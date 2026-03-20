use anyhow::{Context, Result};
use clap::{ArgAction, Parser};
use serde::Deserialize;
use std::io::Write;
use std::path::PathBuf;
use std::process::{Command, Stdio};

const APPLESCRIPT: &str = r#"
use framework "Foundation"
use framework "AppKit"
use framework "PDFKit"
use scripting additions

on run argv
  if (count of argv) < 3 then
    error "pptxrender requires an input file, output directory, and scale"
  end if

  set pptxPath to item 1 of argv
  set outDir to item 2 of argv
  set scaleFactor to (item 3 of argv) as real

  do shell script "mkdir -p " & quoted form of outDir

  set uuid to do shell script "uuidgen"
  set pdfPath to (outDir & "/.__pptxrender_tmp_" & uuid & ".pdf")

  set inAlias to POSIX file pptxPath as alias
  set outURL to (current application's NSURL's fileURLWithPath:pdfPath)
 
  tell application id "PPT3"
    activate
    open inAlias
    delay 0.4
    save active presentation in (outURL as «class furl») as save as PDF
    close active presentation saving no
  end tell

  set pdfURL to current application's NSURL's fileURLWithPath:pdfPath
  set doc to current application's PDFDocument's alloc()'s initWithURL:pdfURL
  if doc is missing value then error "Failed to load exported PDF"

  set pageCount to (doc's pageCount()) as integer
  if pageCount < 1 then error "Exported PDF has no pages"

  repeat with i from 1 to pageCount
    set page to (doc's pageAtIndex:(i - 1))

    set pageRect to (page's boundsForBox:(current application's kPDFDisplayBoxMediaBox)) as list
    set wPts to (item 1 of item 2 of pageRect) as real
    set hPts to (item 2 of item 2 of pageRect) as real

    set outW to (wPts * scaleFactor) as integer
    set outH to (hPts * scaleFactor) as integer

    set img to (page's thumbnailOfSize:(current application's NSMakeSize(outW, outH)) forBox:(current application's kPDFDisplayBoxMediaBox))
    set tiffData to img's TIFFRepresentation()
    set rep to (current application's NSBitmapImageRep's imageRepWithData:tiffData)
    set pngData to (rep's representationUsingType:(current application's NSPNGFileType) |properties|:(current application's NSDictionary's dictionary()))

    set nStr to i as text
    if (length of nStr) = 1 then set nStr to "000" & nStr
    if (length of nStr) = 2 then set nStr to "00" & nStr
    if (length of nStr) = 3 then set nStr to "0" & nStr

    set outPath to (outDir & "/slide-" & nStr & ".png")
    set outFileURL to current application's NSURL's fileURLWithPath:outPath
    pngData's writeToURL:outFileURL atomically:true
  end repeat

  do shell script "rm -f " & quoted form of pdfPath
end run
"#;

#[derive(Parser, Debug)]
#[command(
    name = "pptxrender",
    about = "Render PPTX to slide PNGs using PowerPoint + PDFKit",
    help_template = "usage: pptxrender --in-path <file.pptx> --out-path <out-path> [--scale 2.0] [--transparent-background] [--dark-mode]\n   or: pptxrender --json '{\"inPath\":\"file.pptx\",\"outPath\":\"out-path\",\"scale\":2,\"transparentBackground\":true,\"darkMode\":false}'\n\n{about-section}\n\n{all-args}\n"
)]
struct CliArgs {
    #[arg(long, value_name = "JSON", help = "Read arguments from a JSON object")]
    json: Option<String>,

    #[arg(long, value_name = "FILE.PPTX", help = "Input PPTX file")]
    in_path: Option<PathBuf>,

    #[arg(
        long = "out-path",
        alias = "out-dir",
        value_name = "DIR",
        help = "Destination directory for rendered slide images"
    )]
    out_path: Option<PathBuf>,

    #[arg(long, value_name = "SCALE", help = "Render scale multiplier")]
    scale: Option<f64>,

    #[arg(
        long,
        action = ArgAction::SetTrue,
        help = "Accept transparentBackground from CLI or JSON"
    )]
    transparent_background: bool,

    #[arg(
        long,
        action = ArgAction::SetTrue,
        help = "Accept darkMode from CLI or JSON"
    )]
    dark_mode: bool,
}

#[derive(Debug, Deserialize, Default)]
struct JsonArgs {
    #[serde(rename = "inPath")]
    in_path: Option<PathBuf>,
    #[serde(rename = "outPath")]
    out_path: Option<PathBuf>,
    #[serde(rename = "outDir")]
    out_dir: Option<PathBuf>,
    scale: Option<f64>,
    #[serde(rename = "transparentBackground")]
    transparent_background: Option<bool>,
    #[serde(rename = "darkMode")]
    dark_mode: Option<bool>,
}

#[derive(Debug)]
struct Args {
    in_path: PathBuf,
    out_path: PathBuf,
    scale: f64,
    transparent_background: bool,
    dark_mode: bool,
}

fn resolve_args(cli: CliArgs) -> Result<Args> {
    let json_args = match cli.json {
        Some(payload) => Some(
            serde_json::from_str::<JsonArgs>(&payload).context("failed to parse --json payload")?,
        ),
        None => None,
    };

    let json = json_args.as_ref();
    let in_path = cli
        .in_path
        .or_else(|| json.and_then(|args| args.in_path.clone()))
        .context("missing input path. Pass --in-path or include inPath in --json")?;

    let out_path = cli
        .out_path
        .or_else(|| json.and_then(|args| args.out_path.clone().or_else(|| args.out_dir.clone())))
        .context(
            "missing output path. Pass --out-path/--out-dir or include outPath/outDir in --json",
        )?;

    let scale = cli
        .scale
        .or_else(|| json.and_then(|args| args.scale))
        .unwrap_or(2.0);

    let transparent_background = cli.transparent_background
        || json
            .and_then(|args| args.transparent_background)
            .unwrap_or(false);

    let dark_mode = cli.dark_mode || json.and_then(|args| args.dark_mode).unwrap_or(false);

    Ok(Args {
        in_path,
        out_path,
        scale,
        transparent_background,
        dark_mode,
    })
}

fn main() -> Result<()> {
    let cli = CliArgs::parse();
    let args = resolve_args(cli)?;

    let in_abs = std::fs::canonicalize(&args.in_path)?;
    std::fs::create_dir_all(&args.out_path)?;
    let out_abs = std::fs::canonicalize(&args.out_path)?;
    let _ = (args.transparent_background, args.dark_mode);

    let mut child = Command::new("osascript")
        .arg("-")
        .arg(in_abs.to_string_lossy().to_string())
        .arg(out_abs.to_string_lossy().to_string())
        .arg(format!("{:.3}", args.scale))
        .stdin(Stdio::piped())
        .stdout(Stdio::inherit())
        .stderr(Stdio::inherit())
        .spawn()?;

    {
        let stdin = child.stdin.as_mut().expect("stdin unavailable");
        stdin.write_all(APPLESCRIPT.as_bytes())?;
    }

    let status = child.wait()?;
    if !status.success() {
        std::process::exit(status.code().unwrap_or(1));
    }

    Ok(())
}
