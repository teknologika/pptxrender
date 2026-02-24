use clap::Parser;
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
    error "usage: pptxrender <pptxPath> <outDir> <scale>"
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
    about = "Render PPTX to slide PNGs using PowerPoint + PDFKit"
)]
struct Args {
    #[arg(long)]
    in_path: PathBuf,

    #[arg(long)]
    out_dir: PathBuf,

    #[arg(long, default_value_t = 2.0)]
    scale: f64,
}

fn main() -> anyhow::Result<()> {
    let args = Args::parse();

    let in_abs = std::fs::canonicalize(&args.in_path)?;
    std::fs::create_dir_all(&args.out_dir)?;
    let out_abs = std::fs::canonicalize(&args.out_dir)?;

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
