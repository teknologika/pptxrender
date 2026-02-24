# pptxrender

A macOS command-line tool that renders PowerPoint (`.pptx`) files into
slide images using **Microsoft PowerPoint's native rendering engine**.

`pptxrender` automates PowerPoint to export slides as PDF, then uses
macOS PDFKit to rasterise each slide into PNG images.

The result is deterministic, high-fidelity slide rendering suitable for
automation, testing, and LLM visual QA workflows.

## Why this exists

Most open tools attempt to interpret PPTX files.

This tool does not.

Instead it:

1.  Uses **PowerPoint itself** to render slides (authoritative layout)
2.  Uses **macOS Quartz/PDFKit** to generate images
3.  Produces stable image artefacts for downstream processing

This guarantees output that matches what users actually see in
PowerPoint.

## Requirements

macOS, Microsoft PowerPoint installed, and an interactive user session
(PowerPoint must be able to launch). No additional dependencies are
required.

## Installation

### Build from source

``` bash
git clone https://github.com/<your-org>/pptxrender.git
cd pptxrender
cargo build --release
```

Binary will be located at `target/release/pptxrender`.

## Usage

``` bash
pptxrender   --in-path deck.pptx   --out-dir renders   --scale 2.0
```

Output:

    renders/
      slide-0001.png
      slide-0002.png
      slide-0003.png
      ...

### Arguments

  Flag          Description
  ------------- -----------------------------------------
  `--in-path`   Input PPTX file
  `--out-dir`   Destination directory for images
  `--scale`     Render scale multiplier (default `2.0`)

Higher scale produces higher resolution images.

## First Run (macOS Permission)

On first execution macOS will prompt that the calling app (Terminal or
`pptxrender`) wants to control Microsoft PowerPoint. Click **Allow**.

This permission allows automation of PowerPoint via AppleScript.

## How it works

`pptxrender` drives PowerPoint to export a PDF, then rasterises PDF
pages to PNG using PDFKit. PowerPoint performs layout; macOS performs
rasterisation.

## Limitations

macOS only, requires installed PowerPoint, not supported in headless
environments without a logged-in user, and PowerPoint briefly launches
during rendering.

## Intended use cases

LLM slide QA pipelines, visual regression testing, documentation
automation, slide thumbnail generation, offline rendering workflows.

## Security model

This tool does not parse PPTX files itself. All rendering is delegated
to Microsoft PowerPoint. No network access is performed.

## License

MIT License. See `LICENSE`.

## Contributing

Issues and pull requests are welcome. Please open a discussion before
large architectural changes.

## Acknowledgements

Microsoft PowerPoint rendering engine and macOS Quartz/PDFKit
frameworks.
