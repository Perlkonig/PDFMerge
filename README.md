# PDF Merge

**This software is proof of concept only! As of yet, no helpful error messages, no real documentation, no flexibility.**

An organization I belong to wanted to do a mail merge with PDFs. Acrobat doesn't support that, and they were interested in working with final PDFs instead of doing the mail merge at the production end. This is a simple, standalone .NET WPF app that takes a pre-existing PDF form with form fields and an Excel spreadsheet with header rows matching the names of form fields in the PDF and creates a new PDF for each row in the Excel sheet.

## Getting Started

### Prerequisites

This app targets the .NET 4.0 architecture. Current Windows users shouldn't need to do anything special. If for some reason you're running an older version of .NET, Windows should tell you during the install process and have you upgrade.

### Installing

End users, go to the [Releases page](https://github.com/Perlkonig/PDFMerge/releases), find the version you want, then download the ``PDFMerge-ClickOnce-Package.zip`` file. Unzip that folder then double-click on ``setup.exe``. 

Developers, just fork the repository and go to town.

## Running the tests

In the ``examples`` folder are a sample Excel file and PDF form. Drop those into the appropriate text field, provide an output folder, then click Run. That will at least show you how things work.

## Built With

* [Visual Studio Community 2017](https://www.visualstudio.com/) - Development environment
* [iTextSharp](https://github.com/itext/itextsharp) - PDF library

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

