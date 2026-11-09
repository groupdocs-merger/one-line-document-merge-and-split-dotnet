# Merge and Split PDF, Word and PowerPoint Files in One Line of C#

<div align="center">

[![Product Page](https://img.shields.io/badge/Product%20Page-2865E0?style=for-the-badge&logo=appveyor&logoColor=white)](https://products.groupdocs.com/merger/net)
[![Docs](https://img.shields.io/badge/Docs-2865E0?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.groupdocs.com/merger/net/)
[![Blog](https://img.shields.io/badge/Blog-2865E0?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.groupdocs.com/categories/groupdocs.merger-product-family/)
[![Free Support](https://img.shields.io/badge/Free%20Support-2865E0?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.groupdocs.com/c/merger/)
[![Temporary License](https://img.shields.io/badge/Temporary%20License-2865E0?style=for-the-badge&logo=rocket&logoColor=white)](https://purchase.groupdocs.com/temporary-license)

</div>

## Quick Start

`one-line-document-merge-and-split-dotnet` is a runnable .NET console sample that joins
PDF and Word files, splits PDF pages and PowerPoint slides, and removes PDF pages, each
with a single line of C# against a GroupDocs.Merger.LowCode product. The project targets
.NET 8.0 and references `GroupDocs.Merger.LowCode` 26.10.0 directly from NuGet. It is
built for developers who need one specific document operation and do not want to learn
the general-purpose `Merger` class to get it.

## What You'll Learn

- How to join PDF files with `JoinPdf` and Word documents with `JoinDocx`
- How to split chosen PDF pages into separate files with `SplitPdf`
- How to split chosen PowerPoint slides into separate decks with `SplitPptx`
- How to remove specific pages from a PDF with `RemovePagesPdf` without touching the source
- How a LowCode product rejects a file of the wrong format with `FileTypeNotSupportedException`

## Table of Contents

- [About This Repository](#about-this-repository)
- [Key Features](#key-features)
- [Prerequisites](#prerequisites)
- [Repository Structure](#repository-structure)
- [Implementation Examples](#implementation-examples)
- [Related Resources](#related-resources)
- [Keywords](#keywords)

## About This Repository

This repository contains a single runnable .NET console project built on
GroupDocs.Merger.LowCode, demonstrating six document operations with one line of C# each.
GroupDocs.Merger.LowCode packages `JoinPdf`, `JoinDocx`, `JoinPptx`, `SplitPdf`,
`SplitDocx`, `SplitPptx` and `RemovePagesPdf` as separate, single-purpose products instead
of the general `Merger` class. The examples run against a 2-page `report.pdf`, a 2-page
`appendix.pdf`, a 6-page `contract.docx`, a 1-page `amendment.docx` and a 6-slide
`deck.pptx`, all included under `resources/`. They are written for .NET developers who need
to merge or split PDF, Word and PowerPoint files without learning the full `Merger` API
surface first.

## When should I use GroupDocs.Merger.LowCode instead of the full Merger API?

GroupDocs.Merger.LowCode fits a script or pipeline where one operation - joining two PDFs,
splitting a deck, removing a page - is the whole job, and the full `Merger` class would add
setup the task does not need. The LowCode products hold no `Merger` API surface beyond that
one operation: no cross-format joins, no page-level join options, no output streams, no
document info. Reach for the full [GroupDocs.Merger API](https://reference.groupdocs.com/merger/net/)
when a task needs any of those.

## Key Features

### GroupDocs.Merger.LowCode Capabilities

GroupDocs.Merger.LowCode provides focused, single-purpose document processing:

| Feature | Description |
|---------|-------------|
| **JoinPdf / JoinDocx / JoinPptx** | Append one or more documents of the same format to a source document in a single constructor-plus-`Save` call. |
| **SplitPdf / SplitDocx / SplitPptx** | Split chosen pages or slides into one output file per part, named by a `{0}` placeholder or an appended part number. |
| **RemovePagesPdf** | Write a copy of a PDF with the chosen pages removed; the source file is left unchanged. |
| **Format checking** | Every product checks its source and every joined file, throwing `FileTypeNotSupportedException` and naming the offending file when a format does not match. |
| **`License.Set`** | A single static call applies a metered license file or a standard `.lic` file for the LowCode products. |

### What This Repository Demonstrates

- Joining two PDF files and two Word documents into single output files
- Splitting two PDF pages and two PowerPoint slides into files of their own
- Removing a page from a PDF while leaving the source file untouched
- Rejecting a Word file passed to a PDF-only join with the exact exception message raised

## Prerequisites

Before you begin, ensure you have:

- **NuGet package** – `GroupDocs.Merger.LowCode` (targets .NET 6.0 and later; this sample
  targets net8.0), a separate package from the main `GroupDocs.Merger` library.
- **License (optional)** – without one, the LowCode products run in evaluation mode.
- **Sample inputs** – the five files under `resources/` that ship with this repository.

## Repository Structure

```
one-line-document-merge-and-split-dotnet/
│
├── OneLineMergeDemo.csproj
├── Program.cs
└── resources/
    ├── amendment.docx
    ├── appendix.pdf
    ├── contract.docx
    ├── deck.pptx
    └── report.pdf
```

### File Overview

- **OneLineMergeDemo.csproj** – project file targeting net8.0, referencing
  `GroupDocs.Merger.LowCode` 26.10.0
- **Program.cs** – the demo: license check, six LowCode calls, console output
- **resources/report.pdf** – 2-page PDF used by the join, split and remove examples
- **resources/appendix.pdf** – 2-page PDF appended to `report.pdf`
- **resources/contract.docx** – 6-page Word document used by the join example
- **resources/amendment.docx** – 1-page Word document appended to `contract.docx`
- **resources/deck.pptx** – 6-slide PowerPoint deck used by the split example

## Implementation Examples

### Example 1: Appends one PDF file to another with a single JoinPdf call.

`JoinPdf` takes the source PDF, the files to append, and a `JoinOptions` instance in its
constructor, and performs the join only when `Save` is called with the output path. Every
file passed in - the source and each appended file - must be a PDF, and a relative output
path is resolved against the current working directory. The method returns the size in
bytes of the joined file.

```csharp
new JoinPdf(sourcePath, new[] { appendedPath }, new JoinOptions()).Save(outputPath);

return new FileInfo(outputPath).Length;
```

**What This Example Shows:**

In the sample run, joining `report.pdf` (2 pages) with `appendix.pdf` (2 pages) produced
`report-with-appendix.pdf` at 4 pages and 257987 bytes. The call is the full LowCode
pattern: construct the product with its inputs and options, then call `Save` to run it.

[Program.cs#L94-L99](https://github.com/groupdocs-merger/one-line-document-merge-and-split-dotnet/blob/main/Program.cs#L94-L99)

---

### Example 2: Appends one Word document to another with a single JoinDocx call.

`JoinDocx` follows the same constructor-plus-`Save` pattern as `JoinPdf`, but for DOCX
files: the source document, the documents to append, and a `JoinOptions` instance. The
appended document is placed after the last page of the source document, and every input
must be a Word document. It returns the size of the joined file in bytes.

```csharp
new JoinDocx(sourcePath, new[] { appendedPath }, new JoinOptions()).Save(outputPath);

return new FileInfo(outputPath).Length;
```

**What This Example Shows:**

Joining the 6-page `contract.docx` with the 1-page `amendment.docx` in the sample run
produced a 7-page `contract-with-amendment.docx` of 7915 bytes.

[Program.cs#L108-L114](https://github.com/groupdocs-merger/one-line-document-merge-and-split-dotnet/blob/main/Program.cs#L108-L114)

---

### Example 3: Splits the chosen pages of a PDF file into one file per page with SplitPdf.

`SplitPdf` takes the source PDF and the page numbers to extract, and `Save` writes one file
per requested page. A `{0}` placeholder in the output file name is replaced with the part
number, so `report-page_{0}.pdf` yields `report-page_1.pdf` and `report-page_2.pdf`;
without the placeholder the part number is appended instead. The method returns the count
of files actually written.

```csharp
new SplitPdf(sourcePath, pageNumbers).Save(outputPattern);

return Directory.GetFiles(ResultFolder, "report-page_*.pdf").Length;
```

**What This Example Shows:**

Splitting pages 1 and 2 of `report.pdf` produced `report-page_1.pdf` and
`report-page_2.pdf`, each a single page, in the sample run.

[Program.cs#L124-L130](https://github.com/groupdocs-merger/one-line-document-merge-and-split-dotnet/blob/main/Program.cs#L124-L130)

---

### Example 4: Splits the chosen slides of a presentation into one deck per slide with SplitPptx.

`SplitPptx` works like `SplitPdf` but for PowerPoint slides: it takes the source deck and
the slide numbers to extract. Without a `{0}` placeholder in the output file name, the
part number is appended instead, so `deck-slide.pptx` becomes `deck-slide_1.pptx` and
`deck-slide_2.pptx` in the same folder. It returns the count of deck files written.

```csharp
new SplitPptx(sourcePath, slideNumbers).Save(outputPath);

return Directory.GetFiles(ResultFolder, "deck-slide_*.pptx").Length;
```

**What This Example Shows:**

Splitting slides 1 and 2 of the 6-slide `deck.pptx` produced `deck-slide_1.pptx` and
`deck-slide_2.pptx`, one slide each, in the sample run.

[Program.cs#L140-L146](https://github.com/groupdocs-merger/one-line-document-merge-and-split-dotnet/blob/main/Program.cs#L140-L146)

---

### Example 5: Removes the chosen pages from a PDF file with RemovePagesPdf.

`RemovePagesPdf` takes the source PDF and a `RemoveOptions` instance built from the page
numbers to drop, and `Save` writes a new file without them. The source file itself is left
unchanged, so the original document stays available for other operations. The method
returns the size in bytes of the trimmed output file.

```csharp
new RemovePagesPdf(sourcePath, new RemoveOptions(pageNumbers)).Save(outputPath);

return new FileInfo(outputPath).Length;
```

**What This Example Shows:**

Removing page 2 from `report.pdf` produced `report-without-page-2.pdf` at 1 page and
86647 bytes, with `report.pdf` itself unchanged.

[Program.cs#L156-L161](https://github.com/groupdocs-merger/one-line-document-merge-and-split-dotnet/blob/main/Program.cs#L156-L161)

---

### Example 6: Shows that a PDF join rejects a file of another format before writing anything.

Each LowCode product is built for one format, and `JoinPdf` checks the source and every
joined file before writing anything. When one of them is not a PDF, it throws
`FileTypeNotSupportedException` naming the offending file and the format it actually is.
This method catches that exception and returns its message instead of letting the join
produce any output.

```csharp
try
{
    new JoinPdf(pdfPath, new[] { wordPath }, new JoinOptions()).Save(outputPath);
    return null;
}
catch (FileTypeNotSupportedException exception)
{
    Console.WriteLine($"   {exception.Message}");
    return exception.Message;
}
```

**What This Example Shows:**

When we ran the sample, the PDF join stopped at the Word file and named it: passing
`contract.docx` into a PDF join raised `Input format does not match the expected format
'PDF'; the joined file 'resources\contract.docx' is 'Microsoft Word Open XML Document
(.docx)'.`

[Program.cs#L171-L184](https://github.com/groupdocs-merger/one-line-document-merge-and-split-dotnet/blob/main/Program.cs#L171-L184)

## Related Resources

Explore these additional resources to deepen your understanding of one-line document
merging and splitting with GroupDocs.Merger.LowCode:

* **Step-by-step use case guide in the documentation** – a docs.groupdocs.com walkthrough
  of this exact sample, from installation to the rejected mixed-format join: [Read the
  article →](https://docs.groupdocs.com/merger/net/getting-started/use-cases/merge-and-split-documents-in-one-line/)

* **In-depth blog article about this project** – the companion tutorial on the GroupDocs
  blog, covering the same six LowCode operations with their own code walkthrough: [Read the
  article →](https://blog.groupdocs.com/merger/merge-and-split-documents-in-one-line-net/)

* **GroupDocs.Merger.LowCode installation guide** – how to add the `GroupDocs.Merger.LowCode`
  package alongside the main library: [Read the article →](https://docs.groupdocs.com/merger/net/installation/)

* **GroupDocs.Merger for .NET features overview** – the full set of document operations the
  product family supports beyond the LowCode one-liners: [Read the article →](https://docs.groupdocs.com/merger/net/features-overview/)

* **GroupDocs.Merger.LowCode on NuGet** – the package reference used by this sample's
  `.csproj` file: [Read the article →](https://www.nuget.org/packages/GroupDocs.Merger.LowCode/)

## Keywords

`GroupDocs.Merger.LowCode`, `GroupDocs.Merger`, `.NET`, `C#`, `merge PDF in one line`,
`merge Word documents`, `split PDF pages`, `split PowerPoint slides`, `remove PDF pages`,
`JoinPdf`, `SplitPdf`, `low-code document merging`

---

<div align="center">

**Need help?** [Get Free Support](https://forum.groupdocs.com/c/merger/) | [Get Temporary License](https://purchase.groupdocs.com/temporary-license)

</div>
