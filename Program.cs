// Topic: Merge, split and trim PDF, Word and PowerPoint files with one line of C# each.
// Uses GroupDocs.Merger.LowCode for .NET 26.10: JoinPdf, JoinDocx, SplitPdf, SplitPptx
// and RemovePagesPdf.

using GroupDocs.Merger.Domain.Options;
using GroupDocs.Merger.Exceptions;
using GroupDocs.Merger.LowCode;

namespace Demo.OneLineMerge;

internal static class Program
{
    private const string InputFolder = "resources";
    private const string ResultFolder = "output";

    private static readonly string Report = Path.Combine(InputFolder, "report.pdf");
    private static readonly string Appendix = Path.Combine(InputFolder, "appendix.pdf");
    private static readonly string Contract = Path.Combine(InputFolder, "contract.docx");
    private static readonly string Amendment = Path.Combine(InputFolder, "amendment.docx");
    private static readonly string Deck = Path.Combine(InputFolder, "deck.pptx");

    private static int Main()
    {
        Directory.CreateDirectory(ResultFolder);
        ApplyLicense();

        foreach (string input in new[] { Report, Appendix, Contract, Amendment, Deck })
        {
            if (!File.Exists(input))
            {
                Console.Error.WriteLine($"Missing input: {Path.GetFullPath(input)}");
                return 1;
            }
        }

        Console.WriteLine("1. Two PDF files joined:");
        long joinedPdf = JoinPdfFiles(Report, Appendix, ResultPath("report-with-appendix.pdf"));
        Console.WriteLine($"   bytes: {joinedPdf}");

        Console.WriteLine("2. Two Word documents joined:");
        long joinedDocx = JoinWordDocuments(
            Contract, Amendment, ResultPath("contract-with-amendment.docx"));
        Console.WriteLine($"   bytes: {joinedDocx}");

        Console.WriteLine("3. Pages 1 and 2 of a PDF split into files of their own:");
        int pdfParts = SplitPdfPages(Report, new[] { 1, 2 }, ResultPath("report-page_{0}.pdf"));
        Console.WriteLine($"   files: {pdfParts}");

        Console.WriteLine("4. Slides 1 and 2 of a presentation split into decks of their own:");
        int deckParts = SplitPresentationSlides(Deck, new[] { 1, 2 }, ResultPath("deck-slide.pptx"));
        Console.WriteLine($"   files: {deckParts}");

        Console.WriteLine("5. Page 2 removed from a PDF:");
        long trimmed = RemovePdfPages(Report, new[] { 2 }, ResultPath("report-without-page-2.pdf"));
        Console.WriteLine($"   bytes: {trimmed}");

        Console.WriteLine("6. A Word document passed to the PDF join:");
        string? rejection = RejectMixedFormatJoin(Report, Contract, ResultPath("mixed.pdf"));
        Console.WriteLine(rejection == null ? "   accepted" : "   rejected");

        Console.WriteLine($"Output folder: {Path.GetFullPath(ResultFolder)}");
        bool asExpected = joinedPdf > 0 && joinedDocx > 0 && pdfParts == 2 && deckParts == 2
            && trimmed > 0 && rejection != null;
        return asExpected ? 0 : 2;
    }

    private static string ResultPath(string fileName) =>
        Path.Combine(ResultFolder, fileName);

    private static void ApplyLicense()
    {
        // Point this at your .lic file to remove evaluation limits.
        // Get a free temporary licence: https://purchase.groupdocs.com/temporary-license
        const string licensePath = "REPLACE_WITH_YOUR_LICENSE_PATH";
        if (File.Exists(licensePath))
        {
            License.Set(licensePath);
            Console.WriteLine("[license] applied");
        }
        else
        {
            Console.WriteLine("[license] no licence set - running in evaluation mode");
        }
    }

    /// <summary>
    /// Appends one PDF file to another with a single JoinPdf call.
    /// </summary>
    /// <remarks>
    /// JoinPdf takes the source file, the files to append and the join options, and runs
    /// when Save is called. Every file must be a PDF; a relative output path is resolved
    /// against the current working directory. Returns the size of the joined file.
    /// </remarks>
    private static long JoinPdfFiles(string sourcePath, string appendedPath, string outputPath)
    {
        new JoinPdf(sourcePath, new[] { appendedPath }, new JoinOptions()).Save(outputPath);

        return new FileInfo(outputPath).Length;
    }

    /// <summary>
    /// Appends one Word document to another with a single JoinDocx call.
    /// </summary>
    /// <remarks>
    /// JoinDocx works like JoinPdf for DOCX files: the appended document starts after the
    /// last page of the source document. Returns the size of the joined file.
    /// </remarks>
    private static long JoinWordDocuments(string sourcePath, string appendedPath,
        string outputPath)
    {
        new JoinDocx(sourcePath, new[] { appendedPath }, new JoinOptions()).Save(outputPath);

        return new FileInfo(outputPath).Length;
    }

    /// <summary>
    /// Splits the chosen pages of a PDF file into one file per page with SplitPdf.
    /// </summary>
    /// <remarks>
    /// The {0} placeholder in the output file name is replaced with the part number, so
    /// report-page_{0}.pdf yields report-page_1.pdf, report-page_2.pdf and so on. Returns
    /// the number of files written.
    /// </remarks>
    private static int SplitPdfPages(string sourcePath, int[] pageNumbers,
        string outputPattern)
    {
        new SplitPdf(sourcePath, pageNumbers).Save(outputPattern);

        return Directory.GetFiles(ResultFolder, "report-page_*.pdf").Length;
    }

    /// <summary>
    /// Splits the chosen slides of a presentation into one deck per slide with SplitPptx.
    /// </summary>
    /// <remarks>
    /// Without a {0} placeholder the part number is appended to the file name, so
    /// deck-slide.pptx yields deck-slide_1.pptx, deck-slide_2.pptx and so on in the same
    /// folder. Returns the number of files written.
    /// </remarks>
    private static int SplitPresentationSlides(string sourcePath, int[] slideNumbers,
        string outputPath)
    {
        new SplitPptx(sourcePath, slideNumbers).Save(outputPath);

        return Directory.GetFiles(ResultFolder, "deck-slide_*.pptx").Length;
    }

    /// <summary>
    /// Removes the chosen pages from a PDF file with RemovePagesPdf.
    /// </summary>
    /// <remarks>
    /// RemovePagesPdf takes the page numbers through RemoveOptions and writes a copy of the
    /// document without them; the source file is left unchanged. Returns the size of the
    /// trimmed file.
    /// </remarks>
    private static long RemovePdfPages(string sourcePath, int[] pageNumbers, string outputPath)
    {
        new RemovePagesPdf(sourcePath, new RemoveOptions(pageNumbers)).Save(outputPath);

        return new FileInfo(outputPath).Length;
    }

    /// <summary>
    /// Shows that a PDF join rejects a file of another format before writing anything.
    /// </summary>
    /// <remarks>
    /// Each LowCode product handles one format. JoinPdf checks every file it is given and
    /// throws FileTypeNotSupportedException, naming the file and its format, when one of
    /// them is not a PDF. Returns that message, or null when the join is accepted.
    /// </remarks>
    private static string? RejectMixedFormatJoin(string pdfPath, string wordPath,
        string outputPath)
    {
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
    }
}
