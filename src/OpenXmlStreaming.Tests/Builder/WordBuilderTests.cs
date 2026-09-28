using DocumentFormat.OpenXml.Wordprocessing;

public class WordBuilderTests
{
    [Test]
    public async Task RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true))
        {
            word.AddStyles(
                new(
                    new Style(
                        new StyleName
                        {
                            Val = "Heading 1"
                        },
                        new BasedOn
                        {
                            Val = "Normal"
                        },
                        new NextParagraphStyle
                        {
                            Val = "Normal"
                        },
                        new StyleRunProperties(
                            new Bold(),
                            new FontSize
                            {
                                Val = "32"
                            }))
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = "Heading1"
                    }));

            var footerId = word.AddFooter(
                new(
                    new Paragraph(
                        new Run(new Text("— Confidential —")))));

            word.WriteDocument(
                new(
                    new Body(
                        new Paragraph(
                            new ParagraphProperties(
                                new ParagraphStyleId
                                {
                                    Val = "Heading1"
                                }),
                            new Run(new Text("Quarterly Report"))),
                        new Paragraph(
                            new Run(new Text("Revenue grew 15% year-over-year."))),
                        new SectionProperties(
                            new FooterReference
                            {
                                Type = HeaderFooterValues.Default,
                                Id = footerId
                            }))));
        }

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText).Contains("Quarterly Report");
        await Assert.That(doc.MainDocumentPart.FooterParts.Count()).IsEqualTo(1);
        await Assert.That(doc.MainDocumentPart.StyleDefinitionsPart).IsNotNull();

        stream.Position = 0;
        await Verify(stream, extension: "docx")
            .Snapshot(
                """
                {
                  Text:
                Quarterly Report
                Revenue grew 15% year-over-year.
                }
                """);
    }

    [Test]
    public async Task AddStylesAfterDocument_Throws()
    {
        using var stream = new MemoryStream();
        using var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true);
        word.WriteDocument(new(new Body()));

        await Assert.That(() =>
            word.AddStyles(new())).ThrowsExactly<InvalidOperationException>();
    }

    [Test]
    public async Task DoubleStyles_Throws()
    {
        using var stream = new MemoryStream();
        using var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true);
        word.AddStyles(new());

        await Assert.That(() =>
            word.AddStyles(new())).ThrowsExactly<InvalidOperationException>();
    }
}
