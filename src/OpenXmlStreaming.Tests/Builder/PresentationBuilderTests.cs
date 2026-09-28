using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

public class PresentationBuilderTests
{
    [Test]
    public async Task RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true))
        {
            presentation.AddSlide(TitleSlide("Kickoff"));
            presentation.AddSlide(TitleSlide("Agenda"));
        }

        stream.Position = 0;
        using var doc = PresentationDocument.Open(stream, false);
        var slideIds = doc.PresentationPart!.Presentation!.SlideIdList!.Elements<SlideId>().ToList();

        using (Assert.Multiple())
        {
            await Assert.That(slideIds).Count().IsEqualTo(2);
            await Assert.That(doc.PresentationPart.SlideParts.Count()).IsEqualTo(2);
            await Assert.That(doc.PresentationPart.SlideMasterParts.Count()).IsEqualTo(1);
        }

        stream.Position = 0;
        await Verify(stream, extension: "pptx")
            .Snapshot(
                """
                {
                  SlideCount: 2,
                  Text:
                Kickoff
                ---
                Agenda
                }
                """);
    }

    [Test]
    public async Task NoSlides_StillWritesScaffolding()
    {
        using var stream = new MemoryStream();

        await using (var _ = new StreamingPresentationBuilder(stream, leaveOpen: true))
        {
            // Intentionally no slides added.
        }

        stream.Position = 0;
        using var doc = PresentationDocument.Open(stream, false);
        await Assert.That(doc.PresentationPart!.Presentation).IsNotNull();
    }

    [Test]
    public async Task AddAfterDispose_Throws()
    {
        using var stream = new MemoryStream();
        var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true);
        presentation.Dispose();

        await Assert.That(() =>
            presentation.AddSlide(TitleSlide("Late"))).ThrowsExactly<InvalidOperationException>();
    }

    static Slide TitleSlide(string title) =>
    [
        with(new CommonSlideData(
            new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties
                    {
                        Id = 1,
                        Name = ""
                    },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new Drawing.TransformGroup()),
                new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = 2,
                            Name = "Title"
                        },
                        new NonVisualShapeDrawingProperties(
                            new Drawing.ShapeLocks
                            {
                                NoGrouping = true
                            }),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape
                            {
                                Type = PlaceholderValues.CenteredTitle
                            })),
                    new ShapeProperties(),
                    new TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(
                            new Drawing.Run(
                                new Drawing.RunProperties
                                {
                                    Language = "en-US"
                                },
                                new Drawing.Text(title))))))))
    ];
}
