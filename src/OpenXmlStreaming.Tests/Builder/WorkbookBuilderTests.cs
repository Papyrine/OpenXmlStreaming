using DocumentFormat.OpenXml.Spreadsheet;

public class WorkbookBuilderTests
{
    [Test]
    public async Task RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true))
        {
            workbook.AddWorksheet(
                "Revenue",
                new(
                    new SheetData(
                        new Row(
                            new Cell {
                                CellValue = [with("Q1")],
                                DataType = CellValues.InlineString },
                            new Cell {
                                CellValue = [with("1000")],
                                DataType = CellValues.Number }))));

            workbook.AddWorksheet(
                "Expenses",
                new(
                    new SheetData(
                        new Row(
                            new Cell
                            {
                                CellValue = [with("Rent")], DataType = CellValues.InlineString
                            },
                            new Cell
                            {
                                CellValue = [with("500")],
                                DataType = CellValues.Number
                            }))));
        }

        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var sheets = doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();

        using (Assert.Multiple())
        {
            await Assert.That(sheets).Count().IsEqualTo(2);
            await Assert.That(sheets[0].Name!.Value).IsEqualTo("Revenue");
            await Assert.That(sheets[1].Name!.Value).IsEqualTo("Expenses");
        }

        stream.Position = 0;
        await Verify(stream, extension: "xlsx");
    }

    [Test]
    public async Task AddAfterDispose_Throws()
    {
        using var stream = new MemoryStream();
        var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true);
        workbook.Dispose();

        await Assert.That(() =>
            workbook.AddWorksheet("Late", new(new SheetData()))).ThrowsExactly<InvalidOperationException>();
    }
}
