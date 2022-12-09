using AODL.Document.Content.Tables;
using AODL.Document.Content.Text;
using AODL.Document.SpreadsheetDocuments;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TrackerCommon
{
    /// <summary>
    /// Simple first step class to explore how to use the AODL library to read and write spreadsheets.
    /// </summary>
    public class ODSTools
    {
        public void DuplicateSpreadsheet(string path)
        {
            var newSheet = new SpreadsheetDocument();
            newSheet.New();
            var newTable = new Table(newSheet, "Sheet1", "ta1");
            var newColumn = new Column(newTable, "co1");
            newTable.ColumnCollection.Add(newColumn);
            var newRow = new Row(newTable);
            newTable.RowCollection.Add(newRow);
            var newCell = newTable.CreateCell();
            var newContent = new Paragraph(newSheet);
            var newSimpleText = new SimpleText(newSheet, "Test");
            newContent.TextContent.Add(newSimpleText);
            newCell.Content.Add(newContent);
            newTable.InsertCellAt(0, 0, newCell);
            newSheet.TableCollection.Add(newTable);
            var newPath = @"C:\r\OurTracker\testSheet.ods";
            newSheet.SaveTo(newPath);

            var sheetDocument = new SpreadsheetDocument();
            sheetDocument.Load(path);
            for (var i=0; i<sheetDocument.TableCount; i++)
            {
                Console.WriteLine("=======================================================");
                var table = sheetDocument.TableCollection[i];
                Console.WriteLine($"~~~~[{table.TableName}]~~~~");
                for (var j=0; j<table.RowCollection.Count; j++)
                {
                    var row = table.RowCollection[j];
                    Console.WriteLine("-------------------------------------------------------");
                    for (var k=0; k<row.CellCollection.Count; k++)
                    {
                        var cell = row.CellCollection[k];
                        var styles = cell.Style?.PropertyCollection;
                        var formula = cell.Formula;
                        if (!string.IsNullOrEmpty(formula))
                        {
                            formula = $" ({formula})";
                        }
                        var content = cell.Content.Count > 0 ? cell.Content[0] as Paragraph : null;
                        if (content is Paragraph && content.TextContent.Count > 0)
                        {
                            Console.Write(content.TextContent[0].Text + formula + " | ");
                        }
                    }
                    Console.WriteLine();
                }
            }
            //TODO: the json serialization fails.
            //It's not imperative that I get it working though.
            //Really just curious to see the structure of the document
            //all in one file. :P
            //var options = new JsonSerializerOptions
            //{
            //    WriteIndented = true,
            //    IgnoreReadOnlyProperties = true,
            //    IgnoreReadOnlyFields = true,
            //    ReferenceHandler = ReferenceHandler.Preserve,
            //    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            //    UnknownTypeHandling = JsonUnknownTypeHandling.JsonNode,
            //    IncludeFields = false,
            //};
            //var json = JsonSerializer.Serialize(sheetDocument, options);
            //var writePath = @"C:\r\OurTracker\outputTest.txt";
            //File.WriteAllText(writePath, json);
        }
    }
}