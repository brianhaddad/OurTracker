using AODL.Document.Content.Text;
using AODL.Document.SpreadsheetDocuments;

namespace TrackerCommon
{
    /// <summary>
    /// Simple first step class to explore how to use the AODL library to read and write spreadsheets.
    /// </summary>
    public class ODSTools
    {
        public void ReadOutSpreadsheet(string path)
        {
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
                    for (var k=0; k<row.CellCollection.Count; k++)
                    {
                        var cell = row.CellCollection[k];
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
        }
    }
}