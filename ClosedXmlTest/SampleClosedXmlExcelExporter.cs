using ClosedXML.Excel;

namespace ClosedXmlTest
{
    public class SampleClosedXmlExcelExporter
    {
        public static Stream GetSampleExcelFile()
        {
            var rows = new List<SampleDto>
            {
                new SampleDto("val11", 1111, "first row"),
                new SampleDto("val11", 2222, "second row")
            };

            var stream = new MemoryStream();
            Write(rows, stream);
            return stream;
        }

        /// <summary>
        /// All the ClosedXml-specific code is in here
        /// </summary>
        private static void Write(IEnumerable<SampleDto> rows, Stream stream)
        {
            // event tracking is used to dynamically resize ranges as data is added, so not required here
            // plus disabling it means we don't need to worry about object disposal (according to the docs)
            // https://github.com/closedxml/closedxml/wiki/Where-to-use-the-using-keyword
            var workbook = new XLWorkbook(XLEventTracking.Disabled);

            // self-explanatory really
            var worksheet = workbook.AddWorksheet("mySheet");

            // Cells can be referenced by A1 notation or by numeric row,col
            // InsertTable allows you to add an enumerable and creates a header row with the field names
            // The second parameter switches off the creation of an Excel table with striping/filtering etc.
            worksheet.Cell("A1").InsertTable(rows, false);

            // again, self-explanatory
            worksheet.Range("A2:C3").AddToNamed("MyRange");

            // SaveAs takes a file path or a stream
            workbook.SaveAs(stream);
        }
    }
}