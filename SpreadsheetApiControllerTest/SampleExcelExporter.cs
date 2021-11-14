using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SampleApiControllerTest
{
    internal class SampleExcelExporter
    {
        public static Stream GetSampleExcelFile()
        {
            var stream = new MemoryStream();
            var writer = new ExcelRowWriter();
            writer.Open(stream);
            writer.WriteRow(1, "column heading one,and two,yet three");
            writer.WriteRow(2, "val1,twotwotwo,33333");
            writer.WriteRow(3, "val1,2222,tri");
            writer.CreateNamedRange("MyRange", "mySheet", "$A$2", "$C$3");
            writer.Close();
            return stream;
        }
    }

    internal class ExcelRowWriter
    {
        private SpreadsheetDocument _spreadsheetDocument;
        private WorkbookPart _workbookPart;

        internal void Open(Stream stream)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            Initialise();
        }

        internal void Open(string filepath)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
            Initialise();
        }

        private void Initialise()
        {
            _workbookPart = _spreadsheetDocument.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();
            var worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheets = _workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            var sheet = new Sheet
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);
            _spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
        }

        internal void WriteRow(uint rowIndex, string rowValues)
        {
            var values = rowValues.Split(',');
            var currentColumn = 1;
            var worksheet = _workbookPart.WorksheetParts.First().Worksheet;
            foreach (var value in values)
            {
                var index = InsertSharedStringItem(value);
                var cell = InsertCell(GetExcelColumnName(currentColumn), rowIndex, worksheet);
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                currentColumn++;
            }
        }

        private int InsertSharedStringItem(string text)
        {
            var sharedStringTablePart = _workbookPart.SharedStringTablePart;
            if (sharedStringTablePart.SharedStringTable == null)
            {
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;
            foreach (var item in sharedStringTablePart.SharedStringTable)
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            sharedStringTablePart.SharedStringTable.Save();

            return i;
        }

        private static Cell InsertCell(string columnName, uint rowIndex, Worksheet worksheet)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = $"{columnName}{rowIndex}";

            Row row;
            if (sheetData.Elements<Row>().Any(r => r.RowIndex == rowIndex))
            {
                row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
            }
            else
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            var newCell = new Cell { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);
            worksheet.Save();
            return newCell;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        internal void CreateNamedRange(string rangeName, string sheetName, string fromCell, string toCell)
        {
            _workbookPart.Workbook.DefinedNames = new DefinedNames();
            _workbookPart.Workbook.DefinedNames.AddChild(new DefinedName
            {
                Name = rangeName,
                Text = $"{sheetName}!{fromCell}:{toCell}"
            });
        }

        internal void Close()
        {
            _spreadsheetDocument.WorkbookPart.Workbook.Save();
            _spreadsheetDocument.Close();
        }
    }
}
