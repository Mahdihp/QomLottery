using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QomLottery
{
    public class UtilLottrey
    {
        string Path1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\LotteryResult.txt";
        string Path2 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\LotterySelectionResult.txt";
        public bool OptLottery { get; set; } = false;
        public Task<DataTable> ReadExcel(string fileName, IProgress<int> progress)
        {
            return Task.Run(() =>
            {
                DataTable dataTable = new DataTable();
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Descendants<Row>();

                    foreach (Cell cell in rows.ElementAt(0))
                    {
                        dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                    }
                    int Index = 1;
                    foreach (Row row in rows)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                        {
                            dataRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        }
                        dataTable.Rows.Add(dataRow);
                        if (progress != null)
                        {
                            progress.Report(Index);
                        }
                        Index++;
                        if (OptLottery == true)
                        {
                            if (dataTable.Rows.Count > 0)
                            {
                                dataTable.Rows.RemoveAt(0);
                            }
                            return dataTable;
                        }
                    }

                }
                dataTable.Rows.RemoveAt(0);
                return dataTable;
            });
        }
        public void CreateFileHistory()
        {
            if (!File.Exists(Path1))
            {
                File.Create(Path1);
            }
            if (!File.Exists(Path2))
            {
                File.Create(Path2);
            }
        }
        public void SaveReslt(string lottery, int type, bool IsNewSave)
        {

            TextWriter txt = null;
            if (type == 1)
            {
                //File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\LotteryResult.txt", lottery + Environment.NewLine);
                string myString = "";
                if (IsNewSave == false)
                {
                    StreamReader sr = new StreamReader(Path1);
                    myString = sr.ReadToEnd();
                    sr.Close();
                }
                StreamWriter sw = new StreamWriter(Path1);
                sw.WriteLine(myString + "\n" + lottery + "\n");
                sw.Close();
            }
            else
            {
                //File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\LotterySelectionResult.txt", lottery + Environment.NewLine);
                string myString = "";
                if (IsNewSave == false)
                {
                    StreamReader sr = new StreamReader(Path2);
                    myString = sr.ReadToEnd();
                    sr.Close();
                }
                StreamWriter sw = new StreamWriter(Path2);
                sw.WriteLine(myString + "\n" + lottery + "\n");
                sw.Close();
            }

        }
        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue?.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
        private async Task<DataTable> ReadExcel(string fileName)
        {
            return await Task.Run(() =>
            {
                IronXL.License.LicenseKey = IronXLKeygen.GenerateKey();
                IronXL.WorkBook workbook = IronXL.WorkBook.Load(fileName);
                IronXL.WorkSheet sheet = workbook.DefaultWorkSheet;
                DataTable dataTable = sheet.ToDataTable(true);
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                //09126516617
                    var mobile = dataTable.Rows[i][0].ToString().Trim();
                    if (mobile.Length < 10 || mobile.StartsWith("00"))
                    {
                        dataTable.Rows.RemoveAt(i);
                    }
                }

                return dataTable;
            });

        }

        //public static Task<FastExcel.Worksheet> ReadExcel2(string fileName)
        //{
        //    return Task.Run(() => {

        //       FastExcel.Worksheet worksheet = null;
        //        var inputFile = new FileInfo(fileName);
        //        // Create an instance of Fast Excel
        //        using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
        //        {
        //            // Read the rows using worksheet name
        //            worksheet = fastExcel.Read("sheet1");

        //            // Read the rows using the worksheet index
        //            // Worksheet indexes are start at 1 not 0
        //            // This method is slightly faster to find the underlying file (so slight you probably wouldn't notice)
        //           return worksheet = fastExcel.Read(1);
        //        }

        //    });
        //}
    }
}
