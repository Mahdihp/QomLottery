using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QomLottery
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }
        DataTable dtExcel = new DataTable();
        DataTable dtExcelSearch = new DataTable();
        public static Random random;
        private async void BtnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(openFileDialog.FileName); //get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        label1.Text = "...لطفا صبر کنید";
                        listBox1.Items.Clear();
                        BtnImport.Enabled = false;
                        var progress = new System.Progress<int>(update =>
                        {
                            label1.Text = update + ":تعداد ردیف های خوانده شده";
                        });
                        dtExcel = await ReadExcel(openFileDialog.FileName, progress); //read excel file
                        label1.Text = " تعداد شرکت کنندگان مجاز " + dtExcel.Rows.Count;
                        numericUpDown1.Maximum = dtExcel.Rows.Count - 1;
                        BtnLottery.Enabled = true;
                        BtnImport.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                }
            }
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
                    }

                }

                dataTable.Rows.RemoveAt(0);
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
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
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

        private void BtnLottery_Click(object sender, EventArgs e)
        {
            if (numericUpDown1.Value > dtExcel.Rows.Count)
            {
                MessageBox.Show("تعداد شرکت کنندگان کمتر از تعداد قرعه کشی میباشد.");
                return;
            }
            // numericUpDown1.Enabled = false;
            listBox1.Items.Clear();
            random = new Random((int)DateTime.Now.Ticks);
            for (var i = 0; i < numericUpDown1.Value; i++)
            {
                bool IsFound = false;
                while (IsFound == false)
                {
                    var r = random.Next(0, dtExcel.Rows.Count);
                    var val = dtExcel.Rows[r][4].ToString();
                    IsFound = IsFoundToList(val);
                    if (IsFound == false)
                    {
                        listBox1.Items.Add(dtExcel.Rows[r][4].ToString());
                        dtExcel.Rows.RemoveAt(r);
                        label1.Text = " تعداد شرکت کنندگان مجاز " + dtExcel.Rows.Count;
                        IsFound = true;
                    }
                    else
                    {
                        IsFound = false;
                    }
                }

            }
        }

        private bool IsFoundToList(string val)
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {

                if (listBox1.Items[i].ToString() == val)
                {
                    return true;
                }
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var Result = string.Empty;
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                Result += i + 1 + " - " + listBox1.Items[i].ToString() + Environment.NewLine;
            }
            Clipboard.SetText(Result);
            MessageBox.Show("به حافظه منتقل شد.");
        }

        private async void BtnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(openFileDialog.FileName); //get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        label4.Text = "...لطفا صبر کنید";
                        listBox2.Items.Clear();
                        BtnImportExcel.Enabled = false;
                        var progress = new System.Progress<int>(update =>
                        {
                            label4.Text = update + ":تعداد ردیف های خوانده شده";
                        });
                        dtExcelSearch = await ReadExcel(openFileDialog.FileName, progress); //read excel file
                        label4.Text = " تعداد شرکت کنندگان مجاز " + dtExcelSearch.Rows.Count;
                        numericUpDown2.Maximum = dtExcelSearch.Rows.Count;
                        BtnSearch.Enabled = true;
                        numericUpDown2.Enabled = true;
                        BtnImportExcel.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                }
            }
        }
        private bool IsFoundLottery(string val)
        {
            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                if (listBox2.Items[i].ToString() == val)
                {
                    return true;
                }
            }
            return false;
        }
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (numericUpDown2.Value > dtExcelSearch.Rows.Count)
            {
                MessageBox.Show("این شماره موجود نیست.");
                return;
            }
            if (numericUpDown2.Value > 0)
            {
                int index = Convert.ToInt32(numericUpDown2.Value);
                var lottery = index + " :شماره برنده خوش شانش " + "\n" + dtExcelSearch.Rows[index - 1][4].ToString() + " :شماره فیش " + "\n" +
                    dtExcelSearch.Rows[index - 1][1].ToString() + " :کد نوسازی " + "\n";

                if (!IsFoundLottery(lottery))
                {
                    listBox2.Items.Add(lottery);
                    ShowLottery showLottery = new ShowLottery(lottery);
                    showLottery.ShowDialog();
                }
                else
                {
                    MessageBox.Show("این مورد قبلا انتخاب شده است.");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count <= 0)
            {
                return;
            }
            var Result = string.Empty;
            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                Result += i + 1 + " - " + listBox2.Items[i].ToString() + Environment.NewLine;
            }
            Clipboard.SetText(Result);
            MessageBox.Show("به حافظه منتقل شد.");
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown2.Value > dtExcelSearch.Rows.Count)
            {
                numericUpDown2.Value = 0;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
    }
}
