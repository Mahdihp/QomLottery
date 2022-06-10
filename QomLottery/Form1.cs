using IronXL;
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
                        numericUpDown1.Enabled = true;
                        listBox1.Items.Clear();
                        BtnImport.Enabled = false;
                        dtExcel = await ReadExcel(openFileDialog.FileName); //read excel file
                        label1.Text = " تعداد شرکت کنندگان مجاز " + dtExcel.Rows.Count;
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

        private async Task<DataTable> ReadExcel(string fileName)
        {
            return await Task.Run(() =>
            {
                WorkBook workbook = WorkBook.Load(fileName);
                //// Work with a single WorkSheet.
                ////you can pass static sheet name like Sheet1 to get that sheet
                ////WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
                //You can also use workbook.DefaultWorkSheet to get default in case you want to get first sheet only
                WorkSheet sheet = workbook.DefaultWorkSheet;
                //Convert the worksheet to System.Data.DataTable
                //Boolean parameter sets the first row as column names of your table.
                DataTable dataTable = sheet.ToDataTable(true);
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    //09126516617
                    if (dataTable.Rows[i][0].ToString().Length < 10)
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
            numericUpDown1.Enabled = false;
            listBox1.Items.Clear();
            random = new Random((int)DateTime.Now.Ticks);
            for (var i = 0; i < numericUpDown1.Value; i++)
            {
                bool IsFound = false;
                while (IsFound == false)
                {
                    var r = random.Next(0, dtExcel.Rows.Count);
                    var val = dtExcel.Rows[r][0].ToString();
                    IsFound = IsFoundToList(val);
                    if (IsFound == false)
                    {
                        listBox1.Items.Add(dtExcel.Rows[r][0].ToString());
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
    }
}
