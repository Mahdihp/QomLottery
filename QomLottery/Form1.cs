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
        //32861
        public UtilLottrey _OptLottery { get; set; }
        public UtilLottrey _OptSelection { get; set; }
        DataTable dtExcelLottery = new DataTable();
       // DataTable dtExcelLottery_Removed = new DataTable();
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
                            label1.Text = update + ":فیش های خوانده شده";
                        });
                        _OptLottery = new UtilLottrey();
                        _OptLottery.CreateFileHistory();
                        dtExcelLottery = await _OptLottery.ReadExcel(openFileDialog.FileName, progress);
                        //dtExcelLottery_Removed = dtExcelLottery;
                        label1.Text = " تعداد شرکت کنندگان مجاز " + dtExcelLottery.Rows.Count;
                        numericUpDown1.Maximum = dtExcelLottery.Rows.Count - 1;
                        BtnLottery.Enabled = true;
                        BtnImport.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                        numericUpDown1.Enabled = true;
                        BtnLottery.Enabled = true;
                        BtnImport.Enabled = true;
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    numericUpDown1.Enabled = true;
                    BtnLottery.Enabled = true;
                    BtnImport.Enabled = true;
                }
            }
            else
            {
                numericUpDown1.Enabled = true;
                BtnLottery.Enabled = true;
                BtnImport.Enabled = true;
            }
        }




        private void BtnLottery_Click(object sender, EventArgs e)
        {
            if (numericUpDown1.Value > dtExcelLottery.Rows.Count)
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
                    var r = random.Next(0, dtExcelLottery.Rows.Count);
                    var val = dtExcelLottery.Rows[r][3].ToString();
                    IsFound = IsFoundToList(val);
                    if (IsFound == false)
                    {
                        if (_OptLottery != null)
                        {
                            _OptLottery.SaveReslt(dtExcelLottery.Rows[r][3].ToString(), 1, false);
                        }
                        else
                        {
                            _OptLottery = new UtilLottrey();
                            _OptLottery.SaveReslt(dtExcelLottery.Rows[r][3].ToString(), 1, false);
                        }
                        listBox1.Items.Add(dtExcelLottery.Rows[r][3].ToString());
                      //  var row = dtExcelLottery.Rows[r];
                       // dtExcelLottery_Removed.Rows.Add(row);
                        dtExcelLottery.Rows.RemoveAt(r);
                        label1.Text = " :تعداد شرکت کنندگان مجاز " + dtExcelLottery.Rows.Count;
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
                            label4.Text = update + ":فیش های خوانده شده";
                        });
                        _OptSelection = new UtilLottrey();
                        _OptSelection.CreateFileHistory();
                        dtExcelSearch = await _OptSelection.ReadExcel(openFileDialog.FileName, progress); //read excel file
                        label4.Text = " تعداد شرکت کنندگان مجاز " + dtExcelSearch.Rows.Count;
                        numericUpDown2.Maximum = dtExcelSearch.Rows.Count;
                        BtnSearch.Enabled = true;
                        numericUpDown2.Enabled = true;
                        BtnImportExcel.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                        BtnSearch.Enabled = true;
                        numericUpDown2.Enabled = true;
                        BtnImportExcel.Enabled = true;
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    BtnSearch.Enabled = true;
                    numericUpDown2.Enabled = true;
                    BtnImportExcel.Enabled = true;
                }
            }
            else
            {
                BtnSearch.Enabled = true;
                numericUpDown2.Enabled = true;
                BtnImportExcel.Enabled = true;
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
                var lottery = index + " :شماره "
                    + dtExcelSearch.Rows[index - 1][3].ToString() + " :شماره فیش "
                     + dtExcelSearch.Rows[index - 1][1].ToString() + " :کد نوسازی " + "\n";

                if (!IsFoundLottery(lottery))
                {
                    if (_OptSelection != null)
                    {
                        _OptSelection.SaveReslt(lottery, 2, false);
                    }
                    else
                    {
                        _OptSelection = new UtilLottrey();
                        _OptSelection.SaveReslt(lottery, 2, false);
                    }
                    listBox2.Items.Add(lottery);

                    lottery = "شهرداری قم" + "\n";
                    lottery += "برنده خوش شانش" + "\n";
                    lottery += index + " :شماره " + "\n"
                        + dtExcelSearch.Rows[index - 1][3].ToString() + " :شماره فیش " + "\n"
                        + dtExcelSearch.Rows[index - 1][1].ToString() + " :کد نوسازی " + "\n";
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
            //var tt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //OptLottery = new UtilLottrey();
            //OptLottery.SaveReslt("werwerwerwerw", 1);
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var Result = string.Empty;
            if (listBox1.Items.Count <= 0)
            {
                return;
            }
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                Result += i + 1 + " - " + listBox1.Items[i].ToString() + Environment.NewLine;
            }
            Clipboard.SetText(Result);
            MessageBox.Show("به حافظه منتقل شد.");
        }

        private void BtnLotteryCancel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("آیا مایل به توقف عملیات ورود هستید؟", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (_OptLottery != null)
                {
                    _OptLottery.OptLottery = true;
                }
            }
        }

        private void BtnSelectionCancel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("آیا مایل به توقف عملیات ورود هستید؟", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (_OptSelection != null)
                {
                    _OptSelection.OptLottery = true;
                }
            }
        }

        private void BtnRemoveLottery_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0)
            {
                if (MessageBox.Show("آیا مایل به حذف هستید؟", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    var item = listBox1.Items[listBox1.SelectedIndex].ToString();
                    listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                    if (_OptLottery != null)
                    {
                        var text = "";
                        for (int i = 0; i < listBox1.Items.Count; i++)
                        {
                            text += listBox1.Items[i].ToString() + "\n";
                        }
                        _OptLottery.SaveReslt(text, 1, true);
                        //for (int i = 0; i < dtExcelLottery_Removed.Rows.Count; i++)
                        //{
                        //    if (dtExcelLottery_Removed.Rows[i][3].ToString() == item)
                        //    {
                        //        dtExcelLottery.Rows.Add(dtExcelLottery_Removed.Rows[i]);
                        //    }
                        //}
                    }
                }
            }
        }

        private void BtnRemoveSelection_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count > 0)
            {
                if (MessageBox.Show("آیا مایل به حذف هستید؟", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    listBox2.Items.RemoveAt(listBox2.SelectedIndex);
                    if (_OptSelection != null)
                    {
                        var text = "";
                        for (int i = 0; i < listBox2.Items.Count; i++)
                        {
                            text += listBox2.Items[i].ToString() + "\n";
                        }
                        _OptSelection.SaveReslt(text, 2, true);
                    }
                }
            }
        }
    }
}
