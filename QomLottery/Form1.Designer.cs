namespace QomLottery
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnRemoveLottery = new System.Windows.Forms.Button();
            this.BtnLottery = new System.Windows.Forms.Button();
            this.BtnLotteryCancel = new System.Windows.Forms.Button();
            this.BtnImport = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.BtnRemoveSelection = new System.Windows.Forms.Button();
            this.BtnSelectionCancel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.BtnSearch = new System.Windows.Forms.Button();
            this.BtnImportExcel = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.listBox1.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 27;
            this.listBox1.Location = new System.Drawing.Point(3, 140);
            this.listBox1.Name = "listBox1";
            this.listBox1.ScrollAlwaysVisible = true;
            this.listBox1.Size = new System.Drawing.Size(821, 166);
            this.listBox1.TabIndex = 1;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Font = new System.Drawing.Font("Vazirmatn", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(835, 348);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.numericUpDown1);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.BtnRemoveLottery);
            this.tabPage1.Controls.Add(this.BtnLottery);
            this.tabPage1.Controls.Add(this.BtnLotteryCancel);
            this.tabPage1.Controls.Add(this.BtnImport);
            this.tabPage1.Controls.Add(this.listBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 35);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(827, 309);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "قرعه کشی تصادفی";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Font = new System.Drawing.Font("Vazirmatn FD", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numericUpDown1.Location = new System.Drawing.Point(8, 18);
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(124, 40);
            this.numericUpDown1.TabIndex = 6;
            this.numericUpDown1.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(139, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(117, 40);
            this.label2.TabIndex = 5;
            this.label2.Text = ":تعداد برندگان";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Vazirmatn FD", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(144, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(499, 54);
            this.label1.TabIndex = 4;
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BtnRemoveLottery
            // 
            this.BtnRemoveLottery.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnRemoveLottery.Location = new System.Drawing.Point(8, 80);
            this.BtnRemoveLottery.Name = "BtnRemoveLottery";
            this.BtnRemoveLottery.Size = new System.Drawing.Size(84, 54);
            this.BtnRemoveLottery.TabIndex = 1;
            this.BtnRemoveLottery.Text = "حذف از لیست";
            this.BtnRemoveLottery.UseVisualStyleBackColor = true;
            this.BtnRemoveLottery.Visible = false;
            this.BtnRemoveLottery.Click += new System.EventHandler(this.BtnRemoveLottery_Click);
            // 
            // BtnLottery
            // 
            this.BtnLottery.Enabled = false;
            this.BtnLottery.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnLottery.Location = new System.Drawing.Point(649, 69);
            this.BtnLottery.Name = "BtnLottery";
            this.BtnLottery.Size = new System.Drawing.Size(164, 54);
            this.BtnLottery.TabIndex = 1;
            this.BtnLottery.Text = "قرعه کشی";
            this.BtnLottery.UseVisualStyleBackColor = true;
            this.BtnLottery.Click += new System.EventHandler(this.BtnLottery_Click);
            // 
            // BtnLotteryCancel
            // 
            this.BtnLotteryCancel.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnLotteryCancel.Location = new System.Drawing.Point(479, 9);
            this.BtnLotteryCancel.Name = "BtnLotteryCancel";
            this.BtnLotteryCancel.Size = new System.Drawing.Size(164, 54);
            this.BtnLotteryCancel.TabIndex = 2;
            this.BtnLotteryCancel.Text = "انصراف";
            this.BtnLotteryCancel.UseVisualStyleBackColor = true;
            this.BtnLotteryCancel.Click += new System.EventHandler(this.BtnLotteryCancel_Click);
            // 
            // BtnImport
            // 
            this.BtnImport.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnImport.Location = new System.Drawing.Point(649, 9);
            this.BtnImport.Name = "BtnImport";
            this.BtnImport.Size = new System.Drawing.Size(164, 54);
            this.BtnImport.TabIndex = 2;
            this.BtnImport.Text = "ورود اطلاعات فیش";
            this.BtnImport.UseVisualStyleBackColor = true;
            this.BtnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.BtnRemoveSelection);
            this.tabPage2.Controls.Add(this.BtnSelectionCancel);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.listBox2);
            this.tabPage2.Controls.Add(this.numericUpDown2);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.BtnSearch);
            this.tabPage2.Controls.Add(this.BtnImportExcel);
            this.tabPage2.Location = new System.Drawing.Point(4, 35);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(827, 309);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "قرعه کشی براساس شماره";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.Click += new System.EventHandler(this.tabPage2_Click);
            // 
            // BtnRemoveSelection
            // 
            this.BtnRemoveSelection.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnRemoveSelection.Location = new System.Drawing.Point(178, 81);
            this.BtnRemoveSelection.Name = "BtnRemoveSelection";
            this.BtnRemoveSelection.Size = new System.Drawing.Size(84, 52);
            this.BtnRemoveSelection.TabIndex = 11;
            this.BtnRemoveSelection.Text = "حذف از لیست";
            this.BtnRemoveSelection.UseVisualStyleBackColor = true;
            this.BtnRemoveSelection.Click += new System.EventHandler(this.BtnRemoveSelection_Click);
            // 
            // BtnSelectionCancel
            // 
            this.BtnSelectionCancel.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnSelectionCancel.Location = new System.Drawing.Point(485, 5);
            this.BtnSelectionCancel.Name = "BtnSelectionCancel";
            this.BtnSelectionCancel.Size = new System.Drawing.Size(164, 54);
            this.BtnSelectionCancel.TabIndex = 10;
            this.BtnSelectionCancel.Text = "انصراف";
            this.BtnSelectionCancel.UseVisualStyleBackColor = true;
            this.BtnSelectionCancel.Click += new System.EventHandler(this.BtnSelectionCancel_Click);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Vazirmatn FD", 15.75F);
            this.label4.Location = new System.Drawing.Point(348, 81);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(471, 52);
            this.label4.TabIndex = 8;
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // listBox2
            // 
            this.listBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.listBox2.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 27;
            this.listBox2.Location = new System.Drawing.Point(3, 140);
            this.listBox2.Name = "listBox2";
            this.listBox2.ScrollAlwaysVisible = true;
            this.listBox2.Size = new System.Drawing.Size(821, 166);
            this.listBox2.TabIndex = 7;
            // 
            // numericUpDown2
            // 
            this.numericUpDown2.Enabled = false;
            this.numericUpDown2.Font = new System.Drawing.Font("Vazirmatn FD", 15.75F);
            this.numericUpDown2.Location = new System.Drawing.Point(8, 14);
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(246, 40);
            this.numericUpDown2.TabIndex = 6;
            this.numericUpDown2.ValueChanged += new System.EventHandler(this.numericUpDown2_ValueChanged);
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(260, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(200, 40);
            this.label3.TabIndex = 5;
            this.label3.Text = ":جستجو براساس شماره ردیف";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BtnSearch
            // 
            this.BtnSearch.Enabled = false;
            this.BtnSearch.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnSearch.Location = new System.Drawing.Point(8, 81);
            this.BtnSearch.Name = "BtnSearch";
            this.BtnSearch.Size = new System.Drawing.Size(164, 54);
            this.BtnSearch.TabIndex = 3;
            this.BtnSearch.Text = "جستجو";
            this.BtnSearch.UseVisualStyleBackColor = true;
            this.BtnSearch.Click += new System.EventHandler(this.BtnSearch_Click);
            // 
            // BtnImportExcel
            // 
            this.BtnImportExcel.Font = new System.Drawing.Font("Vazirmatn", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnImportExcel.Location = new System.Drawing.Point(655, 5);
            this.BtnImportExcel.Name = "BtnImportExcel";
            this.BtnImportExcel.Size = new System.Drawing.Size(164, 54);
            this.BtnImportExcel.TabIndex = 3;
            this.BtnImportExcel.Text = "ورود اطلاعات فیش";
            this.BtnImportExcel.UseVisualStyleBackColor = true;
            this.BtnImportExcel.Click += new System.EventHandler(this.BtnImportExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(835, 348);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "سامانه قرعه کشی شهرداری قم";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnLottery;
        private System.Windows.Forms.Button BtnImport;
        private System.Windows.Forms.Button BtnImportExcel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private System.Windows.Forms.Button BtnSearch;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button BtnLotteryCancel;
        private System.Windows.Forms.Button BtnSelectionCancel;
        private System.Windows.Forms.Button BtnRemoveLottery;
        private System.Windows.Forms.Button BtnRemoveSelection;
    }
}

