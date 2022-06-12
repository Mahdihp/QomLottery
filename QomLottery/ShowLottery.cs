using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QomLottery
{
    public partial class ShowLottery : Form
    {
        public ShowLottery(string LotteryFound)
        {
            InitializeComponent();
            this.LotteryFound = LotteryFound;
        }
        public string LotteryFound { get; set; }
        private async void ShowLottery_Load(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            label1.Visible = false;
            SetLocation();
            await Task.Delay(3000);
           
            label1.Visible = true;
            button1.Visible = true;
            label1.Text = LotteryFound;
            pictureBox1.Image = QomLottery.Properties.Resources.conffeti;

            pictureBox1.Location = new Point(0, 0);
            pictureBox1.Height = this.Height;
            pictureBox1.Width = this.Width;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
           
        }

        private void SetLocation()
        {
            var xl = (this.Width / 2) - (label1.Width / 2);
            var yl = (this.Height / 2) - (label1.Height + label1.Height);

            var xb = (this.Width / 2) - (button1.Width / 2);
            var yb = (this.Height / 2) - (button1.Height /5);

            var xp = (this.Width / 2) - (pictureBox1.Width / 2);
            var yp = (this.Height / 2) - (pictureBox1.Height / 2);
            label1.Size = new Size(this.Width, this.Height / 2);
            pictureBox1.Location = new Point(xp, yp);
           // label1.Location = new Point(xl, yl);
          //  button1.Location = new Point(xb, yb);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ShowLottery_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Escape)
            {
                this.Close();
            }
        }
    }
}
