using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Student_Ranking_System
{
    public partial class splashscreen : Form
    {
        public splashscreen()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            this.Hide();
            login.Show();
        }
        int count = 1;
        private void button1_Click(object sender, EventArgs e)
        {

            if (count == 1)
            {
                pictureBox1.Image = Image.FromFile("splash2.jpg");
                count++;
            }
            else if (count == 2)
            {
                pictureBox1.Image = Image.FromFile("splash3.jpg");
                count++;
            }
            else if (count == 3)
            {
                pictureBox1.Image = Image.FromFile("splash4.jpg");
                count++;
            }
            else if (count == 4)
            {
                pictureBox1.Image = Image.FromFile("splash5.jpg");
            }
            //if(c==1)
            //{
            //    Login login = new Login();
            //    this.Hide();
            //    login.Show();
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (count == 1)
            {
                pictureBox1.Image = Image.FromFile("splash1.jpg");
            }
            else if (count == 2)
            {
                pictureBox1.Image = Image.FromFile("splash2.jpg");
                count--;
            }
            else if (count == 3)
            {
                pictureBox1.Image = Image.FromFile("splash3.jpg");
                count--;
            }
            else if (count == 4)
            {
                pictureBox1.Image = Image.FromFile("splash4.jpg");
                count--;
            }
        }
    }
}
