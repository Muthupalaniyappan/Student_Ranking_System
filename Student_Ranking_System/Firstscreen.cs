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
    public partial class Firstscreen : Form
    {
        Timer t = new Timer();
        double pbunit;

        int pbwidth, pbheight, pbcomplete;

        Bitmap bm;
        Graphics g;

        public Firstscreen()
        {
            InitializeComponent();
        }

        private void Firstscreen_Load(object sender, EventArgs e)
        {
            pbwidth = pictureBox3.Width;
            pbheight = pictureBox3.Height;

            pbunit = pbwidth / 100.0;

            pbcomplete = 0;

            bm = new Bitmap(pbwidth, pbheight);

            t.Interval = 100;
            t.Tick += new EventHandler(this.t_tick);
            t.Start();

        }

        private void t_tick(object sender, EventArgs e)
        {
            g = Graphics.FromImage(bm);
            g.FillRectangle(Brushes.LightSkyBlue, new Rectangle(0, 0, (int)(pbcomplete * pbunit),pbheight));

            pictureBox3.Image = bm;
            pbcomplete++;
            if(pbcomplete>100)
            {
                g.Dispose();
                t.Stop();
                splashscreen splash = new splashscreen();
                this.Hide();
                splash.Show();

            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
