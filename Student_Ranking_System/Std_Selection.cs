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
    public partial class Std_Selection : Form
    {
        public Std_Selection()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += Std_Selection_FormClosed;
        }

        private void Std_Selection_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Std_Selection_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Inserting insert = new Inserting();
            this.Hide();
            insert.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _12th_STD_Details retrive = new _12th_STD_Details();
            this.Hide();
            retrive.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Ranking rank = new Ranking();
            this.Hide();
            rank.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Deletion deletion = new Deletion();
            this.Hide();
            deletion.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            this.Hide();
            login.Show();
            MessageBox.Show("Successfully signedout","Information",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Truncation truncate = new Truncation();
            this.Hide();
            truncate.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            TotalAnalysis total = new TotalAnalysis();
            this.Hide();
            total.Show();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SubjectAnalysis subject = new SubjectAnalysis();
            this.Hide();
            subject.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Analysis analysis = new Analysis();
            this.Hide();
            analysis.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MarkRegister mr = new MarkRegister();
            this.Hide();
            mr.Show();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Register r = new Register();
            r.Show();
        }
    }
}
