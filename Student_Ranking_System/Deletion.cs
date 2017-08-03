using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Student_Ranking_System
{
    public partial class Deletion : Form
    {
        public Deletion()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += Deletion_FormClosed;
        }

        private void Deletion_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                string db = comboBox1.Text;

                OleDbCommand deletion = new OleDbCommand("DELETE FROM " + db + " WHERE RollNo = ? and Standard=? and Sec=?", dbconn);
                deletion.Parameters.AddWithValue("RollNo", textBox1.Text);
                deletion.Parameters.AddWithValue("Standard", comboBox2.Text);
                deletion.Parameters.AddWithValue("Sec", comboBox3.Text);
                if(deletion.ExecuteNonQuery()>0)
                    MessageBox.Show("RollNo: " + textBox1.Text + " is deleted", "Information", MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                else
                    MessageBox.Show("RollNo: " + textBox1.Text + " is not there", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Text = "";
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void Deletion_Load(object sender, EventArgs e)
        {

        }
    }
}
