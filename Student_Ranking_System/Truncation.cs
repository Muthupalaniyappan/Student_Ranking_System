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
    public partial class Truncation : Form
    {
        public Truncation()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += Truncation_FormClosed;
        }

        private void Truncation_FormClosed(object sender, FormClosedEventArgs e)
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

                OleDbCommand truncate = new OleDbCommand("DELETE FROM " + db + " where Standard=? and Sec=?", dbconn);
                truncate.Parameters.AddWithValue("Standard", comboBox2.Text);
                truncate.Parameters.AddWithValue("Sec", comboBox3.Text);

                if(truncate.ExecuteNonQuery()>0)
                    MessageBox.Show("DB is truncated for the next year usage", "Information", MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                else
                    MessageBox.Show("DB is already truncated", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void Truncation_Load(object sender, EventArgs e)
        {

        }
    }
}
