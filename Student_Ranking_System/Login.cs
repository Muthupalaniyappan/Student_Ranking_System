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
    public partial class Login : Form
    {
        public Login()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            timer1.Start();
            this.FormClosed += Login_FormClosed;
        }

        private void Login_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username;
            string password;
            string message = "invalid username or password,please try again";

            username = textBox1.Text;
            password = textBox2.Text;


            try {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                OleDbCommand sqlquery = new OleDbCommand("SELECT Username, Password FROM Login_DB WHERE Username = ? AND Password = ?", dbconn);

                sqlquery.Parameters.AddWithValue("Username", username);
                sqlquery.Parameters.AddWithValue("Password", password);

                if (sqlquery.ExecuteScalar() != null)
                {
                    dbconn.Close();
                    this.Hide();
                    Std_Selection s = new Std_Selection();
                    s.Show();
                    textBox1.Clear();
                    textBox2.Clear();
                    MessageBox.Show("welcome user " + username, "information", MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                }
                else
                {
                    dbconn.Close();
                    MessageBox.Show(message, "information", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    textBox1.Clear();
                    textBox2.Clear();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            label3.Text = "" + dt;
        }
    }
}
