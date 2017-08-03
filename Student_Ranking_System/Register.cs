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
    public partial class Register : Form
    {
        public Register()
        {
            InitializeComponent();
        }

        private void label6_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            int flag = 0;

            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                OleDbCommand roll = new OleDbCommand("select RollNo,Standard,Sec from SDetails",dbconn);
                var data = roll.ExecuteReader();
                
                while(data.Read()!=false)
                {
                    if(data.GetInt32(0).ToString()==textBox2.Text && data.GetInt32(1).ToString() == comboBox1.Text && data.GetString(2) == comboBox2.Text)
                    {
                        flag = 1;
                    }
                }

                if(flag == 0)
                {
                    OleDbCommand register = new OleDbCommand("insert into SDetails(Student_name,RollNo,Standard,Sec,PhoneNumber) values('" + textBox1.Text + "','" + textBox2.Text + "','" + comboBox1.Text + "','" + comboBox2.Text + "','" + textBox5.Text + "')", dbconn);
                    register.ExecuteNonQuery();
                    MessageBox.Show("Reistered", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox5.Text = "";
                    comboBox1.Text = "";
                    comboBox2.Text = "";
                }
                else
                {
                    MessageBox.Show("RollNo already Registered","Information",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    textBox2.Text = "";
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(""+ex,"Error",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
            }
        }
    }
}
