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
    public partial class Inserting : Form
    {
        public Inserting()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += Inserting_FormClosed;
        }

        private void Inserting_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int mark1 = int.Parse(textBox3.Text);
                int mark2 = int.Parse(textBox4.Text);
                int mark3 = int.Parse(textBox5.Text);
                int mark4 = int.Parse(textBox6.Text);
                int mark5 = int.Parse(textBox7.Text);
                int mark6 = int.Parse(textBox8.Text);
                int prac1 = int.Parse(textBox10.Text);
                int prac2 = int.Parse(textBox11.Text);
                int prac3 = int.Parse(textBox12.Text);

                int total = mark1 + mark2 + mark3 + mark4 + mark5 + mark6 + prac1 + prac2 + prac3;

                int failcount = 0;


                if (mark1 < 70)
                    failcount++;
                if (mark2 < 70)
                    failcount++;
                if (mark3 < 70)
                    failcount++;
                if (mark4 < 30 || prac1 < 40)
                    failcount++;
                if (mark5 < 30 || prac2 < 40)
                    failcount++;
                if (mark6 < 30 || prac3 < 40)
                    failcount++;

                textBox9.Text = total.ToString();
                textBox13.Text = failcount.ToString();
            }
           catch(Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                string db = comboBox1.Text;

                int mark1 = int.Parse(textBox3.Text);
                int mark2 = int.Parse(textBox4.Text);
                int mark3 = int.Parse(textBox5.Text);
                int mark4 = int.Parse(textBox6.Text);
                int mark5 = int.Parse(textBox7.Text);
                int mark6 = int.Parse(textBox8.Text);
                int prac1 = int.Parse(textBox10.Text);
                int prac2 = int.Parse(textBox11.Text);
                int prac3 = int.Parse(textBox12.Text);

                int physics = mark4 + prac1;
                int chemistry = mark5 + prac2;
                int csorbio = mark6 + prac3;

                int flag = 0;

                OleDbCommand rollno = new OleDbCommand("select RollNo,Standard,Sec from " + db, dbconn);

                var data = rollno.ExecuteReader();

                while(data.Read()!=false)
                {
                    if(data.GetInt32(0).ToString()==textBox2.Text && data.GetInt32(1).ToString() == comboBox2.Text && data.GetString(2) == comboBox3.Text)
                    {
                        flag = 1;
                    }
                }

                if (flag == 0)
                {
                    if ( mark1 <= 200 && mark2 <= 200 && mark3 <=200 && mark4 <=150 && mark5<=150 && mark6 <=150 && prac1 <=50 && prac2 <=50 && prac3 <=50)
                    {
                        OleDbCommand sname = new OleDbCommand("select Student_name from SDetails where RollNo=? and Standard=? and Sec = ?",dbconn);
                        sname.Parameters.AddWithValue("RollNo", textBox2.Text);
                        sname.Parameters.AddWithValue("Standard",comboBox2.Text);
                        sname.Parameters.AddWithValue("Sec",comboBox3.Text);
                        var datas = sname.ExecuteReader();
                        datas.Read();
                        String Studentname = datas.GetString(0);

                        OleDbCommand insertquery = new OleDbCommand("insert into " + db + " (Student_name,Tamil,English,Maths,Physics,Chemistry,CsORBio,Total,RollNo,Failcount,PhysicsT,ChemistryT,CsORBioT,PhysicsP,ChemistryP,CsORBioP,Standard,Sec) values ('" + Studentname + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + physics + "','" + chemistry + "','" + csorbio + "','" + textBox9.Text + "','" + textBox2.Text + "','" + textBox13.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + textBox10.Text + "','" + textBox11.Text + "','" + textBox12.Text + "','" + comboBox2.Text + "','" + comboBox3.Text + "')", dbconn);
                        insertquery.ExecuteNonQuery();
                        MessageBox.Show("Record Inserted", "Information", MessageBoxButtons.OK,MessageBoxIcon.Information);

                       // textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                        textBox4.Text = "";
                        textBox5.Text = "";
                        textBox6.Text = "";
                        textBox7.Text = "";
                        textBox8.Text = "";
                        textBox9.Text = "";
                        textBox10.Text = "";
                        textBox11.Text = "";
                        textBox12.Text = "";
                        textBox13.Text = "";

                    }
                    else
                    {
                        MessageBox.Show("Please enter mark below 200 for non practical subject and below 150 for practical subject and enter practical mark in practical column", "Information", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("Roll No: " + textBox2.Text + " already exist", "Information", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    textBox2.Text = "";
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(" "+ex, "Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void Inserting_Load(object sender, EventArgs e)
        {

        }
    }
}
