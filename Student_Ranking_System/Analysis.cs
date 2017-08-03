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
    public partial class Analysis : Form
    {
        public Analysis()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += Analysis_FormClosed;
        }

        private void Analysis_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            label11.Text = "";
            label12.Text = "";
            label13.Text = "";
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            label18.Text = "";
            label20.Text = "";
            
            var combo1 = comboBox1.Text;
            var combo2 = comboBox2.Text;
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                OleDbCommand selection = new OleDbCommand("select " + combo1 + " from " + combo2 + " where Standard=? and Sec=?", dbconn);
                selection.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection.Parameters.AddWithValue("Sec", comboBox4.Text);

                float average = 0.0f;
                float total = 0.0f, counter = 0.0f;
                var data1 = selection.ExecuteReader();
                while(data1.Read()!= false)
                {
                    if(data1.GetInt32(0)!=0)
                    {
                        total = total + data1.GetInt32(0);
                        counter++;
                    }
                }

                average = total / counter;
                label16.Text = "" + average; 


                OleDbCommand Max = new OleDbCommand("select MAX(" + combo1 + ") from " + combo2 + " where Standard=? and Sec=?", dbconn);
                Max.Parameters.AddWithValue("Standard", comboBox3.Text);
                Max.Parameters.AddWithValue("Sec", comboBox4.Text);

                label14.Text = "" + Max.ExecuteScalar();



                OleDbCommand Min = new OleDbCommand("select MIN(" + combo1 + ") from " + combo2 + " where Standard=? and Sec=?", dbconn);
                Min.Parameters.AddWithValue("Standard", comboBox3.Text);
                Min.Parameters.AddWithValue("Sec", comboBox4.Text);

                label15.Text = "" + Min.ExecuteScalar();



                OleDbCommand count = new OleDbCommand("select count(" + combo1 + ") from " + combo2+ " where Standard=? and Sec=?", dbconn);
                count.Parameters.AddWithValue("Standard", comboBox3.Text);
                count.Parameters.AddWithValue("Sec", comboBox4.Text);

                label11.Text = "" + count.ExecuteScalar();


                int npass = 0;
                int fail = 0;
                int abs = 0;

                if (combo1 == "Tamil" || combo1 == "English" || combo1 == "Maths")
                {
                    OleDbCommand pass = new OleDbCommand("select " + combo1 + ",RollNo,Student_name from " + combo2 + " where Standard=? and Sec=?", dbconn);
                    pass.Parameters.AddWithValue("Standard", comboBox3.Text);
                    pass.Parameters.AddWithValue("Sec", comboBox4.Text);
                    var data = pass.ExecuteReader();
                    while(data.Read()!=false)
                    {
                        if(data.GetInt32(0)>=70)
                        {
                            npass++;
                        }
                        else if(data.GetInt32(0) == 0)
                        {
                            label20.Text += "RollNo :" + data.GetInt32(1) + "  Name :" + data.GetString(2) + "\n";
                            abs++;
                        }
                        else
                        {
                            label18.Text += "RollNo :" + data.GetInt32(1) + "  Name :" + data.GetString(2) + "\n";
                            fail++;
                        }
                    }

                    label12.Text = "" + npass;
                    label13.Text = "" + fail;
                    label17.Text = "" + abs;

                    float Tcount = float.Parse(label11.Text);
                    float Pcount = float.Parse(label12.Text);
                    float percent = (Pcount / Tcount) * 100;

                    label24.Text = "" + percent;
                }
                else
                {
                    var theory = combo1 + "T";
                    var practical = combo1 + "P";

                    OleDbCommand pass1 = new OleDbCommand("select " + theory + "," + practical + ",RollNo,Student_name from " + combo2 + " where Standard=? and Sec=?", dbconn);
                    pass1.Parameters.AddWithValue("Standard", comboBox3.Text);
                    pass1.Parameters.AddWithValue("Sec", comboBox4.Text);
                    var data = pass1.ExecuteReader();
                    while (data.Read() != false)
                    {
                        if (data.GetInt32(0) >= 30 && data.GetInt32(1) >= 40) 
                        {
                            npass++;
                        }
                        else if (data.GetInt32(0) == 0)
                        {
                            label20.Text += "RollNo :" + data.GetInt32(2) + "  Name :" + data.GetString(3) + "\n";
                            abs++;
                        }
                        else
                        {
                            label18.Text += "RollNo :" + data.GetInt32(2) + "  Name :" + data.GetString(3) + "\n";
                            fail++;
                        }
                    }

                    label12.Text = "" + npass;
                    label13.Text = "" + fail;
                    label17.Text = "" + abs;


                    float Tcount = float.Parse(label11.Text);
                    float Pcount = float.Parse(label12.Text);
                    float percent = (Pcount / Tcount) * 100;
  
                    label24.Text = "" + percent;

                }

                OleDbCommand totalavg = new OleDbCommand("select AVG(Total) from " + combo2 + " where Standard=? and Sec=?", dbconn);
                totalavg.Parameters.AddWithValue("Standard", comboBox3.Text);
                totalavg.Parameters.AddWithValue("Sec", comboBox4.Text);

                label22.Text = "" + totalavg.ExecuteScalar();

            }
            catch (Exception ex)
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

        private void Analysis_Load(object sender, EventArgs e)
        {

        }
    }
}
