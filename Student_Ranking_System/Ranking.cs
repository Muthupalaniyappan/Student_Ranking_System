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
    public partial class Ranking : Form
    {
        public Ranking()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += Ranking_FormClosed;
        }

        private void Ranking_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Ranking_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                string db = comboBox1.Text;

                OleDbCommand retrive = new OleDbCommand("select Student_name,Tamil,English,Maths,Physics,Chemistry,CsORBio,Total,RollNo,Failcount from " + db + " where Standard=? and Sec=? order by Failcount asc,Total desc", dbconn);

                retrive.Parameters.AddWithValue("Standard", comboBox2.Text);
                retrive.Parameters.AddWithValue("Sec", comboBox3.Text);

                int count=0;
                int oldtotal = 0;
                int c = 1;
                int flag = 1;
                int failflag = -1;


                var data = retrive.ExecuteReader();
                while (data.Read() != false)
                {
                    var r_no = "" + data.GetInt32(8);
                    //if(data.GetInt32(1)>=70&& data.GetInt32(2)>=70&& data.GetInt32(3)>=70 && data.GetInt32(4)>=70 && data.GetInt32(5)>=70 && data.GetInt32(6)>=70)
                    //{
                        if (data.GetInt32(7) != oldtotal || (data.GetInt32(7) == oldtotal && data.GetInt32(9)!=failflag ))
                        {
                            if (flag == 1)
                                count++;
                            else
                            {
                                count = count + c;
                                flag = 1;
                                c = 1;
                            }
                        }
                        else
                        {
                            c++;
                            flag = 0;
                        }


                    OleDbCommand updater = new OleDbCommand("UPDATE " + db + " SET Rank = ? WHERE RollNo = ? and Standard=? and Sec=?", dbconn);
                    updater.Parameters.AddWithValue("Rank", count);
                    updater.Parameters.AddWithValue("RollNo", r_no);
                    updater.Parameters.AddWithValue("Standard", comboBox2.Text);
                    updater.Parameters.AddWithValue("Sec", comboBox3.Text);
                    updater.ExecuteNonQuery();
                    oldtotal = data.GetInt32(7);
                    failflag = data.GetInt32(9);
                    //}
                }


                MessageBox.Show("Rank Generated", "Information", MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                //dbconn.Close();
            }
            catch (Exception ex)
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
    }
}
