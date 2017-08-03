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
    public partial class StudentComparition : Form
    {
        public StudentComparition()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
            dbconn.Open();

            string db = comboBox1.Text;

            OleDbCommand Student_1_first = new OleDbCommand("select Tamil,English,Maths,Physics,Chemistry,CsORBio,Total from " + db + " where RollNo = ? and Standard = ? and Sec = ?",dbconn);
            Student_1_first.Parameters.AddWithValue("RollNo",textBox1.Text);
            Student_1_first.Parameters.AddWithValue("Standard", comboBox2.Text);
            Student_1_first.Parameters.AddWithValue("Sec", comboBox3.Text);

            var student_1_data = Student_1_first.ExecuteReader();

            student_1_data.Read();

            OleDbCommand Student_2_first = new OleDbCommand("select Tamil,English,Maths,Physics,Chemistry,CsORBio,Total from " + db + " where RollNo = ? and Standard = ? and Sec = ?",dbconn);
            Student_2_first.Parameters.AddWithValue("RollNo", textBox2.Text);
            Student_2_first.Parameters.AddWithValue("Standard", comboBox2.Text);
            Student_2_first.Parameters.AddWithValue("Sec", comboBox3.Text);

            var student_2_data = Student_2_first.ExecuteReader();

            student_2_data.Read();


            this.chart1.Series["Student_1"].Points.AddXY("Tamil",student_1_data.GetInt32(0));
            this.chart1.Series["Student_1"].Points[0].Label = student_1_data.GetInt32(0).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("Tamil", student_2_data.GetInt32(0));
            this.chart1.Series["Student_2"].Points[0].Label = student_2_data.GetInt32(0).ToString();



            this.chart1.Series["Student_1"].Points.AddXY("English", student_1_data.GetInt32(1));
            this.chart1.Series["Student_1"].Points[1].Label = student_1_data.GetInt32(1).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("English", student_2_data.GetInt32(1));
            this.chart1.Series["Student_2"].Points[1].Label = student_2_data.GetInt32(1).ToString();



            this.chart1.Series["Student_1"].Points.AddXY("Maths", student_1_data.GetInt32(2));
            this.chart1.Series["Student_1"].Points[2].Label = student_1_data.GetInt32(2).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("Maths", student_2_data.GetInt32(2));
            this.chart1.Series["Student_2"].Points[2].Label = student_2_data.GetInt32(2).ToString();



            this.chart1.Series["Student_1"].Points.AddXY("Physics", student_1_data.GetInt32(3));
            this.chart1.Series["Student_1"].Points[3].Label = student_1_data.GetInt32(3).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("Physics", student_2_data.GetInt32(3));
            this.chart1.Series["Student_2"].Points[3].Label = student_2_data.GetInt32(3).ToString();



            this.chart1.Series["Student_1"].Points.AddXY("Chemistry", student_1_data.GetInt32(4));
            this.chart1.Series["Student_1"].Points[4].Label = student_1_data.GetInt32(4).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("Chemistry", student_2_data.GetInt32(4));
            this.chart1.Series["Student_2"].Points[4].Label = student_2_data.GetInt32(4).ToString();



            this.chart1.Series["Student_1"].Points.AddXY("Cs/Bio", student_1_data.GetInt32(5));
            this.chart1.Series["Student_1"].Points[5].Label = student_1_data.GetInt32(5).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("Cs/Bio", student_2_data.GetInt32(5));
            this.chart1.Series["Student_2"].Points[5].Label = student_2_data.GetInt32(5).ToString();



            this.chart1.Series["Student_1"].Points.AddXY("Total", student_1_data.GetInt32(6));
            this.chart1.Series["Student_1"].Points[6].Label = student_1_data.GetInt32(6).ToString();
            this.chart1.Series["Student_2"].Points.AddXY("Total", student_2_data.GetInt32(6));
            this.chart1.Series["Student_2"].Points[6].Label = student_2_data.GetInt32(6).ToString();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.chart1.Series["Student_1"].Points.Clear();
            this.chart1.Series["Student_2"].Points.Clear();
        }

        private void label9_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
