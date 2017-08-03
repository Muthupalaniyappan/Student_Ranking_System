using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;

using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Windows.Forms.DataVisualization.Charting;

namespace Student_Ranking_System
{
    public partial class TotalAnalysis : Form
    {
        public TotalAnalysis()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += TotalAnalysis_FormClosed;
        }
        int count = 0;
        private void TotalAnalysis_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void TotalAnalysis_Load(object sender, EventArgs e)
        {
            
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            count = 5;
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                var combo = comboBox1.Text;

                OleDbCommand selection = new OleDbCommand("select " + combo + " from FirstMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data = selection.ExecuteReader();
                data.Read();

                float first = data.GetInt32(0);

                OleDbCommand selection1 = new OleDbCommand("select " + combo + " from Quaterly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection1.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection1.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection1.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data1 = selection1.ExecuteReader();
                data1.Read();

                float quaterly = data1.GetInt32(0);

                OleDbCommand selection2 = new OleDbCommand("select " + combo + " from SecondMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection2.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection2.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection2.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data2 = selection2.ExecuteReader();
                data2.Read();

                float second = data2.GetInt32(0);

                OleDbCommand selection3 = new OleDbCommand("select " + combo + " from Halfyearly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection3.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection3.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection3.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data3 = selection3.ExecuteReader();
                data3.Read();

                float Halfyearly = data3.GetInt32(0);

                OleDbCommand selection4 = new OleDbCommand("select " + combo + " from ThirdMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection4.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection4.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection4.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data4 = selection4.ExecuteReader();
                data4.Read();

                float third = data4.GetInt32(0);

                OleDbCommand selection5 = new OleDbCommand("select " + combo + " from Annual where RollNo = ? and Standard=? and Sec=?", dbconn);
   
                selection5.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection5.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection5.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data5 = selection5.ExecuteReader();
                data5.Read();

                float annual = data5.GetInt32(0);

                this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Halfyearly", data3.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("ThirdMidTerm", data4.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Annual", data5.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                this.chart1.Series[comboBox1.Text].Points[1].Label = quaterly.ToString();
                this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                this.chart1.Series[comboBox1.Text].Points[2].Label = second.ToString();
                this.chart1.Series[comboBox1.Text].Points[3].AxisLabel = "Halfyearly";
                this.chart1.Series[comboBox1.Text].Points[3].Label = Halfyearly.ToString();
                this.chart1.Series[comboBox1.Text].Points[4].AxisLabel = "ThirdMidTerm";
                this.chart1.Series[comboBox1.Text].Points[4].Label = third.ToString();
                this.chart1.Series[comboBox1.Text].Points[5].AxisLabel = "Annual";
                this.chart1.Series[comboBox1.Text].Points[5].Label = annual.ToString();

                float average = (first + second + third + quaterly + Halfyearly + annual)/6;

                this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Halfyearly", average);
                this.chart1.Series["Average"].Points.AddXY("ThirdMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Annual", average);

                label2.Text = "Average : " + average;

                this.chart2.Series[combo].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Halfyearly", data3.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("ThirdMidTerm", data4.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Annual", data5.GetInt32(0));


            }
            catch(Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var combo = comboBox1.Text;
                this.chart1.Series[combo].Points.Clear();
                this.chart1.Series["Average"].Points.Clear();
                this.chart2.Series[combo].Points.Clear();
                label2.Text = "";
            }
            catch(Exception ex)
            {
                MessageBox.Show("Select the correct category to unload " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            count = 1;

            try {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                var combo = comboBox1.Text;

                OleDbCommand selection = new OleDbCommand("select " + combo + " from FirstMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data = selection.ExecuteReader();
                data.Read();

                float first = data.GetInt32(0);

                OleDbCommand selection1 = new OleDbCommand("select " + combo + " from Quaterly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection1.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection1.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection1.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data1 = selection1.ExecuteReader();
                data1.Read();

                float Quaterly = data1.GetInt32(0);


                this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();

                float average = (first + Quaterly) / 2;

                this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Quaterly", average);

                label2.Text = "Average : " + average;

                this.chart2.Series[combo].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Quaterly", data1.GetInt32(0));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            count = 2;
            try {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                var combo = comboBox1.Text;

                OleDbCommand selection = new OleDbCommand("select " + combo + " from FirstMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);
                
                selection.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection.Parameters.AddWithValue("Sec", comboBox3.Text);


                var data = selection.ExecuteReader();
                data.Read();

                float first = data.GetInt32(0);

                OleDbCommand selection1 = new OleDbCommand("select " + combo + " from Quaterly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection1.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection1.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection1.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data1 = selection1.ExecuteReader();
                data1.Read();

                float Quaterly = data1.GetInt32(0);

                OleDbCommand selection2 = new OleDbCommand("select " + combo + " from SecondMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection2.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection2.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection2.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data2 = selection2.ExecuteReader();
                data2.Read();

                float second = data2.GetInt32(0);

                this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                this.chart1.Series[comboBox1.Text].Points[2].Label = second.ToString();

                float average = (first + Quaterly + second) / 3;

                this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);

                label2.Text = "Average : " + average;

                this.chart2.Series[combo].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            count = 3;
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                var combo = comboBox1.Text;

                OleDbCommand selection = new OleDbCommand("select " + combo + " from FirstMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data = selection.ExecuteReader();
                data.Read();

                float first = data.GetInt32(0);

                OleDbCommand selection1 = new OleDbCommand("select " + combo + " from Quaterly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection1.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection1.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection1.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data1 = selection1.ExecuteReader();
                data1.Read();

                float Quaterly = data1.GetInt32(0);

                OleDbCommand selection2 = new OleDbCommand("select " + combo + " from SecondMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection2.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection2.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection2.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data2 = selection2.ExecuteReader();
                data2.Read();

                float second = data2.GetInt32(0);

                OleDbCommand selection3 = new OleDbCommand("select " + combo + " from Halfyearly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection3.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection3.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection3.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data3 = selection3.ExecuteReader();
                data3.Read();

                float Halfyearly = data3.GetInt32(0);

                this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Halfyearly", data3.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                this.chart1.Series[comboBox1.Text].Points[2].Label = second.ToString();
                this.chart1.Series[comboBox1.Text].Points[3].AxisLabel = "Halfyearly";
                this.chart1.Series[comboBox1.Text].Points[3].Label = Halfyearly.ToString();

                float average = (first + Quaterly + second + Halfyearly) / 4;

                this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Halfyearly", average);

                label2.Text = "Average : " + average;

                this.chart2.Series[combo].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Halfyearly", data3.GetInt32(0));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            count = 4;
            try {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                var combo = comboBox1.Text;

                OleDbCommand selection = new OleDbCommand("select " + combo + " from FirstMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data = selection.ExecuteReader();
                data.Read();

                float first = data.GetInt32(0);

                OleDbCommand selection1 = new OleDbCommand("select " + combo + " from Quaterly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection1.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection1.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection1.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data1 = selection1.ExecuteReader();
                data1.Read();

                float Quaterly = data1.GetInt32(0);

                OleDbCommand selection2 = new OleDbCommand("select " + combo + " from SecondMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection2.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection2.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection2.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data2 = selection2.ExecuteReader();
                data2.Read();

                float second = data2.GetInt32(0);

                OleDbCommand selection3 = new OleDbCommand("select " + combo + " from Halfyearly where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection3.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection3.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection3.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data3 = selection3.ExecuteReader();
                data3.Read();

                float Halfyearly = data3.GetInt32(0);

                OleDbCommand selection4 = new OleDbCommand("select " + combo + " from ThirdMidTerm where RollNo = ? and Standard=? and Sec=?", dbconn);

                selection4.Parameters.AddWithValue("RollNo", textBox1.Text);
                selection4.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection4.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data4 = selection4.ExecuteReader();
                data4.Read();

                float third = data4.GetInt32(0);

                this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("Halfyearly", data3.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points.AddXY("ThirdMidTerm", data4.GetInt32(0));
                this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                this.chart1.Series[comboBox1.Text].Points[2].Label = second.ToString();
                this.chart1.Series[comboBox1.Text].Points[3].AxisLabel = "Halfyearly";
                this.chart1.Series[comboBox1.Text].Points[3].Label = Halfyearly.ToString();
                this.chart1.Series[comboBox1.Text].Points[4].AxisLabel = "ThirdMidTerm";
                this.chart1.Series[comboBox1.Text].Points[4].Label = third.ToString();

                float average = (first + Quaterly + second + Halfyearly + third) / 5;

                this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);
                this.chart1.Series["Average"].Points.AddXY("Halfyearly", average);
                this.chart1.Series["Average"].Points.AddXY("ThirdMidTerm", average);

                label2.Text = "Average : " + average;

                this.chart2.Series[combo].Points.AddXY("FirstMidTerm", data.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Quaterly", data1.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("SecondMidTerm", data2.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("Halfyearly", data3.GetInt32(0));
                this.chart2.Series[combo].Points.AddXY("ThirdMidTerm", data4.GetInt32(0));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

            try
            {
                string combo = null;

                if (count == 1)
                {
                    combo = "Quaterly";
                }
                else if (count == 2)
                {
                    combo = "SecondMidTerm";
                }
                else if (count == 3)
                {
                    combo = "Halfyearly";
                }
                else if (count == 4)
                {
                    combo = "ThirdMidTerm";
                }
                else if (count == 5)
                {
                    combo = "Annual";
                }


                Document doc = new Document(iTextSharp.text.PageSize.A4, 30, 30, 42, 35);
                PdfWriter pdf = PdfWriter.GetInstance(doc, new FileStream("RollNo_" + textBox1.Text + "_analysis_on_" + comboBox1.Text + "_upto_" + combo + "_"+comboBox2.Text+"_"+comboBox3.Text+".pdf", FileMode.Create));
                doc.Open();
                Paragraph paragraph = new Paragraph("                                 STUDENT ANALYSIS PICTORIAL REPRESENTATION \nBAR GRAPH:\n        This graph shows the analysis of the student whose RollNo is " + textBox1.Text + " on " + comboBox1.Text + " up to " + combo + " Exam and it is also gives the average up to current exam \n" + label2.Text + "\n\n\n");
                doc.Add(paragraph);
                var chartimage = new MemoryStream();
                chart1.SaveImage(chartimage, ChartImageFormat.Png);
                iTextSharp.text.Image chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                doc.Add(chart_image);
                Paragraph paragraph1 = new Paragraph("\n\nPIE CHART:\n This graph shows the analysis of the student whose RollNo is " + textBox1.Text + " on " + comboBox1.Text + " up to " + combo + " Exam \n\n\n");
                doc.Add(paragraph1);
                var chartimage1 = new MemoryStream();
                chart2.SaveImage(chartimage1, ChartImageFormat.Png);
                iTextSharp.text.Image chart_image1 = iTextSharp.text.Image.GetInstance(chartimage1.GetBuffer());
                doc.Add(chart_image1);
                doc.Close();
                System.Diagnostics.Process.Start("RollNo_" + textBox1.Text + "_analysis_on_" + comboBox1.Text + "_upto_" + combo + "_" + comboBox2.Text + "_" + comboBox3.Text + ".pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Email email = new Email();
            email.Show();
        }

        private void label6_Click(object sender, EventArgs e)
        {
            StudentComparition sd = new StudentComparition();
            sd.Show();
        }
    }
}
