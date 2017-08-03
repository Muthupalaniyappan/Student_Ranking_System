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
    public partial class SubjectAnalysis : Form
    {
        public SubjectAnalysis()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += SubjectAnalysis_FormClosed;
        }

        private void SubjectAnalysis_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var combo1 = comboBox1.Text;
            var combo2 = comboBox2.Text;
            float first = 0;
            float Quaterly = 0;
            float Second = 0;
            float Halfyearly = 0;
            float Third = 0;
            float Annual = 0;
            var first1 = 0;
            var Quaterly1 = 0;
            var Second1 = 0;
            var Halfyearly1 = 0;
            var Third1 = 0;
            var Annual1 = 0;

            int c = 0;
            try
            {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                OleDbCommand selection = new OleDbCommand("select " + combo1 + " from FirstMidTerm where Standard=? and Sec=?", dbconn);
                selection.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection.Parameters.AddWithValue("Sec", comboBox4.Text);

                var data = selection.ExecuteReader();

                while(data.Read()!=false)
                {
                    if (data.GetInt32(0) != 0)
                    {
                        first1 = first1 + data.GetInt32(0);
                        c++;
                    }
                }

                first = first1 / c;

                c = 0;

                OleDbCommand selection1 = new OleDbCommand("select " + combo1 + " from Quaterly where Standard=? and Sec=?", dbconn);
                selection1.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection1.Parameters.AddWithValue("Sec", comboBox4.Text);

                var data1 = selection1.ExecuteReader();
                data1.Read();

                while (data1.Read() != false)
                {
                    if (data1.GetInt32(0) != 0)
                    {
                        Quaterly1 = Quaterly1 + data1.GetInt32(0);
                        c++;
                    }
                }

                Quaterly = Quaterly1 / c;

                c = 0;

                OleDbCommand selection2 = new OleDbCommand("select " + combo1 + " from SecondMidTerm where Standard=? and Sec=?", dbconn);
                selection2.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection2.Parameters.AddWithValue("Sec", comboBox4.Text);
                var data2 = selection2.ExecuteReader();

                while (data2.Read() != false)
                {
                    if (data2.GetInt32(0) != 0)
                    {
                        Second1 = Second1 + data2.GetInt32(0);
                        c++;
                    }
                }

                Second = Second1 / c;

                c = 0;

                OleDbCommand selection3 = new OleDbCommand("select " + combo1 + " from Halfyearly where Standard=? and Sec=?", dbconn);
                selection3.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection3.Parameters.AddWithValue("Sec", comboBox4.Text);
                var data3 = selection3.ExecuteReader();

                while (data3.Read() != false)
                {
                    if (data3.GetInt32(0) != 0)
                    {
                        Halfyearly1 = Halfyearly1 + data3.GetInt32(0);
                        c++;
                    }
                }

                Halfyearly = Halfyearly1 / c;

                c = 0;

                OleDbCommand selection4 = new OleDbCommand("select " + combo1 + " from ThirdMidTerm where Standard=? and Sec=?", dbconn);
                selection4.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection4.Parameters.AddWithValue("Sec", comboBox4.Text);
                var data4 = selection4.ExecuteReader();
                while (data4.Read() != false)
                {
                    if (data4.GetInt32(0) != 0)
                    {
                        Third1 = Third1 + data4.GetInt32(0);
                        c++;
                    }
                }

                Third = Third1 / c;

                c = 0;

                OleDbCommand selection5 = new OleDbCommand("select " + combo1 + " from Annual where Standard=? and Sec=?", dbconn);
                selection5.Parameters.AddWithValue("Standard", comboBox3.Text);
                selection5.Parameters.AddWithValue("Sec", comboBox4.Text);

                var data5 = selection5.ExecuteReader();
                while (data5.Read() != false)
                {
                    if (data5.GetInt32(0) != 0)
                    {
                        Annual1 = Annual1 + data5.GetInt32(0);
                        c++;
                    }
                }


                Annual = Annual1 / c;
                c = 0;
               
                if (combo2 == "Quaterly")
                {
                    this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", first);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", Quaterly);
                    this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                    this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                    this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();

                    float average = (first + Quaterly) / 2;

                    label3.Text = "Average : " + average;

                    this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Quaterly", average);

                    this.chart2.Series[combo1].Points.AddXY("FirstMidTerm", first);
                    this.chart2.Series[combo1].Points.AddXY("Quaterly", Quaterly);
                }
                else if(combo2 == "SecondMidTerm")
                {
                    this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", first);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", Quaterly);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", Second);
                    //this.chart1.Series[comboBox1.Text].Points.Add(first);
                    this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                    this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                    this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                    this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[2].Label = Second.ToString();
                    

                    float average = (first + Quaterly + Second) / 3;

                    this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                    this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);

                    label3.Text = "Average : " + average;

                    this.chart2.Series[combo1].Points.AddXY("FirstMidTerm", first);
                    this.chart2.Series[combo1].Points.AddXY("Quaterly", Quaterly);
                    this.chart2.Series[combo1].Points.AddXY("SecondMidTerm", Second);
                    
                }
                else if(combo2 == "Halfyearly")
                {
                    this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", first);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", Quaterly);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", Second);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Halfyearly", Halfyearly);
                    this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                    this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                    this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                    this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[2].Label = Second.ToString();
                    this.chart1.Series[comboBox1.Text].Points[3].AxisLabel = "Halfyearly";
                    this.chart1.Series[comboBox1.Text].Points[3].Label = Halfyearly.ToString();


                    float average = (first + Second + Quaterly + Halfyearly) / 4;

                    this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                    this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Halfyearly", average);

                    label3.Text = "Average : " + average;

                    this.chart2.Series[combo1].Points.AddXY("FirstMidTerm", first);
                    this.chart2.Series[combo1].Points.AddXY("Quaterly", Quaterly);
                    this.chart2.Series[combo1].Points.AddXY("SecondMidTerm", Second);
                    this.chart2.Series[combo1].Points.AddXY("Halfyearly", Halfyearly);
                }
                else if(combo2 == "ThirdMidTerm")
                {

                    this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", first);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", Quaterly);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", Second);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Halfyearly", Halfyearly);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("ThirdMidTerm", Third);
                    this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                    this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                    this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                    this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[2].Label = Second.ToString();
                    this.chart1.Series[comboBox1.Text].Points[3].AxisLabel = "Halfyearly";
                    this.chart1.Series[comboBox1.Text].Points[3].Label = Halfyearly.ToString();
                    this.chart1.Series[comboBox1.Text].Points[4].AxisLabel = "ThirdMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[4].Label = Third.ToString();

                    float average = (first + Quaterly + Second + Halfyearly + Third) / 5;

                    this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                    this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Halfyearly", average);
                    this.chart1.Series["Average"].Points.AddXY("ThirdMidTerm", average);

                    label3.Text = "Average : " + average;

                    this.chart2.Series[combo1].Points.AddXY("FirstMidTerm", first);
                    this.chart2.Series[combo1].Points.AddXY("Quaterly", Quaterly);
                    this.chart2.Series[combo1].Points.AddXY("SecondMidTerm", Second);
                    this.chart2.Series[combo1].Points.AddXY("Halfyearly", Halfyearly);
                    this.chart2.Series[combo1].Points.AddXY("ThirdMidTerm", Third);
                }
                else if(combo2 == "Annual")
                {
                    this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", first);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Quaterly", Quaterly);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("SecondMidTerm", Second);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Halfyearly", Halfyearly);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("ThirdMidTerm", Third);
                    this.chart1.Series[comboBox1.Text].Points.AddXY("Annual", Annual);
                    this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                    this.chart1.Series[comboBox1.Text].Points[1].AxisLabel = "Quaterly";
                    this.chart1.Series[comboBox1.Text].Points[1].Label = Quaterly.ToString();
                    this.chart1.Series[comboBox1.Text].Points[2].AxisLabel = "SecondMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[2].Label = Second.ToString();
                    this.chart1.Series[comboBox1.Text].Points[3].AxisLabel = "Halfyearly";
                    this.chart1.Series[comboBox1.Text].Points[3].Label = Halfyearly.ToString();
                    this.chart1.Series[comboBox1.Text].Points[4].AxisLabel = "ThirdMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[4].Label = Third.ToString();
                    this.chart1.Series[comboBox1.Text].Points[5].AxisLabel = "Annual";
                    this.chart1.Series[comboBox1.Text].Points[5].Label = Annual.ToString();

                    float average = (first + Quaterly + Second + Halfyearly + Third + Annual) / 6;

                    this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Quaterly", average);
                    this.chart1.Series["Average"].Points.AddXY("SecondMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Halfyearly", average);
                    this.chart1.Series["Average"].Points.AddXY("ThirdMidTerm", average);
                    this.chart1.Series["Average"].Points.AddXY("Annual", average);

                    label3.Text = "Average : " + average;

                    this.chart2.Series[combo1].Points.AddXY("FirstMidTerm", first);
                    this.chart2.Series[combo1].Points.AddXY("Quaterly", Quaterly);
                    this.chart2.Series[combo1].Points.AddXY("SecondMidTerm", Second);
                    this.chart2.Series[combo1].Points.AddXY("Halfyearly", Halfyearly);
                    this.chart2.Series[combo1].Points.AddXY("ThirdMidTerm", Third);
                    this.chart2.Series[combo1].Points.AddXY("Annual", Annual);

                }
                else
                {
                    this.chart1.Series[comboBox1.Text].Points.AddXY("FirstMidTerm", first);
                    this.chart1.Series[comboBox1.Text].Points[0].AxisLabel = "FirstMidTerm";
                    this.chart1.Series[comboBox1.Text].Points[0].Label = first.ToString();
                    float average = first;
                    label3.Text = "Average : " + average;
                    this.chart1.Series["Average"].Points.AddXY("FirstMidTerm", average);
                    this.chart2.Series[combo1].Points.AddXY("FirstMidTerm", first);
                }


            }
            catch(Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                this.chart1.Series[comboBox1.Text].Points.Clear();
                this.chart1.Series["Average"].Points.Clear();
                this.chart2.Series[comboBox1.Text].Points.Clear();
                label3.Text = "";
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex, "Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void SubjectAnalysis_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = new Document(iTextSharp.text.PageSize.A4, 30, 30, 42, 35);
                PdfWriter pdf = PdfWriter.GetInstance(doc, new FileStream("subjectanalysis_" + comboBox1.Text + "_upto_" + comboBox2.Text + "for_"+comboBox3.Text+"_"+comboBox4.Text+"_section.pdf", FileMode.Create));
                doc.Open();
                Paragraph paragraph = new Paragraph("                                 SUBJECT ANALYSIS PICTORIAL REPRESENTATION \nBAR GRAPH:\n        This graph shows the analysis on the subject " + comboBox1.Text + " up to " + comboBox2.Text + " Exam and it is also gives the average of the particular subject up to current exam \n" + label3.Text + "\n\n\n");
                doc.Add(paragraph);
                var chartimage = new MemoryStream();
                chart1.SaveImage(chartimage, ChartImageFormat.Png);
                iTextSharp.text.Image chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                doc.Add(chart_image);
                Paragraph paragraph1 = new Paragraph("\n\nPIE CHART:\n This graph shows the analysis on the subject " + comboBox1.Text + " up to " + comboBox2.Text + " Exam \n\n\n");
                doc.Add(paragraph1);
                var chartimage1 = new MemoryStream();
                chart2.SaveImage(chartimage1, ChartImageFormat.Png);
                iTextSharp.text.Image chart_image1 = iTextSharp.text.Image.GetInstance(chartimage1.GetBuffer());
                doc.Add(chart_image1);
                doc.Close();
                System.Diagnostics.Process.Start("subjectanalysis_" + comboBox1.Text + "_upto_" + comboBox2.Text + "for_" + comboBox3.Text + "_" + comboBox4.Text + "_section.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Email email = new Email();
            email.Show();
        }
    }
}
