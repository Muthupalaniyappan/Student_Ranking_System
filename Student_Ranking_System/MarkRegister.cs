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


namespace Student_Ranking_System
{
    public partial class MarkRegister : Form
    {
        public MarkRegister()
        {

            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.FormClosed += MarkRegister_FormClosed;
        }

        private void MarkRegister_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void MarkRegister_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                string db = comboBox1.Text;

                OleDbCommand selection = new OleDbCommand("select Student_name,RollNo,Tamil,English,Maths,Physics,Chemistry,CsORBio,Total,Failcount,Rank from " + db + " where Standard=? and Sec=? ", dbconn);
                selection.Parameters.AddWithValue("Standard", comboBox2.Text);
                selection.Parameters.AddWithValue("Sec", comboBox3.Text);

                OleDbDataAdapter da = new OleDbDataAdapter(selection);

                DataTable dt = new DataTable();

                da.Fill(dt);

                dataGridView1.DataSource = dt;

                dataGridView1.Sort(dataGridView1.Columns[10], ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error); ;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = new Document(iTextSharp.text.PageSize.A4.Rotate(), 10, 10, 42, 35);
                PdfWriter pdf = PdfWriter.GetInstance(doc, new FileStream("ConsolidatedMarkSheet_for_Standard" + comboBox2.Text +"_" +comboBox3.Text+"_section for_"+ comboBox1.Text + ".pdf", FileMode.Create));
                doc.Open();
                PdfPTable table = new PdfPTable(dataGridView1.Columns.Count);

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    PdfPCell cell = new PdfPCell();
                    cell.BackgroundColor = BaseColor.GRAY;
                    cell.AddElement(new Chunk(dataGridView1.Columns[i].HeaderText));
                    table.AddCell(cell);
                }
                table.SetWidths(new float[] { 3f, 1f, 1f, 1f, 1f, 1f, 1f, 1f, 1f, 1f, 1f });
                table.HeaderRows = 1;

                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    for (int k = 0; k < dataGridView1.Columns.Count; k++)
                    {
                        if (dataGridView1[k, j].Value != null)
                        {
                            table.AddCell(new Phrase(dataGridView1[k, j].Value.ToString()));
                        }
                    }
                }

                //string image = "rank.jpg";

                //iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(image);

                //jpg.Alignment = iTextSharp.text.Image.UNDERLYING;
                //float width = PageSize.A4.Width;
                //float height = PageSize.A4.Height;
                //jpg.ScaleToFit(width,height);

                //doc.Add(jpg);

                //Paragraph paragraph = new Paragraph("                                                                            Consolidated Marksheet for " + comboBox1.Text + " Exam\n");
                //doc.Add(paragraph);
                doc.Add(table);

                doc.Close();

                System.Diagnostics.Process.Start("ConsolidatedMarkSheet_for_Standard" + comboBox2.Text + "_" + comboBox3.Text + "_section for_" + comboBox1.Text + ".pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Email email = new Email();
            email.Show();
        }
    }
}
