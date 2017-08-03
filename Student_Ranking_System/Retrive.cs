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

using System.Net;
using System.Net.Cache;

using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Student_Ranking_System
{
    public partial class _12th_STD_Details : Form
    {
        int fail = 0;

        TextBox CustomTextBox = new TextBox();
        Label CustomLabel = new Label();

        public _12th_STD_Details()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            siteUrl = string.Format("http://{0}.way2sms.com", WorkingSite);
            this.FormClosed += _12th_STD_Details_FormClosed;
        }

        private void _12th_STD_Details_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try {
                OleDbConnection dbconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                dbconn.Open();

                string db = comboBox1.Text;

                OleDbCommand retrive = new OleDbCommand("select Student_name,Tamil,English,Maths,Physics,Chemistry,CsORBio,Total,Rank,RollNo,Failcount from " + db + " where RollNo = ? and Standard=? and Sec=?", dbconn);

                retrive.Parameters.AddWithValue("RollNo", textBox1.Text);
                retrive.Parameters.AddWithValue("Standard", comboBox2.Text);
                retrive.Parameters.AddWithValue("Sec", comboBox3.Text);

                var data = retrive.ExecuteReader();
                data.Read();

                fail = data.GetInt32(10);
                var r_no = "" + data.GetInt32(9);

                    label2.Text = "Student_Name";
                    label3.Text = "Tamil";
                    label4.Text = "English";
                    label5.Text = "Maths";
                    label6.Text = "Physics";
                    label7.Text = "Chemistry";
                    label8.Text = "CsORBio";
                    label9.Text = "Total";
                    label10.Text = "Rank";
                    label11.Text = "" + data.GetString(0);
                    label12.Text = "" + data.GetInt32(1);
                    label13.Text = "" + data.GetInt32(2);
                    label14.Text = "" + data.GetInt32(3);
                    label15.Text = "" + data.GetInt32(4);
                    label16.Text = "" + data.GetInt32(5);
                    label17.Text = "" + data.GetInt32(6);
                    label18.Text = "" + data.GetInt32(7);
                    label19.Text = "" + data.GetInt32(8);
                    label20.Text = "/200";
                    label21.Text = "/200";
                    label22.Text = "/200";
                    label23.Text = "/200";
                    label24.Text = "/200";
                    label25.Text = "/200";
                    label26.Text = "/1200";
         
            }
            catch(Exception ex)
            {
                label2.Text = "";
                label3.Text = "";
                label4.Text = "";
                label5.Text = "";
                label6.Text = "";
                label7.Text = "";
                label8.Text = "";
                label9.Text = "";
                label10.Text = "";
                label11.Text = "";
                label12.Text = "";
                label13.Text = "";
                label14.Text = "";
                label15.Text = "";
                label16.Text = "";
                label17.Text = "";
                label18.Text = "";
                label19.Text = "";
                label20.Text = "";
                label21.Text = "";
                label22.Text = "";
                label23.Text = "";
                label24.Text = "";
                label25.Text = "";
                label26.Text = "";


                if(comboBox1.Text == null)
                MessageBox.Show("Error " + ex,"Information", MessageBoxButtons.OK,MessageBoxIcon.Error);
                else
                MessageBox.Show("Enter a Valid RollNo", "Informtion", MessageBoxButtons.OK,MessageBoxIcon.Error);
               
            }
        }

        private void _12th_STD_Details_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Std_Selection std = new Std_Selection();
            this.Hide();
            std.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Document doc = new Document(iTextSharp.text.PageSize.A4, 0, 0, 0, 0);
            PdfWriter pdf = PdfWriter.GetInstance(doc, new FileStream("RollNo_" +textBox1.Text+"_MerkSheet for_" + comboBox1.Text +"_Exam in Standard_"+ comboBox2.Text +"_Section"+comboBox3.Text+".pdf", FileMode.Create));
            doc.Open();
            string image = "rank.jpg";

            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(image);

            jpg.Alignment = iTextSharp.text.Image.UNDERLYING;
            float width = PageSize.A4.Width;
            float height = PageSize.A4.Height;
            jpg.ScaleToFit(width, height);

            doc.Add(jpg);

            Paragraph SchoolName = new Paragraph("\n\n\n\n                                                     "+textBox2.Text);
            doc.Add(SchoolName);
            if (comboBox1.Text == "SecondMidTerm" || comboBox1.Text == "ThirdMidTerm" || comboBox1.Text == "FirstMidTerm")
            {
                Paragraph exam = new Paragraph("\n\n                                                                                 " + comboBox1.Text + " Test");
                doc.Add(exam);
            }
            else
            {
                Paragraph exam = new Paragraph("\n\n                                                                                 " + comboBox1.Text + " Exam");
                doc.Add(exam);
            }
            Paragraph name = new Paragraph("\n\n\n\n                       "+label2.Text+":"+label11.Text+"                                                                    RollNo:"+textBox1.Text);
            doc.Add(name);
            Paragraph paragraph = new Paragraph("\n\n\n\n\n                               1.                               " +label3.Text+ "                                                         " + label12.Text);
            doc.Add(paragraph);
            Paragraph paragraph1 = new Paragraph("\n                               2.                               " + label4.Text + "                                                      " + label13.Text);
            doc.Add(paragraph1);
            Paragraph paragraph2 = new Paragraph("\n                               3.                               " + label5.Text + "                                                        " + label14.Text);
            doc.Add(paragraph2);
            Paragraph paragraph3 = new Paragraph("\n                               4.                               " + label6.Text + "                                                     " + label15.Text);
            doc.Add(paragraph3);
            Paragraph paragraph4 = new Paragraph("\n                               5.                             " + label7.Text + "                                                   " + label16.Text);
            doc.Add(paragraph4);
            Paragraph paragraph5 = new Paragraph("\n                               6.              Computer Science Or Biology                                   " + label17.Text);
            doc.Add(paragraph5);
            Paragraph paragraph6 = new Paragraph("\n\n\n\n                                                                                                                     " + label18.Text + "                 1200");
            doc.Add(paragraph6);
            Paragraph paragraph7 = new Paragraph("\n\n                                                                                                                     " + label19.Text);
            doc.Add(paragraph7);

            string remar = null;

            if(fail==0)
            {
                if(label19.Text=="1")
                {
                    remar = "Very Good Keep it up!"; 
                }
                else if(label19.Text == "2" || label19.Text == "3")
                {
                    remar = "Keep it up!";
                }
                else
                {
                    remar = "Try to get more marks and get Top 3 Ranks";
                }
            }
            else if(fail == 1)
            {
                remar = "Try to get pass mark in that subject and try to get good total";
            }
            else if (fail == 2)
            {
                remar = "Try to get pass mark in these two subject and try to get good total";
            }
            else if (fail == 3)
            {
                remar = "Meet the class Teacher with your parents";
            }
            else if (fail == 4)
            {
                remar = "Poor,Meet the class Teacher with your parents";
            }
            else if (fail == 5)
            {
                remar = "Very Poor,Meet the class Teacher with your parents";
            }
            else
            {
                remar = "Very Poor,Meet the class Teacher with your parents";
            }

            Paragraph remark = new Paragraph("\n\n                                                 "+remar);
            doc.Add(remark);
            doc.Close();
            System.Diagnostics.Process.Start("RollNo_" + textBox1.Text + "_MerkSheet for_" + comboBox1.Text + "_Exam in Standard_" + comboBox2.Text + "_Section" + comboBox3.Text + ".pdf");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            btnSend_Click(sender, e);
        }

        string userName;
        string passWord;
        HttpWebRequest req;
        CookieContainer cookieCntr;
        HttpWebResponse response;
        string url;
        string TokenStr;
        string WorkingSite = "site23";
        string siteUrl = string.Empty;

        protected void btnSend_Click(object sender, EventArgs e)
        {
            if (this.textBox3.Text != "" && this.textBox4.Text != "")
            {
                try
                {

                    userName = textBox3.Text.Trim();
                    passWord = textBox4.Text.Trim();
                    //string msg = textBox3.Text.Trim();

                    OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                    conn.Open();

                    string db = comboBox1.Text;

                    OleDbCommand c = new OleDbCommand("select count(*) from " + db + " where Standard=? and Sec=?", conn);
                    c.Parameters.AddWithValue("Standard", comboBox2.Text);
                    c.Parameters.AddWithValue("Sec", comboBox3.Text);

                    string rank = "" + c.ExecuteScalar();

                    OleDbCommand ms = new OleDbCommand("select Tamil,English,Maths,Physics,Chemistry,CsORBio,Total,Rank,Failcount from " + db + " where RollNo = ? and Standard=? and Sec=?", conn);
                    ms.Parameters.AddWithValue("RollNo", textBox1.Text);
                    ms.Parameters.AddWithValue("Standard", comboBox2.Text);
                    ms.Parameters.AddWithValue("Sec", comboBox3.Text);

                    var data = ms.ExecuteReader();
                    data.Read();

                    

                    OleDbCommand no = new OleDbCommand("select PhoneNumber from SDetails where RollNo=? and Standard=? and Sec=?", conn);
                    no.Parameters.AddWithValue("RollNo", textBox1.Text);
                    no.Parameters.AddWithValue("Standard", comboBox2.Text);
                    no.Parameters.AddWithValue("Sec", comboBox3.Text);

                    var d = no.ExecuteReader();
                    d.Read();

                    string msg = "Hello Sir/Madam,Your child marks Tam:" + data.GetInt32(0) + ",Eng:" + data.GetInt32(1) + ",Mat:" + data.GetInt32(2) + ",Phy:" + data.GetInt32(3) + ",Che:" + data.GetInt32(4) + ",Cs/Bio:" + data.GetInt32(5) + " and the total of " + data.GetInt32(6) + "/1200 with Rank:" + data.GetInt32(7) + "/" + rank + ". No.Fail:" + data.GetInt32(8);
                    //string msg = "Hello Sir/Madam,Your child marks Tam:200,Eng;200,Mat:200,Phy:200,Che:200,Cs/Bio:200 and the total of 1200/1200 with Rank:120/120.No.Fail:6";

                    string pno = "" + d.GetDecimal(0);

                    if (Connect() == true)
                    {
                        if (checkBox1.Checked == false)
                        {
                            SendMessage(pno, msg);
                            MessageBox.Show("Message sent");
                        }
                        else
                        {
                            SendMessage(pno, CustomTextBox.Text);
                            MessageBox.Show("Message sent");
                        }
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("" + ex, "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }
        bool Connect()
        {
            try
            {
                url = string.Empty;
                this.req = (HttpWebRequest)WebRequest.Create(string.Format("{0}/Login1.action", siteUrl));
                this.req = SetHeaders(this.req, false, string.Format("{0}/content/index.html", siteUrl), true);

                url = string.Format("username={0}&password={1}&userLogin=yes", userName, passWord);
                this.req.ContentLength = url.Length;

                StreamWriter writer = new StreamWriter(this.req.GetRequestStream(), System.Text.Encoding.ASCII);
                writer.Write(url);
                writer.Close();
                this.response = (HttpWebResponse)this.req.GetResponse();
                this.cookieCntr = this.req.CookieContainer;
                this.response.Close();

                string TokenStrtmp1 = response.ResponseUri.Query.Split(new string[] { "Token=" }, StringSplitOptions.None)[1];
                this.TokenStr = TokenStrtmp1.Split(new string[] { "&" }, StringSplitOptions.None)[0];
                return true;

            }
            catch (Exception ex) { MessageBox.Show("" + ex); return false; }
        }

        HttpWebRequest SetHeaders(HttpWebRequest request, bool useEncoding, string referer, bool isNewCookie)
        {
            request.Method = "POST";
            request.KeepAlive = true;
            HttpRequestCachePolicy cachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.Refresh);
            request.CachePolicy = cachePolicy;
            request.Headers.Add("Origin", siteUrl);
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.4 (KHTML, like Gecko) Chrome/22.0.1229.79 Safari/537.4";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            if (!string.IsNullOrEmpty(referer))
                request.Referer = referer;
            if (useEncoding)
                request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Headers.Add("Accept-Language", "en-US,en;q=0.8");
            request.Headers.Add("Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.3");
            request.AllowAutoRedirect = true;
            if (isNewCookie)
                request.CookieContainer = new CookieContainer();
            else
                request.CookieContainer = this.cookieCntr;
            return request;
        }

        bool SendMessage(string mobileNo, string Message)
        {
            try
            {

                this.req = (HttpWebRequest)WebRequest.Create(string.Format("{0}/smstoss.action", siteUrl));
                this.req = SetHeaders(this.req, false, string.Format("{0}/sendSMS?Token={1}", siteUrl, this.TokenStr), false);
                int msgLen = 140 - Encoding.ASCII.GetBytes(Message).Length;
                if (msgLen < 0)
                    throw new NotImplementedException();
                this.url = string.Format("ssaction=ss&Token={0}&mobile={1}&message={2}&msgLen={3}", TokenStr, mobileNo, Uri.EscapeDataString(Message).Replace("%20", "+"), msgLen);
                this.req.ContentLength = this.url.Length;
                StreamWriter writer = new StreamWriter(this.req.GetRequestStream(), System.Text.Encoding.ASCII);
                writer.Write(this.url);
                writer.Close();
                this.response = (HttpWebResponse)this.req.GetResponse();
                string ss = new StreamReader(this.response.GetResponseStream()).ReadToEnd();
                this.response.Close();
                return true;
            }
            catch { return false; }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                CustomTextBox.Size = new Size(322, 82);
                CustomTextBox.Location = new Point(830, 591);
                CustomTextBox.Multiline = true;
                CustomTextBox.MaxLength = 140;
                CustomTextBox.Anchor = AnchorStyles.Top;
                this.Controls.Add(CustomTextBox);

                CustomLabel.Size = new Size(52, 16);
                CustomLabel.Location = new Point(1105, 565);
                CustomLabel.Anchor = AnchorStyles.Top;
                CustomLabel.Text = "0/140";
                this.Controls.Add(CustomLabel);

                CustomTextBox.TextChanged += Textcount;
            }
            if (checkBox1.Checked == false)
            {
                this.Controls.Remove(CustomLabel);
                this.Controls.Remove(CustomTextBox);
            }
        }

        private void Textcount(object sender, EventArgs e)
        {
            int textcount = CustomTextBox.Text.Length;
            CustomLabel.Text = textcount.ToString() + "/" + CustomTextBox.MaxLength.ToString();

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton1.Checked==true)
            {
                BulkMessage bm = new BulkMessage(textBox1.Text, comboBox1.Text, comboBox2.Text, comboBox3.Text);
                bm.Show();
                radioButton1.Checked = false;
            }
        }
    }
}