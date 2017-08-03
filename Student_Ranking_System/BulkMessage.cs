using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Net.Cache;


namespace Student_Ranking_System
{
    public partial class BulkMessage : Form
    {
        TextBox CustomTextBox = new TextBox();
        Label CustomLabel = new Label();

        string combo1;
        string combo2;
        string combo3;
        string text1;

        public BulkMessage(string t,string c1,string c2,string c3)
        {
            InitializeComponent();
            combo1 = c1;
            combo2 = c2;
            combo3 = c3;
            text1 = t;
            siteUrl = string.Format("http://{0}.way2sms.com", WorkingSite);
        }

        private void button1_Click(object sender, EventArgs e)
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
            if (this.textBox1.Text != "" && this.textBox2.Text != "")
            {
                try
                {
                    int count = 0;
                    userName = textBox1.Text.Trim();
                    passWord = textBox2.Text.Trim();
                    //string msg = textBox3.Text.Trim();

                    OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Student_Ranking_System.accdb");
                    conn.Open();

                    string db = combo1;

                    OleDbCommand c = new OleDbCommand("select count(*) from " + db + " where Standard=? and Sec=?", conn);
                    c.Parameters.AddWithValue("Standard", combo2);
                    c.Parameters.AddWithValue("Sec", combo3);

                    string rank = "" + c.ExecuteScalar();

                    OleDbCommand ms = new OleDbCommand("select Tamil,English,Maths,Physics,Chemistry,CsORBio,Total,Rank,Failcount,RollNo from " + db + " where Standard=? and Sec=?", conn);
                    ms.Parameters.AddWithValue("Standard", combo2);
                    ms.Parameters.AddWithValue("Sec", combo3);

                    var data = ms.ExecuteReader();


                    while (data.Read() != false)
                    {
                        OleDbCommand no = new OleDbCommand("select PhoneNumber from SDetails where RollNo=? and Standard=? and Sec=?", conn);
                        no.Parameters.AddWithValue("RollNo", data.GetInt32(9));
                        no.Parameters.AddWithValue("Standard", combo2);
                        no.Parameters.AddWithValue("Sec", combo3);

                        var d = no.ExecuteReader();
                        d.Read();

                        string msg = "Hello Sir/Madam,Your child marks Tam:" + data.GetInt32(0) + ",Eng:" + data.GetInt32(1) + ",Mat:" + data.GetInt32(2) + ",Phy:" + data.GetInt32(3) + ",Che:" + data.GetInt32(4) + ",Cs/Bio:" + data.GetInt32(5) + " and the total of " + data.GetInt32(6) + "/1200 with Rank:" + data.GetInt32(7) + "/" + rank + ". No.Fail:" + data.GetInt32(8);
                        //string msg = "Hello Sir/Madam,Your child marks Tam:200,Eng;200,Mat:200,Phy:200,Che:200,Cs/Bio:200 and the total of 1200/1200 with Rank:120/120.No.Fail:6";

                        string pno = "" + d.GetDecimal(0);

                        if (pno != null)
                        {
                            if (Connect() == true)
                            {
                                if (checkBox1.Checked == false)
                                {
                                    SendMessage(pno, msg);
                                    //MessageBox.Show("Message sent");
                                    count++;
                                }
                                else
                                {
                                    SendMessage(pno, CustomTextBox.Text);
                                    //MessageBox.Show("Message sent");
                                    count++;
                                }
                            }
                        }
                    }

                    MessageBox.Show("Message sent for " + count + " students", "Inforammtion", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                catch (Exception ex)
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
                CustomTextBox.Size = new Size(287, 83);
                CustomTextBox.Location = new Point(285, 84);
                CustomTextBox.Multiline = true;
                CustomTextBox.MaxLength = 140;
                //CustomTextBox.Anchor = AnchorStyles.Top;
                this.Controls.Add(CustomTextBox);

                CustomLabel.Size = new Size(45, 16);
                CustomLabel.Location = new Point(527, 62);
                //CustomLabel.Anchor = AnchorStyles.Top;
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

        private void label4_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
