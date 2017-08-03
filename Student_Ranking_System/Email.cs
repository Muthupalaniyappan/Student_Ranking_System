using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;

namespace Student_Ranking_System
{
    public partial class Email : Form
    {
        public Email()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Multiselect = true;
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = openFileDialog1.FileName.ToString();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(textBox1.Text);
                message.To.Add(textBox3.Text);
                message.Body = textBox6.Text;
                message.Subject = textBox5.Text;
                client.UseDefaultCredentials = false;
                client.EnableSsl = true;

                if (textBox4.Text != "")
                {
                    message.Attachments.Add(new Attachment(textBox4.Text));
                }

                client.Credentials = new System.Net.NetworkCredential(textBox1.Text, textBox2.Text);
                client.Send(message);

                message = null;
                MessageBox.Show("Mail sent", "Information", MessageBoxButtons.OK);
            }
            catch(Exception ex)
            {
                MessageBox.Show(""+ex,"Information",MessageBoxButtons.OK);
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
