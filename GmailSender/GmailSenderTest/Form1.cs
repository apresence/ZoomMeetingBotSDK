using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GmailSenderTest
{
    public partial class Form1 : Form
    {
        GmailSenderLib.GmailSender gmailSender = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            if (gmailSender is null)
            {
                gmailSender = new GmailSenderLib.GmailSender(Form1.ActiveForm.Text);
            }

            string msgId = gmailSender.Send(new GmailSenderLib.SimpleMailMessage(textBoxSubject.Text, textBoxBody.Text, textBoxTo.Text, null, checkBoxIsHTML.Checked));

            MessageBox.Show("Email sent; Id: " + msgId);
        }
    }
}
