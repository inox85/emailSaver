
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Pop3;

namespace MailDownloader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a POP3 client  
            Pop3Client pop = new Pop3Client();
            //Set host, username, password etc. for the client  
            pop.Host = "pop.chemitec.it";
            pop.Username = "d.innocenti@chemitec.it";
            pop.Password = "@Verde6100";
            pop.Port = 995;
            pop.EnableSsl = true;
            
            //Connect the server  
            pop.Connect();
            //Get the first message by its sequence number  
            Pop3MessageInfoCollection m = pop.GetAllMessages();
            MailMessage message = pop.GetMessage(54);
            //Parse the message 
            message.Save("Sample.msg", MailMessageFormat.Msg);

            Console.WriteLine("------------------ HEADERS ---------------");
            Console.WriteLine("From : " + message.From.ToString());
            Console.WriteLine("To : " + message.To.ToString());
            Console.WriteLine("Date : " + message.Date.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Subject: " + message.Subject);
            Console.WriteLine("------------------- BODY -----------------");
            Console.WriteLine(message.BodyText);
            Console.WriteLine("------------------- END ------------------");
            //Save the message to disk using its subject as file name  
            message.Save(message.Subject + ".eml", MailMessageFormat.Eml);
            Console.WriteLine("Message Saved.");
        }
    }
}
