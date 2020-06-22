using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using GINtool.Properties;
using System.Reflection;

namespace GINtool
{
    public partial class GINtaskpane : UserControl
    {
        
        public GINtaskpane()
        {
            InitializeComponent();

            ///Properties.Settings.Default.htmlHelp;

            //string yourRTFText = @"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard This is some {\b bold} text.\par }"; 
            //MemoryStream stream = new MemoryStream(UTF8Encoding.Default.GetBytes(yourRTFText));
            //richTextBox1.Text = stream.ToString();            

           // webView1.Navigate("http://www.contoso.com");

            //webBrowser1.DocumentText = ReadResource("GINtool.Resources.user_manual.htm");
        }

        public void LoadHtml()
        {
            //string html = Properties.Settings.Default.htmlHelp;
          
        }

        public string ReadResource(string resourceName)
        {
            // Determine path
            var assembly = Assembly.GetExecutingAssembly();            

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                string result = reader.ReadToEnd();
                return result;
            }

            return "";
            
        }
    }
}
