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
using System.Drawing.Imaging;
using HtmlAgilityPack;

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

            //sWrite.WriteLine("<p><img src='data:image/jpeg;base64," + Base64Encoded(Resource.Image) + "' height='10%' width='5%' > </p>");


            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); 

            string html = ReadResource("GINtool.Resources.user_manual.htm"); 
            doc.LoadHtml(html);

            var urls = doc.DocumentNode.Descendants("img")
                                .Select(e => e.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            //var str = doc.DocumentNode.InnerText;
            //doc.DocumentNode.

            foreach(string _s in urls)
            {
                var assembly = Assembly.GetExecutingAssembly();
                string _orig = string.Format("GINtool.Resources.{0}",_s);


                //string gif = ReadResource(_orig);
                Bitmap image = new Bitmap(assembly.GetManifestResourceStream(_orig));
                string img = Base64Encoded(image);

                string _nwe = string.Format("src='data:image/jpeg;base64, {0}'", img);
                string rep = string.Format("src=\"{0}\"",_s);
                
                html = html.Replace(rep,_nwe);
            }

            webBrowser1.DocumentText = html; // ReadResource("GINtool.Resources.user_manual.htm");
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
            {
                if (stream != null)
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string result = reader.ReadToEnd();
                        return result;
                    }
            }
            return "";
                       
        }


        public static String Base64Encoded(Image image)
        {
            using (MemoryStream m = new MemoryStream())
            {
                image.Save(m, ImageFormat.Jpeg);
                byte[] imageBytes = m.ToArray();

                // Convert byte[] to Base64 String
                string base64String = Convert.ToBase64String(imageBytes);
                return base64String;
            }
        }
    }
}
