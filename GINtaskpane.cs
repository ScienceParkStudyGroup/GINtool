using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace GINtool
{
    public partial class GINtaskpane : UserControl
    {

        public delegate void UpdateButtonStatus(bool visible);
        public UpdateButtonStatus updateButtonStatus = null;

        public GINtaskpane()
        {
            InitializeComponent();

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

            string html = ReadResource("GINtool.Resources.user_manual.htm");
            doc.LoadHtml(html);

            // parse document for images and return the src value
            var urls = doc.DocumentNode.Descendants("img")
                                .Select(e => e.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            // for each src entry found, read it from resources and display it as inline base64 encoded text
            foreach (string _s in urls)
            {
                var assembly = Assembly.GetExecutingAssembly();
                string _orig = string.Format("GINtool.Resources.{0}", _s);

                Bitmap image = new Bitmap(assembly.GetManifestResourceStream(_orig));
                string img = Base64Encoded(image);

                string _nwe = string.Format("src='data:image/jpeg;base64, {0}'", img);
                string rep = string.Format("src=\"{0}\"", _s);

                html = html.Replace(rep, _nwe);
            }

            // show text in webbrowser component
            webBrowser1.DocumentText = html;
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
