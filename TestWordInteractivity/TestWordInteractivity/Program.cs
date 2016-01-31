using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;


namespace TestWordInteractivity
{
    class Program
    {
        static void Main(string[] args)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.IndentChars = ("    ");
            settings.CloseOutput = true;
            settings.OmitXmlDeclaration = true;
            using (XmlWriter thisWriter = XmlWriter.Create(@"C:\Users\patri\Source\Repos\charp_word\TestWordInteractivity\dummyOptions.xml",settings))
            {
                

                thisWriter.WriteStartElement("ListBox1");
                for (int i = 0; i < 2000; i++)
                {
                    string thisVal = "Option" + i.ToString();

                    thisWriter.WriteElementString("value", thisVal);

                }
                thisWriter.WriteEndElement();
                thisWriter.Flush();
            }
        }
    }
}
