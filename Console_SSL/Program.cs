using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using wrd10 = DocumentFormat.OpenXml.Office2010.Word;
using wrd13 = DocumentFormat.OpenXml.Office2013.Word;

namespace Console_SSL
{
    class Program
    {
        static void Main(string[] args)
        {
            //const string TEMPLATE_PATH_NAME = @"D:\Dev Projects\SSL\Documents\Sole Source Letter v4.dotx";
            const string TEMPLATE_PATH_NAME = @"C:\Users\ajones\Documents\Automation\Code\Word\SSL Work\Sole Source Letter v4.dotx";
            /* Create new document from template */
            //SSLDocument sslnewdoc = new SSLDocument(TEMPLATE_PATH_NAME);
            CMDocument newdoc = new CMDocument(TEMPLATE_PATH_NAME);
            /* OLD AutoText Choices
             * Jeremy Singer
             * Auditi Chakravarty
             * Trevor Packer
             * Cyndie Schmeiser
             * Trevor Packer
             * SSL-K12
             * SSL-HED
             */

            /* NEW AutoText Choices
             * TP - Category: Signatures
             * AC - Category: Signatures
             * JS - Category: Signatures
             * DMJ - Category: Signatures
             * HED - Category: Signatures
             * K12 - Category: Signatures
             * ***************************
             * For the Future
             * TVP - Category: Signatures
             * ADC - Category: Signatures
             * JYS - Category: Signatures
             * DMJ - Category: Signatures
             * ***************************
            */
            //sslnewdoc.BuildDocument(new string[] {"SSL-K12", "Cyndie Schmeiser" });
            newdoc.BuildDocument(new string[] { "K12", "DMJ" });
            Console.WriteLine("Finished");
            Console.ReadLine();
        }
    }
}
