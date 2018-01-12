﻿using System;
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
            const string TEMPLATE_PATH_NAME = @"D:\Dev Projects\SSL\Documents\Sole Source Letter v4.dotx";
            /* Create new document from template */
            SSLDocument sslnewdoc = new SSLDocument(TEMPLATE_PATH_NAME);
            sslnewdoc.BuildDocument(new string[] {"SSL-HED", "David C Meade Jr" });
            Console.WriteLine("Finished");
            Console.ReadLine();
        }
    }
}
