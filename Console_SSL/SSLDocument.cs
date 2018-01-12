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
    class SSLDocument
    {
        /* Create new document from template
         * Add AutoText to new document
         * Update Custom XML in new document
         * Save document
         * Display new document
         */
        private string template = string.Empty;
        private WordprocessingDocument wrddoc;
        private MainDocumentPart sslmdp;
        private Document ssldoc;

        private GlossaryDocumentPart gdp;
        private GlossaryDocument gd;
        private CBAutoText atx;
        private const string DOC_PATH_NAME = @"D:\Dev Projects\SSL\Documents\SSL_Doc.docx";

        public SSLDocument() { }
        public SSLDocument(string templatefullname)
        {
            template = templatefullname;
            ssldoc = NewDocument();
        }

        private Document NewDocument()
        {
            WordprocessingDocument newdoc = WordprocessingDocument.CreateFromTemplate(template);
            wrddoc = newdoc;
            sslmdp = newdoc.MainDocumentPart;
            return newdoc.MainDocumentPart.Document;
        }

        public void BuildDocument(string[] AutoTextName)
        {

            /* Here's where I look in the Glossary Document Part to determine it there IS AutoText
             * If there isn't, then there's no use in creating the List of AutoText objects!
             */
            gdp = sslmdp.GlossaryDocumentPart;
            if (gdp != null)
            {
                gd = gdp.GlossaryDocument;
                foreach (string atxname in AutoTextName)
                {
                    atx = new CBAutoText(atxname, gd.DocParts);
                    InsertAutoText();
                }
            }
            wrddoc.SaveAs(DOC_PATH_NAME);
        }

        public void InsertAutoText()
        {
            // When doing the signature content contrld this dies!
            var cctrl = (from sdtCtrl in sslmdp.Document.Descendants<SdtElement>()
                         where (sdtCtrl.Descendants<Tag>().First().Val.ToString() == atx.Category)
                         || (sdtCtrl.Descendants<SdtAlias>().First().Val.ToString() == atx.Name)
                         select sdtCtrl).Single();

            Console.WriteLine();
            Console.WriteLine("atxt.AutoTextContent InnerText: {0}", atx.Content);
            cctrl.InnerXml = atx.Content;
        }

        class CBAutoText
        {
            private string name = string.Empty;
            private string category = string.Empty;
            private string content = string.Empty;

            private string containername = string.Empty;

            public CBAutoText(string autotextname, DocParts gdps)
            {
                var atx = (from dp in gdps
                            where dp.Descendants<DocPartProperties>().First().DocPartName.Val == autotextname
                           select dp).Single();

                name = autotextname;
                content = atx.First().InnerText;
                category = atx.Descendants<DocPartProperties>().First().Category.Name.Val;
            }

            public string Name { get { return name; } }
            public string Category { get { return category; } }
            public string Content { get { return content; } }
        }

        class CBData { }
    }
}
