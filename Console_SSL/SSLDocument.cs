using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

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
        private const string DOC_PATH_NAME = @"D:\Dev Projects\SSL\Documents\SSL_Doc.docx";

        private GlossaryDocumentPart gdp;
        private GlossaryDocument gd;
        private CBAutoText atx;

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
            return sslmdp.Document;
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
                    atx = new CBAutoText(atxname, gd.DocParts, ssldoc.Body);
                    atx.Insert();
                }
            }
            wrddoc.SaveAs(DOC_PATH_NAME);
        }

        class CBAutoText
        {
            private string name = string.Empty;
            private string category = string.Empty;
            private string content = string.Empty;
            private string containername = string.Empty;
            private string conatinercontent = string.Empty;
            private Body sslbody;

            // The easiest way to link to the parent.

            public CBAutoText(string autotextname, DocParts gdps, Body docbody)
            {
                var atx = (from dp in gdps
                            where dp.GetFirstChild<DocPartProperties>().DocPartName.Val == autotextname
                            select dp).Single();

                name = autotextname;
                category = atx.GetFirstChild<DocPartProperties>().Category.Name.Val; 
                content = atx.GetFirstChild<DocPartBody>().InnerXml;
                sslbody = docbody;
            }

            public string Name { get { return name; } }
            public string Category { get { return category; } }
            public string Content { get { return content; } }

            public void Insert()
            {
                Console.WriteLine("SdtElement is used");
                var cctrl = (from sdtCtrl in sslbody.Descendants<SdtElement>()
                            where sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == this.Category
                            || sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == this.Name
                            select sdtCtrl).Single();

                switch (cctrl.GetType().Name)
                {
                    case "SdtRun":
                        Console.WriteLine("Inserting this AutoText InnerXML==>{0}",this.Content);
                        //Console.ReadLine();
                        cctrl.GetFirstChild<SdtContentRun>().InnerXml=this.Content; 
                        break;

                    case "SdtBlock":
                        Console.WriteLine("Inserting this AutoText InnerXML==>{0}", this.Content);
                        //Console.ReadLine();
                        cctrl.GetFirstChild<SdtContentBlock>().InnerXml=this.Content;
                        break;

                    default:
                        break;
                }

                Console.WriteLine("Count after loop finished {0}", cctrl.Count());
                Console.WriteLine();
            }
        }

        class CBData { }
    }
}
