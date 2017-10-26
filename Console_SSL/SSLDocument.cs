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
        private Document ssldoc;
        private MainDocumentPart sslmdp;

        public SSLDocument() { }
        public SSLDocument(string templatefullname)
        {
            template = templatefullname;
            ssldoc = NewDocument();
        }

        public Document Doc { get; set; }
        public MainDocumentPart Mdp { get; set; }
        public WordprocessingDocument WrdDoc { get; set; }

        private Document NewDocument()
        {
            WordprocessingDocument newdoc = WordprocessingDocument.CreateFromTemplate(template);
            sslmdp = newdoc.MainDocumentPart;
            this.WrdDoc = newdoc;
            this.Mdp = sslmdp;
            this.Doc = sslmdp.Document;
            return newdoc.MainDocumentPart.Document;
        }

        public void AddAutoText(string[] AutoTextName)
        {
            List<CBAutoText> atxs = new List<CBAutoText>();
            foreach (string atxname in AutoTextName)
            {
                atxs.Add(new CBAutoText(this, atxname));
            }

            foreach (var atx in atxs)
            {
                Console.WriteLine("AutoText Name: {0}\n ", atx.Name);
                Console.WriteLine();
                int i = 0;
                foreach (object content in atx.AutoTextContent)
                {
                    i++;
                    Console.WriteLine("Content {1}: {0}\n", content, i);
                }

            }
        }

        class CBAutoText
        {
            private SSLDocument ssldoc;
            private GlossaryDocumentPart gdp;
            private GlossaryDocument gd;
            private DocParts gdocparts;
            private string containername = string.Empty;
            private string content = string.Empty;

            public CBAutoText(SSLDocument parentdoc, string autotextname)
            {
                ssldoc = parentdoc;
                this.Name = autotextname;
                gdp = ssldoc.Mdp.GlossaryDocumentPart;
                if (gdp != null)
                {
                    gd = gdp.GlossaryDocument;
                    gdocparts = gd.DocParts;
                }
                // 10/25/2017 Here I should cancel creating an autotext object if there is not autotext
            }

            public string Name { get; set; }

            public Array AutoTextContent
            {
                get
                {
                    var content = from gdocpart in gdocparts
                                   where gdocpart.Descendants<DocPartProperties>().First().DocPartName.Val == this.Name
                                   select new
                                   {
                                       aXml=gdocpart.Descendants<DocPartBody>().FirstOrDefault().InnerXml,
                                       aText=gdocpart.Descendants<DocPartBody>().FirstOrDefault().InnerText
                                   };
;
                    return content.ToArray();
                }
            }
        }

        class CBData { }
    }
}
