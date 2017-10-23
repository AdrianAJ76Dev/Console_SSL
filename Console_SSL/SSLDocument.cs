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

        private Document NewDocument()
        {
            WordprocessingDocument newdoc = WordprocessingDocument.CreateFromTemplate(template);
            sslmdp = newdoc.MainDocumentPart;
            return newdoc.MainDocumentPart.Document;
        }

        public void AddAutoText(string[] AutoTextName)
        {
            List<CBAutoText> atxs = new List<CBAutoText>();
            foreach (string atxname in AutoTextName)
            {
                atxs.Add(new CBAutoText(this));
            }
        }

        class CBAutoText
        {
            private SSLDocument ssldoc;
            private GlossaryDocument gd;
            private GlossaryDocumentPart gdp;
            private string name = string.Empty;
            private string containername = string.Empty;
            private string content = string.Empty;

            public CBAutoText(SSLDocument parent)
            {
                ssldoc = parent;
                gdp = ssldoc.Mdp.GlossaryDocumentPart;
            }

            public string AutoTextContent
            {
                get{return content;}
            }
        }

        class CBData { }
    }
}
