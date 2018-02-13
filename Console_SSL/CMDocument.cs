using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Open XML References
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using wrd10 = DocumentFormat.OpenXml.Office2010.Word;
using wrd13 = DocumentFormat.OpenXml.Office2013.Word;

// For...
using System.Reflection;


namespace Console_SSL
{
    class CMDocument
    {
        /* Create new document from template
         * Add AutoText to new document
         * Update Custom XML in new document
         * Save document
         * Display new document
         */
        private string template = string.Empty;
        private WordprocessingDocument wrddoc;
        private MainDocumentPart mdp;
        private Document doc;
        private GlossaryDocumentPart gdp;

        private const string DOC_PATH_NAME = @"D:\Dev Projects\SSL\Documents\SSL_Doc.docx";

        private CBAutoText atxt;

        public CMDocument(string templatefullname)
        {
            template = templatefullname;
            doc = NewDocument();
        }

        private Document NewDocument()
        {
            WordprocessingDocument newdoc = WordprocessingDocument.CreateFromTemplate(template);
            wrddoc = newdoc;
            mdp = newdoc.MainDocumentPart;
            return mdp.Document;
        }

        public void BuildDocument(string[] AutoTextName)
        {
            /* Here's where I look in the Glossary Document Part to determine it there IS AutoText
             * If there isn't, then there's no use in creating the List of AutoText objects!
             */
            gdp = mdp.GlossaryDocumentPart;
            if (gdp != null)
            {
                foreach (string atxname in AutoTextName)
                {
                    atxt = new CBAutoText();
                    atxt.DocPartsDoc= gdp.GlossaryDocument;
                    atxt.Name = atxname;
                    ReplaceContentControlWithAutoTextInAContentControl();
                    wrddoc.SaveAs(DOC_PATH_NAME);
                }
            }
        }

        private void ReplaceContentControlWithAutoTextInAContentControl()
        {
            Console.WriteLine("Count of Content Controls is {0}\n", doc.Body.Descendants<SdtElement>().Count());
            var cctrl = (from sdtCtrl in doc.Body.Descendants<SdtElement>()
                         where sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == atxt.Category
                         || sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == atxt.Name
                         select sdtCtrl).Single();

            cctrl.InnerXml = atxt.Content;

            //Drawing ImageSignatory = gDocPartBodySignatures.FirstOrDefault<DocPartBody>().Descendants<Drawing>().FirstOrDefault();
            if (mdp.GetPartsCountOfType<GlossaryDocumentPart>() != 0)
            {
                GlossaryDocumentPart gNewDoc = mdp.GetPartsOfType<GlossaryDocumentPart>().FirstOrDefault();
                DocumentFormat.OpenXml.Drawing.Blip blpSignature = mdp.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                string OldRelID = blpSignature.Embed.Value;
                ImagePart ImageSignatory = (ImagePart)gNewDoc.GetPartById(OldRelID);
                if (ImageSignatory != null)
                {
                    // Fails here because it ASSIGNED a relationship ID the 1st time around
                    mdp.CreateRelationshipToPart(ImageSignatory, "rId30");
                    blpSignature.Embed.Value = "rId30";
                }
            }

        }

    }

    class CBAutoText
    {
        private string name = string.Empty;
        private string category = string.Empty;
        private string content = string.Empty;

        private GlossaryDocument gdoc;
        private DocParts dps;
        private SdtElement cctrl;
        private DocPart dp;

        public string Category { get { return category; } }
        public string Content { get { return content; } }

        public CBAutoText() { }

        public GlossaryDocument DocPartsDoc
        {
            set
            {
                gdoc = value;
                dps = gdoc.DocParts;
            }
        }

        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
                var atxt = (from dp in dps
                           where dp.GetFirstChild<DocPartProperties>().DocPartName.Val == name
                           select dp).Single();

                category = atxt.GetFirstChild<DocPartProperties>().Category.Name.Val;
                content = atxt.GetFirstChild<DocPartBody>().Descendants<SdtElement>().FirstOrDefault().InnerXml;
            }
        }

    }
}
