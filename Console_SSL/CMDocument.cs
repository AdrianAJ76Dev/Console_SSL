using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Open XML References
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
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

        //private const string DOC_PATH_NAME = @"D:\Dev Projects\SSL\Documents\SSL_Doc.docx";
        private const string DOC_PATH_NAME = @"C:\Users\ajones\Documents\Automation\Code\Word\SSL Work\SSL_Doc.docx";

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
            wrddoc.SaveAs(DOC_PATH_NAME);
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
                    atxt.ParentMdp = mdp;
                    atxt.GDP = gdp;
                    atxt.Name = atxname;
                    Console.WriteLine("AutoText Name ==> {0}", atxt.Name);

                    //atxt.IdentifyPartsAndRelationships();
                    //atxt.IdentifyPartsAndRelationshipsMDP();
                    //atxt.InvestigatingDocPart();
                    atxt.CheckForRelationshipInAutoTextEntry();
                    Console.ReadLine();
                    //ReplaceContentControlWithAutoTextInAContentControl();
                }
                //wrddoc.SaveAs(DOC_PATH_NAME);
            }
        }

        private void ReplaceContentControlWithAutoTextInAContentControl()
        {
            Console.WriteLine("Count of Content Controls is {0}\n", doc.Body.Descendants<SdtElement>().Count());
            Console.ReadLine();
            var cctrl = (from sdtCtrl in doc.Body.Descendants<SdtElement>()
                         where sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == atxt.Category
                         || sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == atxt.Name
                         select sdtCtrl).Single();

            //CheckIfImageInAutoText();
            cctrl.InnerXml = atxt.Content;
        }

        private void CheckIfImageInAutoText()
        {
            Blip blpSignature = mdp.Document.Descendants<Blip>().FirstOrDefault();
            if (blpSignature != null)
            {
                string OldRelID = blpSignature.Embed.Value;
                ImagePart ImageSignatory = (ImagePart)mdp.GetPartById(OldRelID);
                if (ImageSignatory != null)
                {
                    mdp.CreateRelationshipToPart(ImageSignatory, "rId20");
                    blpSignature.Embed.Value = "rId20";
                }
            }
        }
    }

    class CBAutoText
    {
        // The AutoText "structures" in a document
        // The MS Word Document Parts or XML containers
        private MainDocumentPart parentmdp;
        private GlossaryDocumentPart gdp;
        private DocParts dps;

        // The description & content of AutoText
        private OpenXmlElement autotextDocPart;
        
        // Fields
        private string name = string.Empty;             // This is how I reference/call the AutoText
        private string category = string.Empty;         // This is the name of the content control the AutoText goes in
        private string content = string.Empty;          // This is the contents of the AutoText: Content Control with text all retrieved as XML
        private string containername = string.Empty;    // This IS the SAME as category. Category is where AutoText keeps the name of its content control
        private bool hasrelationship = false;
        private string relationshipid = string.Empty;

        public CBAutoText() { }

        public MainDocumentPart ParentMdp
        {
            set { parentmdp = value; }
        }

        public GlossaryDocumentPart GDP
        {
            set
            {
                gdp = value;
                dps = gdp.GlossaryDocument.DocParts;
            }
        }

        public void IdentifyPartsAndRelationships()
        {
            Console.WriteLine("GlossaryDocument Parts Count ==> {0}", gdp.Parts.Count());
            foreach (IdPartPair partrel in gdp.Parts)
            {
                Console.WriteLine("OpenXmlPart = {0}", partrel.OpenXmlPart);
                Console.WriteLine("Relationship ID = {0}",partrel.RelationshipId);
                Console.WriteLine();
            }
        }

        public void IdentifyPartsAndRelationshipsMDP()
        {
            Console.WriteLine("MainDocumentPart Parts Count ==> {0}", parentmdp.Parts.Count());
            foreach (IdPartPair partrel in parentmdp.Parts)
            {
                Console.WriteLine("OpenXmlPart = {0}", partrel.OpenXmlPart);
                Console.WriteLine("Relationship ID = {0}", partrel.RelationshipId);
                Console.WriteLine();
            }
        }

        public void InvestigatingDocPart()
        {
            int DescendentsCount = autotextDocPart.GetFirstChild<DocPartBody>().Descendants().Count();
            Console.WriteLine("Descendents Count ==> {0}",DescendentsCount);
            CheckForRelationshipInAutoTextEntry();

            /*
            foreach (OpenXmlElement item in autotextDocPart.GetFirstChild<DocPartBody>().Descendants())
            {
                Console.WriteLine("OpenXmlElement:Local Name ==> {0}", item.LocalName);
                Console.WriteLine("OpenXmlElement:Attr Count ==> {0}", item.GetAttributes().Count());
                foreach (var attr in item.GetAttributes())
                {
                    Console.WriteLine("Attr Value ==> {0}", attr.Value);
                    Console.WriteLine("Attr Name ==> {0}", attr.LocalName);
                    Console.WriteLine("Attr Namespace URI ==> {0}", attr.NamespaceUri);
                }
            }
            */

            Console.WriteLine();
        }

        
        public void CheckForRelationshipInAutoTextEntry()
        {
            var aAttr = from el in autotextDocPart.GetFirstChild<DocPartBody>().Descendants<OpenXmlElement>()
                        where el.HasAttributes
                        select el.GetAttributes();
            foreach (var eattr in aAttr)
            {
                Console.WriteLine("Count of eattr Items ==> {0}", eattr.Count()); 
                foreach (var item in eattr)
                {
                    Console.WriteLine("\t{1}:Item Value ==> {0}", item.Value, item.LocalName); 
                }
            }

            var attr = from attrb in aAttr
                       select attrb;
            foreach (var items in attr)
            {
                Console.WriteLine("Items Count ==> {0}", items.Count());
                foreach (var item in items)
                {
                    Console.WriteLine("\t{1}:Item Value ==> {0}", item.Value,item.LocalName);
                }
            }
        }

        // Properties for the fields
        public string Category { get { return category; } }
        public string Content { get { return content; } }
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

                autotextDocPart = atxt;
            }
        }
        public string RelationshipID
        {
            get { return relationshipid; }
        }
    }
}
