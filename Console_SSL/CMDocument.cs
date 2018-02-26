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
                    atxt = new CBAutoText
                    {
                        ParentMdp = mdp,
                        GDP = gdp,
                        Name = atxname
                    };
                    Console.WriteLine("AutoText Name ==> {0}", atxt.Name);

                    // Create a new relationship in the NEW document with the AutoText FOUND in the template
                    atxt.CheckForRelationshipInAutoTextEntry();

                    //atxt.IdentifyPartsAndRelationships();
                    //atxt.IdentifyPartsAndRelationshipsMDP();
                    //atxt.InvestigatingDocPart();
                    ReplaceContentControlWithAutoTextInAContentControl();
                    Console.ReadLine();
                }
                wrddoc.SaveAs(DOC_PATH_NAME);

                // Form the relationships found in the AutoText in the Glossary Document
                foreach (string relshpID in atxt.RelationshipID)
                {
                    OpenXmlPart AutoTextRelationshipPart = gdp.GetPartById(relshpID);
                    switch (AutoTextRelationshipPart.GetType().Name)
                    {
                        // Figure out what to switch on.  It'll be on OpenXmlPart Type
                        case "ImagePart":
                            ImagePart ImageSignatory = (ImagePart)gdp.GetPartById(relshpID);
                            if (ImageSignatory != null)
                            {
                                mdp.CreateRelationshipToPart(ImageSignatory, "rId20"); //This hardcoded Relationship ID has to be changed.
                            }
                            break;

                        default:
                            break;
                    }
                }
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
        private List<string> relationshipids;

        public CBAutoText()
        {
            relationshipids = new List<string>();
        }

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
            Console.WriteLine();
        }

        public void CheckForRelationshipInAutoTextEntry()
        {
            hasrelationship = false;
            var ElementsWithRelID = from el in autotextDocPart.GetFirstChild<DocPartBody>().Descendants<OpenXmlElement>()
                                        where el.HasAttributes
                                        select from attr in el.GetAttributes()
                                            where attr.Value.Contains("rId")
                                            select attr.Value;
            foreach (var elems in ElementsWithRelID)
            {
                foreach (var relid in elems)
                {
                    hasrelationship = true;
                    relationshipids.Add(relid.ToString());
                }
            }
        }

        // Properties for the fields
        public string Category { get { return category; } }
        public string Content { get { return content; } }
        public List<string> RelationshipID { get { return relationshipids; } }
        public bool HasARelationship { get { return hasrelationship; } }
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
    }