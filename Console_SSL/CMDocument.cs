using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml;
using System.Xml.Linq;

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

        private const string DOC_PATH_NAME = @"C:\Users\Adria\Documents\Dev Projects\SSL\Documents\SSL_Doc.docx";
        //private const string DOC_PATH_NAME = @"C:\Users\ajones\Documents\Automation\Code\Word\SSL Work\SSL_Doc.docx";

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
                    atxt.SearchForRelationshipInAutoTextEntry();
                }
                wrddoc.SaveAs(DOC_PATH_NAME);
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

        // 3-23-18
        // Standard XML element
        private XElement xeAutoText = null;

        // 2-28-2018 Addition
        private string RelIDAutoText;
        private string RelIDDocument;

        // The description & content of AutoText
        private OpenXmlElement autotextDocPart;
        
        // Fields
        private string name = string.Empty;             // This is how I reference/call the AutoText
        private string category = string.Empty;         // This is the name of the content control the AutoText goes in
        private string content = string.Empty;          // This is the contents of the AutoText: Content Control with text all retrieved as XML
        private string containername = string.Empty;    // This IS the SAME as category. Category is where AutoText keeps the name of its content control
        private bool hasrelationship = false;
        private List<string> relationshipidsatxt;
        private List<string> relationshipidsdoc;

        public CBAutoText()
        {
            relationshipidsatxt = new List<string>();
            relationshipidsdoc = new List<string>();
        }

        public void SearchForRelationshipInAutoTextEntry()
        {
            // Retrieve RELATIONSHIP IDs from the document/document.xml in GLOSSARY PART/AUTOTEXT GALLERY
            XElement docAutoText = XElement.Parse(autotextDocPart.OuterXml);
            IEnumerable<XAttribute> autotextPartAttribs = docAutoText.Descendants().Attributes();
            // LINQ over an XElement is easier than LINQ over an OpenXmlElement
            var AutoTextRelIDs = from attrb in autotextPartAttribs
                                 where attrb.Value.Contains("rId")
                                 select new { attrb.Value, attrb.Parent, attrb.Name};



            if (AutoTextRelIDs.Count()==0)
            {
                Console.WriteLine("There are no relationship IDs in this Autotext/DocPart");
                hasrelationship = false;
            }
            else
            {
                Console.WriteLine("Relationship IDs found in AutoTextRelIDs");
                foreach (var relID in AutoTextRelIDs)
                {
                    RelIDAutoText = relID.Value;
                    Console.WriteLine("attrb ==> {0}\t{1}\t{2}", relID, gdp.GetPartById(relID.Value).GetType().Name, gdp.GetPartById(relID.Value).Uri);
                }
                hasrelationship = true;
            }
            Console.ReadLine();


            // Retrieve RELATIONSHIP IDs from the document/document.xml in GLOSSARY PART/AUTOTEXT GALLERY
            XElement docMain = XElement.Parse(parentmdp.Document.OuterXml);

            // THIS brings back ALL attributes instead of attributes under a specific OpenXmlElement
            IEnumerable<XAttribute> mainDocAttribs = docMain.Descendants().Attributes();
            var MainDocRelIDs = from attrb in mainDocAttribs
                                 where attrb.Value.Contains("rId")
                                 select attrb.Value;

            if (MainDocRelIDs.Count() == 0)
            {
                hasrelationship = false;
                Console.WriteLine("There are no relationship IDs in this main document part");
            }
            else
            {
                Console.WriteLine("Relationship IDs found in MainDocRelIDs");
                foreach (var relID in MainDocRelIDs)
                {
                    Console.WriteLine("attrb ==> {0}\t{1}\t{2}", relID, gdp.GetPartById(relID).GetType().Name, gdp.GetPartById(relID).Uri);
                    //Establish new relationship
                    IdPartPair RelationshipPair = (from autotextrel in parentmdp.Parts
                                                       where autotextrel.RelationshipId.Equals(relID)
                                                       select autotextrel).SingleOrDefault();

                    parentmdp.DeleteReferenceRelationship(RelIDDocument);
                    parentmdp.CreateRelationshipToPart(RelationshipPair.OpenXmlPart, RelIDDocument);
                }
                hasrelationship = true;
            }
            Console.ReadLine();

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

        public void PartRelPairGlossaryDoc()
        {
            Console.WriteLine("{0}", gdp.RootElement.GetType().Name);
            Console.WriteLine("Relationship Count ==> {0}", gdp.Parts.Count());
            Console.WriteLine();
            foreach (IdPartPair item in gdp.Parts)
            {
                Console.WriteLine("Content Type ==> {0}", item.OpenXmlPart.ContentType);
                Console.WriteLine("Uri ==> {0}", item.OpenXmlPart.Uri);
                Console.WriteLine("RelationshipId ==> {0}", item.RelationshipId);
                Console.WriteLine("OpenXmlPart ==> {0}", item.OpenXmlPart.GetType().Name);
                Console.WriteLine();
            }
            Console.WriteLine("ImagePart Count ==> {0}", gdp.GetPartsCountOfType<ImagePart>());
            Console.ReadLine();
        }

        public void PartRelPairMainDoc()
        {
            Console.WriteLine("{0}", parentmdp.RootElement.GetType().Name);
            Console.WriteLine("Relationship Count ==> {0}", parentmdp.Parts.Count());
            Console.WriteLine();
            foreach (IdPartPair item in parentmdp.Parts)
            {
                Console.WriteLine("Content Type ==> {0}", item.OpenXmlPart.ContentType);
                Console.WriteLine("Uri ==> {0}", item.OpenXmlPart.Uri);
                Console.WriteLine("RelationshipId ==> {0}", item.RelationshipId);
                Console.WriteLine("OpenXmlPart ==> {0}", item.OpenXmlPart.GetType().Name);
                Console.WriteLine();
            }
            Console.WriteLine("ImagePart Count ==> {0}", parentmdp.GetPartsCountOfType<ImagePart>());
            Console.ReadLine();
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
                           select dp).SingleOrDefault();

                // Name of content control to insert retrieved autotext. Rename field to ContainControlName or something like that
                category = atxt.GetFirstChild<DocPartProperties>().Category.Name.Val; 
                
                // Containt to go into content control
                content = atxt.GetFirstChild<DocPartBody>().Descendants<SdtElement>().FirstOrDefault().InnerXml;
                autotextDocPart = atxt;
            }
        }




        // Properties for the fields
        public string Category { get { return category; } }
        public string Content { get { return content; } }
        public bool HasARelationship { get { return hasrelationship; } }
        //public List<string> RelationshipIDs { get { return RelIDAutoText; } }
    }
}