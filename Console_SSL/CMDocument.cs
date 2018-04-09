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

        //private const string DOC_PATH_NAME = @"C:\Users\Adria\Documents\Dev Projects\SSL\Documents\SSL_Doc.docx";
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
                    atxt = new CBAutoText
                    {
                        GDP = gdp,
                        Name = atxname
                    };
                    Console.WriteLine("AutoText Name ==> {0}", atxt.Name);

                    // Create a new relationship in the NEW document with the AutoText FOUND in the template
                    // Think about changing how this code is called.
                    atxt.CheckForRelationship();

                    // Retrieve RELATIONSHIP IDs from the document/document.xml in Main document/Document being created
                    Console.WriteLine("Count of Content Controls in this document is {0}\n", doc.Body.Descendants<SdtElement>().Count());
                    Console.ReadLine();

                    var cctrl = (from sdtCtrl in doc.Body.Descendants<SdtElement>()
                                    where sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == atxt.CCName
                                    || sdtCtrl.SdtProperties.GetFirstChild<SdtAlias>().Val == atxt.Name
                                    select sdtCtrl).SingleOrDefault();

                    XElement cc = XElement.Parse(cctrl.OuterXml);
                    IEnumerable<XAttribute> ccAttrs = cc.Descendants().Attributes();
                    var ccRelIDs = from attrib in ccAttrs
                                   where attrib.Value.Contains("rId")
                                   select attrib;

                    foreach (var item in mdp.Parts)
                    {
                        Console.WriteLine("RelIDs ==> {0}", item.RelationshipId);
                    }
                    Console.WriteLine();


                    if (ccRelIDs.Count() > 0)
                    {
                        foreach (var RelIDs in ccRelIDs)
                        {
                            Console.WriteLine("RelIDs = {0}", RelIDs.Value);
                            mdp.DeleteReferenceRelationship(ccRelIDs.FirstOrDefault().Value);
                        }
                    }


                    if (atxt.HasARelationship)
                    {
                        int i = 0;
                        foreach (var RelPart in atxt.RelationshipParts)
                        {
                            //Establish new relationship
                            atxt.NewRelID = mdp.CreateRelationshipToPart(RelPart);
                            i++;
                        }
                    }
                    cctrl.InnerXml = atxt.Content;
                }
                wrddoc.SaveAs(DOC_PATH_NAME);
                wrddoc.Close();

            }
        }



    }

    class CBAutoText
    {
        // The AutoText "structures" in a document
        // The MS Word Document Parts or XML containers
        private GlossaryDocumentPart gdp;
        private DocParts dps;

        // The description & content of AutoText
        private OpenXmlElement autotextDocPart;

        // Fields
        private string name = string.Empty;             // This is how I reference/call the AutoText
        private string category = string.Empty;         // This is the name of the content control the AutoText goes in
        private string content = string.Empty;          // This is the contents of the AutoText: Content Control with text all retrieved as XML
        private string contentcontainername = string.Empty;    // This IS the SAME as category. Category is where AutoText keeps the name of its content control
        private bool hasrelationship = false;
        private List<string> relationshipids;
        private List<OpenXmlPart> relationshipparts;

        // Added on 04/03/2018
        private string newrelid = string.Empty;

        public CBAutoText()
        {
            relationshipids = new List<string>();
            relationshipparts = new List<OpenXmlPart>();
        }

        public void CheckForRelationship()
        {
            // Retrieve RELATIONSHIP IDs from the document/document.xml in GLOSSARY PART/AUTOTEXT GALLERY
            XElement  AutoTextContent = XElement.Parse(this.content);
            IEnumerable<XAttribute> autotextPartAttribs = AutoTextContent.Descendants().Attributes();

            // LINQ over an XElement is easier than LINQ over an OpenXmlElement
            var AutoTextRelIDs = from attrb in autotextPartAttribs
                                 where attrb.Value.Contains("rId")
                                 select attrb;

            if (AutoTextRelIDs.Count() == 0)
            {
                Console.WriteLine("There are no relationship IDs in this Autotext/DocPart");
                hasrelationship = false;
            }
            else
            {
                hasrelationship = true;
                foreach (var relID in AutoTextRelIDs)
                {
                    /* Maybe, instead of adding relationship Parts I should raise an event
                     * to ADD a NEW RELATIONSHIP passing the part to the parent document
                     * AND returning the Relationship ID to the AutoText
                     * AND the newly received Rel ID is INSERTED into the AutoText content.
                    */
                    relationshipparts.Add(gdp.GetPartById(relID.Value));
                }
            }
        }


        public GlossaryDocumentPart GDP
        {
            set
            {
                gdp = value;
                dps = gdp.GlossaryDocument.DocParts;
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
                DocPartProperties autotextprops; // Using to navigate to the correct group of xml elements
                DocPartBody autotextbody; // Using to navigate to the correct group of xml elements
                SdtElement autotextcc;

                name = value;
                var atxt = (from dp in dps
                           where dp.GetFirstChild<DocPartProperties>().DocPartName.Val == name
                           select dp).SingleOrDefault();

                autotextprops = atxt.GetFirstChild<DocPartProperties>();
                autotextbody = atxt.GetFirstChild<DocPartBody>();
                autotextcc = autotextbody.Descendants<SdtElement>().SingleOrDefault();

                // Name of content control to insert retrieved autotext. Rename field to ContainControlName or something like that
                contentcontainername = autotextprops.Category.Name.Val;

                // Content to go into content control.  
                // 04/09/2018 Watch this because this is delicate and can invalidate the document.
                content = autotextcc.OuterXml;
                autotextDocPart = atxt;
            }
        }

        
        // Properties for the fields
        public string CCName { get { return contentcontainername; } }
        public string Content { get { return content; } }
        public bool HasARelationship { get { return hasrelationship; } }
        public string NewRelID { get { return newrelid; } set { newrelid = value; } }
        public List<OpenXmlPart> RelationshipParts { get {return relationshipparts; } }
    }
}