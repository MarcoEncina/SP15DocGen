using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using System.Xml.XPath;

namespace Sezai.SPDocGen
{
    // Convert the XML to a word doc using XSLT
    public class FarmDocGen
    {
        #region constructors
        public FarmDocGen(string FarmXmlFilename, string OutputDocFileName, string XsltFileName)
        {
            this.XmlFilename = FarmXmlFilename;
            this.XsltFileName = XsltFileName;
            this.OutputDocFileName = OutputDocFileName;
        }
        public FarmDocGen(string FarmXmlFilename)
        {
            this.XmlFilename = FarmXmlFilename;
            this.OutputDocFileName = FarmXmlFilename.Replace(".xml", ".doc");
            this.XsltFileName = "DocGen.xslt";
        }
        public FarmDocGen(XmlDocument farmXml)
        {
            FarmXml=farmXml;
        }
        #endregion

        #region public properties
        public string XmlFilename;
        public string OutputDocFileName;
        public XmlDocument FarmXml=null;
        public XmlDocument FarmDoc;
        public string XsltFileName = "DocGen.xslt";
        #endregion

        #region public methods
        public void CreateFarmWordDoc()
        {
            // Create XML Reader (to read FarmXml) and XML Writer, to write DOC file
            XmlReader reader = XmlReader.Create(XmlFilename);
            XmlWriter writer = XmlWriter.Create(OutputDocFileName);

            // Create and Load the XSLT file
            XsltSettings settings = new XsltSettings();
            settings.EnableScript = true;
            XslCompiledTransform transform = new XslCompiledTransform();
            transform.Load(XsltFileName,settings,null);
            
            // Transform the FarmXml to the DOC file
            transform.Transform(reader,writer);
        }
        #endregion
    }
}
