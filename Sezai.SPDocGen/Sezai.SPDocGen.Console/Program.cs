using System;
using Microsoft.CSharp;
using System.Collections.Generic;
using System.Text;
using Sezai.SPDocGen;
using Microsoft.SharePoint.Administration;

namespace Sezai.SPDocGen.Console
{
    public class Program
    {
        /// <summary>
        /// TODO: Implement parameters to the application to let users control what is output
        /// </summary>        
        static void Main(string[] args)
        {
            System.Console.WriteLine("Sezai.SPDocGen http://spdocgen.codeplex.com");
            System.Console.WriteLine();
            DateTime startTime = DateTime.Now;
            System.Console.WriteLine("Generating Documentation from SharePoint Farm...");
            System.Console.WriteLine();

            // Pass in SPFarm.Local and kick off the Farm XML Generation process
            FarmXmlGen farmXmlGen = new FarmXmlGen(SPFarm.Local);

            string exceptionMessage = "";
            try
            {
                // Some exceptions will be written to the XML file
                farmXmlGen.BuildFarmXml();
            }
            catch (Exception e)
            {
                exceptionMessage=e.ToString();
            }
            if (exceptionMessage == "")
            {
                // if there were no exceptions thrown from farmXmlGen.BuildFarmXml(); dump XML to screen
                System.Console.WriteLine(farmXmlGen.FarmXml.InnerXml);
                System.Console.WriteLine();
                System.Console.WriteLine();
                System.Console.WriteLine();
            }

            // Generate XML Filename
            string fileName = "SPDocGen_Farm_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string xmlFileName = fileName + ".xml";

            // Save farmXmlGen.FarmXml to disk
            System.Console.WriteLine("Saving Documentation to XML File: " + xmlFileName);
            System.Console.WriteLine();
            farmXmlGen.SaveFarmXml(xmlFileName);

            // Generate DOC Filename and specify XSLT transform file
            string docFileName = fileName + ".doc";
            string xsltFileName = "DocGen.xslt";

            // Transform and save the XML to DOC
            System.Console.WriteLine("Using " + xsltFileName + " to Transform ");
            System.Console.WriteLine(xmlFileName);
            System.Console.WriteLine("   to");
            System.Console.WriteLine(docFileName);
            System.Console.WriteLine();                       
            FarmDocGen farmDocGen = new FarmDocGen(xmlFileName, docFileName, xsltFileName);
            farmDocGen.CreateFarmWordDoc();

            // All done.
            DateTime finishTime = DateTime.Now;
            System.Console.WriteLine("Finished Generating Documentation, total time taken " + (finishTime-startTime).TotalSeconds + " seconds.");
            System.Console.WriteLine();
            System.Console.WriteLine("Hit Any Key To Exit...");
            System.Console.ReadLine();
        }
    }
}
