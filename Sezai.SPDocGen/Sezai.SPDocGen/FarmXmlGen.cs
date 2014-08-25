using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Sezai.SPDocGen
{
    public class FarmXmlGen
    {
        #region Public Properties
        public XmlDocument FarmXml;
        public SPFarm Farm;
        public XmlNode farmSiteCollectionsNode;
        #endregion

        #region Constructors
        public FarmXmlGen(SPFarm farm)
        {
            Farm = farm;
            FarmXml = new XmlDocument();

            // <?xml version="1.0" encoding="utf-8"?>
            XmlDeclaration xmlDeclaration = FarmXml.CreateXmlDeclaration("1.0", "utf-8", null);
            FarmXml.InsertBefore(xmlDeclaration, FarmXml.DocumentElement);
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Write the FarmXml XmlDocument property to disk
        /// </summary>
        /// <param name="FilePath">Filename of XML file to write to disk</param>
        public void SaveFarmXml(string FilePath)
        {
            XmlTextWriter xmlWriter = new XmlTextWriter(FilePath, null);
            FarmXml.Save(xmlWriter);
            xmlWriter.Close();
        }

        public void BuildFarmXml()
        {
            // Start with the Farm
            Farm = SPFarm.Local;
            XmlElement farmNode = FarmXml.CreateElement("Farm");

            FarmXml.AppendChild(farmNode);
            
            // StartTime
            XmlNode startTimeAttribute = FarmXml.CreateAttribute("StartTime");
            startTimeAttribute.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff");
            farmNode.Attributes.SetNamedItem(startTimeAttribute);

            // FarmVersion
            // TODO config file to store the description of SharePoint versions for lookup
            XmlNode versionNode = FarmXml.CreateElement("FarmVersion");
            Version farmVersion = Farm.BuildVersion;
            versionNode.InnerText = farmVersion.ToString();
            farmNode.AppendChild(versionNode);

            // FarmName
            XmlNode farmNameNode = FarmXml.CreateElement("FarmName");
            string farmName = Farm.DisplayName;
            farmNameNode.InnerText = farmName;
            farmNode.AppendChild(farmNameNode);

            // FarmId
            XmlNode farmIdNode = FarmXml.CreateElement("FarmId");
            Guid farmId = Farm.Id;
            farmIdNode.InnerText = farmId.ToString();
            farmNode.AppendChild(farmIdNode);

            // FarmStatus
            XmlNode farmStatusNode = FarmXml.CreateElement("FarmStatus");
            SPObjectStatus farmStatus = Farm.Status;
            farmStatusNode.InnerText = farmStatus.ToString();
            farmNode.AppendChild(farmStatusNode);

            // FarmServers            
            farmNode.AppendChild(GetServersNode(Farm.Servers, "FarmServers"));
            
            // FarmServices
            XmlNode farmServicesNode = FarmXml.CreateElement("FarmServices");
            farmServicesNode = GetFarmServicesNode(Farm.Services, farmServicesNode);
            farmNode.AppendChild(farmServicesNode);

            // FarmSolutions
            XmlNode farmSolutionsNode = GetFarmSolutionsNode(Farm.Solutions);
            farmNode.AppendChild(farmSolutionsNode);

            // FarmFeatureDefinitions
            XmlNode farmFeatureDefinitionsNode = FarmXml.CreateElement("FarmFeatureDefinitions");
            farmFeatureDefinitionsNode = GetFarmFeatureDefinitionsNode(Farm.FeatureDefinitions, farmFeatureDefinitionsNode);
            farmNode.AppendChild(farmFeatureDefinitionsNode);

            // FarmSiteCollections
            farmSiteCollectionsNode = FarmXml.CreateElement("FarmSiteCollections");

            // WebApplications
            SPWebServiceCollection webServices = new SPWebServiceCollection(Farm);
            // note this also populates the FarmSiteCollections
            XmlNode farmWebApplicationsNode = GetFarmWebApplicationsNode(webServices);
            farmNode.AppendChild(farmWebApplicationsNode);
            farmNode.AppendChild(farmSiteCollectionsNode);

            //ApplicationPools
            XmlNode farmApplicationPoolsNode = GetFarmApplicationPoolsNode(webServices);
            farmNode.AppendChild(farmApplicationPoolsNode);

            // FinishTime
            XmlNode finishTimeAttribute = FarmXml.CreateAttribute("FinishTime");
            DateTime finishTime = DateTime.Now;
            finishTimeAttribute.Value = finishTime.ToString("yyyy-MM-dd HH:mm:ss:fff");
            farmNode.Attributes.SetNamedItem(finishTimeAttribute);

            // A nicer GenerationTime format
            XmlNode generationTimeAttribute = FarmXml.CreateAttribute("GenerationTime");
            generationTimeAttribute.Value = finishTime.ToLongDateString() + " " + finishTime.ToLongTimeString();
            farmNode.Attributes.SetNamedItem(generationTimeAttribute);
        }
        #endregion

        #region Private Methods

        private XmlNode GetFarmApplicationPoolsNode(SPWebServiceCollection webServices)
        {
            XmlNode farmApplicationPoolsNode = FarmXml.CreateElement("FarmApplicationPools");
            int appPoolCount = 0;
            foreach (SPWebService webService in webServices)
            {
                foreach (SPApplicationPool appPool in webService.ApplicationPools)
                {
                    appPoolCount++;
                    XmlNode AppPoolNode = FarmXml.CreateElement("FarmApplicationPool");

                    List<AttributeValuePair> appPoolAttributes = SPAttributes.GetSPApplicationPoolAttributes(appPool);
                    foreach (AttributeValuePair appPoolAttribute in appPoolAttributes)
                    {
                        AppPoolNode.Attributes.SetNamedItem(GetXmlAttribute(appPoolAttribute));
                    }

                    farmApplicationPoolsNode.AppendChild(AppPoolNode);
                }
            }
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = appPoolCount.ToString();
            farmApplicationPoolsNode.Attributes.SetNamedItem(countAttribute);

            return farmApplicationPoolsNode;
        }

        private XmlNode GetFarmWebApplicationsNode(SPWebServiceCollection webServices)
        {
            XmlNode farmWebApplicationsNode = FarmXml.CreateElement("FarmWebApplications");
            int webAppCount = 0;
            foreach (SPWebService webService in webServices)
            {
                foreach (SPWebApplication webApplication in webService.WebApplications)
                {
                    webAppCount++;
                    XmlNode WebAppNode = FarmXml.CreateElement("WebApplication");
                    List<AttributeValuePair> webAppAttributes = SPAttributes.GetSPWebApplicationAttributes(webApplication);
                    foreach (AttributeValuePair webAppAttribute in webAppAttributes)
                    {
                        WebAppNode.Attributes.SetNamedItem(GetXmlAttribute(webAppAttribute));
                    }

                    // Get the Application Pool for the Web Application
                    XmlNode webAppApplicationPool = FarmXml.CreateElement("ApplicationPool");
                    SPApplicationPool appPool = webApplication.ApplicationPool;

                    List<AttributeValuePair> appPoolAttributes = SPAttributes.GetSPApplicationPoolAttributes(appPool);
                    foreach (AttributeValuePair appPoolAttribute in appPoolAttributes)
                    {
                        webAppApplicationPool.Attributes.SetNamedItem(GetXmlAttribute(appPoolAttribute));
                    }
                    WebAppNode.AppendChild(webAppApplicationPool);

                    // Get the site collections for the Web Application
                    XmlNode siteCollectionsNode = GetSiteCollectionsNode(webApplication.Sites,true);
                    WebAppNode.AppendChild(siteCollectionsNode);

                    // Get the content databases for the web application
                    XmlNode contentDatabasesNode = GetContentDatabasesNode(webApplication.ContentDatabases);
                    WebAppNode.AppendChild(contentDatabasesNode);

                    farmWebApplicationsNode.AppendChild(WebAppNode);
                }
            }
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = webAppCount.ToString();
            farmWebApplicationsNode.Attributes.SetNamedItem(countAttribute);

            return farmWebApplicationsNode;
        }

        private XmlNode GetContentDatabasesNode(SPContentDatabaseCollection contentDatabases)
        {
            XmlNode contentDatabasesNode = FarmXml.CreateElement("ContentDatabases");
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = contentDatabases.Count.ToString();
            contentDatabasesNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPContentDatabase contentDb in contentDatabases)
            {
                XmlNode contentDatabaseNode = FarmXml.CreateElement("ContentDatabase");

                List<AttributeValuePair> contentDbAttributes = SPAttributes.GetSPContentDatabaseAttributes(contentDb);
                foreach (AttributeValuePair contentDbAttribute in contentDbAttributes)
                {
                    contentDatabaseNode.Attributes.SetNamedItem(GetXmlAttribute(contentDbAttribute));
                }
                // If theres multiple content dbs, we also want to know what site collections are in each
                if (contentDatabases.Count > 1)
                {
                    XmlNode siteCollectionsNode = GetSiteCollectionsNode(contentDb.Sites,false);
                    contentDatabaseNode.AppendChild(siteCollectionsNode);
                }
                contentDatabasesNode.AppendChild(contentDatabaseNode);
            }

            return contentDatabasesNode;
        }

        /// <summary>
        /// Note the second parameter we also want to build a list of farm level site collections, but we don't want to add the same site collection twice
        /// </summary>        
        private XmlNode GetSiteCollectionsNode(SPSiteCollection siteCollections, bool AddToFarmSiteCollectionsNode)
        {
            XmlNode siteCollectionsNode = FarmXml.CreateElement("SiteCollections");
            // count
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = siteCollections.Count.ToString();
            siteCollectionsNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPSite site in siteCollections)
            {
                XmlNode siteCollectionNode = FarmXml.CreateElement("SiteCollection");

                List<AttributeValuePair> siteAttributes = SPAttributes.GetSPSiteAttributes(site);
                foreach (AttributeValuePair siteAttribute in siteAttributes)
                {
                    siteCollectionNode.Attributes.SetNamedItem(GetXmlAttribute(siteAttribute));
                }        
                // SPSiteAdministration gives us extra info on the SPSite, how cool is that!
                SPSiteAdministration siteAdmin = new SPSiteAdministration(site.ID);
                List<AttributeValuePair> siteAdminAttributes = SPAttributes.GetSPSiteAdministrationAttributes(siteAdmin);
                foreach (AttributeValuePair siteAdminAttribute in siteAdminAttributes)
                {
                    siteCollectionNode.Attributes.SetNamedItem(GetXmlAttribute(siteAdminAttribute));
                }
                siteCollectionsNode.AppendChild(siteCollectionNode);
                if (AddToFarmSiteCollectionsNode)
                {
                    XmlNode nodeclone = siteCollectionNode.Clone();
                    farmSiteCollectionsNode.AppendChild(nodeclone);
                }
                site.Dispose();
            }
            return siteCollectionsNode;
        }

        private XmlNode GetSiteCollectionAdminsNode(SPSite site)
        {
            SPWeb rootWeb = site.RootWeb;
            XmlNode siteCollectionAdminsNode = FarmXml.CreateElement("SiteCollectionAdministrators");

            int adminCount = 0;
            foreach (SPUser user in rootWeb.AllUsers)
            {
                if (user.IsSiteAdmin)
                {
                    XmlNode siteCollectionAdminNode = FarmXml.CreateElement("SiteCollectionAdministrator");

                    List<AttributeValuePair> userAttributes = SPAttributes.GetSPUserAttributes(user);
                    foreach (AttributeValuePair userAttribute in userAttributes)
                    {
                        siteCollectionAdminNode.Attributes.SetNamedItem(GetXmlAttribute(userAttribute));
                    }
                    siteCollectionAdminsNode.AppendChild(siteCollectionAdminNode);
                    adminCount++;
                }
            }
            // admin count
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = adminCount.ToString();
            siteCollectionAdminsNode.Attributes.SetNamedItem(countAttribute);

            rootWeb.Dispose();

            return siteCollectionAdminsNode;
        }

        private XmlNode GetFarmFeatureDefinitionsNode(SPFeatureDefinitionCollection farmFeatureDefinitions, XmlNode farmFeatureDefinitionsNode)
        {
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = farmFeatureDefinitions.Count.ToString();
            farmFeatureDefinitionsNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPFeatureDefinition featureDefinition in farmFeatureDefinitions)
            {
                XmlNode featureDefinitionNode = FarmXml.CreateElement("FeatureDefinition");
                List<AttributeValuePair> featureDefAttributes = SPAttributes.GetSPFeatureDefinitionAttributes(featureDefinition);
                foreach (AttributeValuePair featureDefAttribute in featureDefAttributes)
                {
                    featureDefinitionNode.Attributes.SetNamedItem(GetXmlAttribute(featureDefAttribute));
                }
                farmFeatureDefinitionsNode.AppendChild(featureDefinitionNode);
            }
            return farmFeatureDefinitionsNode;
        }

        private XmlNode GetFarmSolutionsNode(SPSolutionCollection farmSolutions)
        {
            XmlNode farmSolutionsNode = FarmXml.CreateElement("FarmSolutions");
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = farmSolutions.Count.ToString();
            farmSolutionsNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPSolution solution in farmSolutions)
            {
                XmlNode solutionNode = FarmXml.CreateElement("Solution");

                List<AttributeValuePair> solutionAttributes = SPAttributes.GetSPSolutionAttributes(solution);
                foreach (AttributeValuePair solutionAttribute in solutionAttributes)
                {
                    solutionNode.Attributes.SetNamedItem(GetXmlAttribute(solutionAttribute));
                }
                try
                {
                    solutionNode.AppendChild(GetServersNode(solution.DeployedServers, "DeployedServers"));
                }
                catch (Exception e)
                { 
                solutionNode.Attributes.SetNamedItem(GetXmlAttribute(new AttributeValuePair("Exception",e.ToString())));
                }
                try
                {
                    solutionNode.AppendChild(GetWebApplicationsNode(solution.DeployedWebApplications, "DeployedWebApplications"));
                }
                catch (Exception e)
                {
                    solutionNode.Attributes.SetNamedItem(GetXmlAttribute(new AttributeValuePair("Exception", e.ToString())));
                }
                farmSolutionsNode.AppendChild(solutionNode);
            }
            return farmSolutionsNode;
        }

        private XmlNode GetWebApplicationsNode(Collection<SPWebApplication> webApplications, string nodeName)
        {
            XmlNode webApplicationsNode = FarmXml.CreateElement(nodeName);
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = webApplications.Count.ToString();
            webApplicationsNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPWebApplication deployedWebApplication in webApplications)
            {
                XmlNode deployedWebApplicationNode = FarmXml.CreateElement("WebApplication");

                List<AttributeValuePair> webAppAttributes = SPAttributes.GetSPWebApplicationAttributes(deployedWebApplication);
                foreach (AttributeValuePair webAppAttribute in webAppAttributes)
                {
                    deployedWebApplicationNode.Attributes.SetNamedItem(GetXmlAttribute(webAppAttribute));
                }
                webApplicationsNode.AppendChild(deployedWebApplicationNode);
            }
            return webApplicationsNode;
        }
        /// <summary>
        // Note GetServersNode is overloaded to support Collection<SPServer> which is from SPSolution.DeployedServers
        /// </summary>
        private XmlNode GetServersNode(Collection<SPServer> servers, string nodeName)
        {            
            XmlNode serversNode = FarmXml.CreateElement(nodeName);
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = servers.Count.ToString();
            serversNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPServer deployedServer in servers)
            {
                XmlNode deployedServerNode = FarmXml.CreateElement("Server");

                List<AttributeValuePair> serverAttributes = SPAttributes.GetSPServerAttributes(deployedServer);
                foreach (AttributeValuePair serverAttribute in serverAttributes)
                {
                    deployedServerNode.Attributes.SetNamedItem(GetXmlAttribute(serverAttribute));
                }
                serversNode.AppendChild(deployedServerNode);
            }
            return serversNode;
        }
                
        private XmlNode GetServersNode(SPServerCollection servers, string nodeName)
        {
            XmlNode serversNode = FarmXml.CreateElement(nodeName);

            AttributeValuePair serverCount;
            serverCount.Attribute = "Count";
            serverCount.Value = servers.Count.ToString();

            XmlNode countAttribute = FarmXml.CreateAttribute(serverCount.Attribute);
            countAttribute.Value = serverCount.Value;
            serversNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPServer deployedServer in servers)
            {
                XmlNode deployedServerNode = FarmXml.CreateElement("Server");
                List<AttributeValuePair> serverAttributes = SPAttributes.GetSPServerAttributes(deployedServer);
                foreach (AttributeValuePair serverAttribute in serverAttributes)
                {
                    deployedServerNode.Attributes.SetNamedItem(GetXmlAttribute(serverAttribute));
                }

                // Get Server's service instances
                SPServiceInstanceCollection serverServiceInstances = deployedServer.ServiceInstances;
                XmlNode serverServicesNode = GetServerServiceInstanceNode(serverServiceInstances);
                deployedServerNode.AppendChild(serverServicesNode);
                serversNode.AppendChild(deployedServerNode);

                // Get Database Service Instances
                XmlNode databaseServiceInstances = GetDatabaseServiceInstancesNode(serverServiceInstances);
                if (databaseServiceInstances != null)
                    serversNode.AppendChild(databaseServiceInstances);
            }
            return serversNode;
        }

        // We're looking for servers that have a service instance of "Windows SharePoint Services Database"
        private XmlNode GetDatabaseServiceInstancesNode(SPServiceInstanceCollection serverServiceInstances)
        {
            XmlNode databases = null;
            foreach (SPServiceInstance serviceInstance in serverServiceInstances)
            {
                if (serviceInstance.TypeName == "Windows SharePoint Services Database")
                {
                    databases = GetServerDatabasesNode(serviceInstance);
                }
            }
            return databases;
        }
        private XmlNode GetServerServiceInstanceNode(SPServiceInstanceCollection serverServiceInstances)
        {
            XmlNode serverServicesNode = FarmXml.CreateElement("ServerServices");
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = serverServiceInstances.Count.ToString();
            serverServicesNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPServiceInstance serviceInstance in serverServiceInstances)
            {
                XmlNode serviceInstanceNode = FarmXml.CreateElement("Service");
                List<AttributeValuePair> serviceInstanceAttributes = SPAttributes.GetSPServiceInstanceAttributes(serviceInstance);

                foreach (AttributeValuePair serviceInstanceAttribute in serviceInstanceAttributes)
                {
                    serviceInstanceNode.Attributes.SetNamedItem(GetXmlAttribute(serviceInstanceAttribute));
                }
                serverServicesNode.AppendChild(serviceInstanceNode);
            }
            return serverServicesNode;
        }

        private XmlNode GetServerDatabasesNode(SPServiceInstance serviceInstance)
        {
            XmlNode databaseServiceInstances = FarmXml.CreateElement("DatabaseServiceInstances");
            SPDatabaseServiceInstanceCollection databaseServiceInstanceCollection = new SPDatabaseServiceInstanceCollection(serviceInstance.Server);

            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = databaseServiceInstanceCollection.Count.ToString();
            databaseServiceInstances.Attributes.SetNamedItem(countAttribute);

            // Need to test this in a farm using multiple SQL Servers to store databases to validate this
            foreach (SPDatabaseServiceInstance databaseServiceInstance in databaseServiceInstanceCollection)
            {
                XmlNode databaseServiceInstanceNode = FarmXml.CreateElement("DatabaseServiceInstance");

                XmlNode databases = FarmXml.CreateElement("Databases");
                XmlNode count = FarmXml.CreateAttribute("Count");
                count.Value = databaseServiceInstance.Databases.Count.ToString();
                databases.Attributes.SetNamedItem(count);

                foreach (SPDatabase spDatabase in databaseServiceInstance.Databases)
                {
                    XmlNode database = FarmXml.CreateElement("Database");
                    List<AttributeValuePair> databaseAttribues = SPAttributes.GetSPDatabaseAttributes(spDatabase);

                    foreach (AttributeValuePair databaseAttribute in databaseAttribues)
                    {
                        database.Attributes.SetNamedItem(GetXmlAttribute(databaseAttribute));
                    }
                    databases.AppendChild(database);
                }
                databaseServiceInstanceNode.AppendChild(databases);
                databaseServiceInstances.AppendChild(databaseServiceInstanceNode);
            }
            return databaseServiceInstances;
        }
        
        private XmlNode GetFarmServicesNode(SPServiceCollection farmServices, XmlNode farmServicesNode)
        {
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = farmServices.Count.ToString();
            farmServicesNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPService service in farmServices)
            {
                XmlNode serviceNode = FarmXml.CreateElement("FarmService");

                List<AttributeValuePair> serviceAttributes = SPAttributes.GetSPServiceAttributes(service);
                foreach (AttributeValuePair serviceAttribute in serviceAttributes)
                {
                    serviceNode.Attributes.SetNamedItem(GetXmlAttribute(serviceAttribute));
                }

                // Get the service's Job Definitions if it has any
                SPJobDefinitionCollection serviceJobDefinitions = service.JobDefinitions;
                serviceNode.AppendChild(GetJobDefinitionsNode(serviceJobDefinitions, "ServiceJobDefinitions"));
                
                farmServicesNode.AppendChild(serviceNode);
            }

            return farmServicesNode;
        }

        private XmlNode GetJobDefinitionsNode(SPJobDefinitionCollection serviceJobDefinitions, string nodeName)
        {
            XmlNode jobDefinitionsNode = FarmXml.CreateElement(nodeName);
            XmlNode countAttribute = FarmXml.CreateAttribute("Count");
            countAttribute.Value = serviceJobDefinitions.Count.ToString();
            jobDefinitionsNode.Attributes.SetNamedItem(countAttribute);

            foreach (SPJobDefinition jobDefinition in serviceJobDefinitions)
            {
                XmlNode jobDefinitionNode = FarmXml.CreateElement("JobDefinition");

                List<AttributeValuePair> jobDefAttributes = SPAttributes.GetSPJobDefinitionAttributes(jobDefinition);
                foreach (AttributeValuePair jobDefAttribute in jobDefAttributes)
                {
                    jobDefinitionNode.Attributes.SetNamedItem(GetXmlAttribute(jobDefAttribute));
                }
                jobDefinitionsNode.AppendChild(jobDefinitionNode);
            }
            return jobDefinitionsNode;
        }

        /// <summary>
        /// Creates an XML Element from an AttributeValurPair
        /// </summary>        
        private XmlNode GetXmlAttribute(AttributeValuePair attributeValuePair)
        {
            XmlNode attribute = FarmXml.CreateAttribute(attributeValuePair.Attribute);
            attribute.Value = attributeValuePair.Value;
            return attribute;
        }
        #endregion
    }
}
