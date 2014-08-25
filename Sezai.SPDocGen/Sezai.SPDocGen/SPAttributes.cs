using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using Microsoft.SharePoint.Utilities;

namespace Sezai.SPDocGen
{
    /// <summary>
    /// Each method in this class returns a generic List of AttributeValuePair for a given SharePoint Object
    /// </summary>
    public static class SPAttributes
    {
        public static List<AttributeValuePair> GetSPServerAttributes(SPServer server)
        {
            List<AttributeValuePair> serverAttributes = new List<AttributeValuePair>();
            try
            {
                serverAttributes.Add(new AttributeValuePair("Name", server.Name));
                serverAttributes.Add(new AttributeValuePair("Id", server.Id.ToString()));
                serverAttributes.Add(new AttributeValuePair("Status", server.Status.ToString()));
                serverAttributes.Add(new AttributeValuePair("Address", server.Address));
                serverAttributes.Add(new AttributeValuePair("Role", server.Role.ToString()));
            }
            catch (Exception e)
            {
                serverAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            }
            return serverAttributes;
        }

        public static List<AttributeValuePair> GetSPServiceInstanceAttributes(SPServiceInstance serviceInstance)
        {
            List<AttributeValuePair> serviceInstanceAttributes = new List<AttributeValuePair>();
            try
            {
                serviceInstanceAttributes.Add(new AttributeValuePair("Instance", serviceInstance.Instance));
                serviceInstanceAttributes.Add(new AttributeValuePair("TypeName", serviceInstance.TypeName));
                serviceInstanceAttributes.Add(new AttributeValuePair("Id", serviceInstance.Id.ToString()));
                serviceInstanceAttributes.Add(new AttributeValuePair("SystemService", serviceInstance.SystemService.ToString()));
            }
            catch (Exception e)
            {
                serviceInstanceAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return serviceInstanceAttributes;
        }

        public static List<AttributeValuePair> GetSPServiceAttributes(SPService service)
        {
            List<AttributeValuePair> serviceAttributes = new List<AttributeValuePair>();
            try
            {
                serviceAttributes.Add(new AttributeValuePair("TypeName", service.TypeName));
                serviceAttributes.Add(new AttributeValuePair("Name", service.Name));
                serviceAttributes.Add(new AttributeValuePair("Id", service.Id.ToString()));
            }
            catch (Exception e)
            {
                serviceAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return serviceAttributes;
        }

        public static List<AttributeValuePair> GetSPDatabaseAttributes(SPDatabase database)
        {
            List<AttributeValuePair> databaseAttributes = new List<AttributeValuePair>();
            try
            {
                databaseAttributes.Add(new AttributeValuePair("Name", database.Name));
                databaseAttributes.Add(new AttributeValuePair("Server", database.Server.Name));
                databaseAttributes.Add(new AttributeValuePair("DisplayName", database.Server.Name));
                databaseAttributes.Add(new AttributeValuePair("Id", database.Id.ToString()));
                databaseAttributes.Add(new AttributeValuePair("ConnectionString", database.DatabaseConnectionString));
                databaseAttributes.Add(new AttributeValuePair("DiskSizeRequired", (database.DiskSizeRequired / 1024 / 1024).ToString() + " MB"));
            }
            catch (Exception e)
            {
                databaseAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return databaseAttributes;
        }

        public static List<AttributeValuePair> GetSPContentDatabaseAttributes(SPContentDatabase contentDatabase)
        {
            List<AttributeValuePair> contentDatabaseAttributes = new List<AttributeValuePair>();
            try
            {
                contentDatabaseAttributes.Add(new AttributeValuePair("Id", contentDatabase.Id.ToString()));
                contentDatabaseAttributes.Add(new AttributeValuePair("DisplayName", contentDatabase.DisplayName));
                contentDatabaseAttributes.Add(new AttributeValuePair("Name", contentDatabase.Name));
                contentDatabaseAttributes.Add(new AttributeValuePair("Server", contentDatabase.Server));
                contentDatabaseAttributes.Add(new AttributeValuePair("ConnectionString", contentDatabase.DatabaseConnectionString));
                contentDatabaseAttributes.Add(new AttributeValuePair("Status", contentDatabase.Status.ToString()));
                contentDatabaseAttributes.Add(new AttributeValuePair("DiskSizeRequired", (contentDatabase.DiskSizeRequired / 1024 / 1024).ToString() + " MB"));
                contentDatabaseAttributes.Add(new AttributeValuePair("CurrentSiteCount", contentDatabase.CurrentSiteCount.ToString()));
                contentDatabaseAttributes.Add(new AttributeValuePair("WarningSiteCount", contentDatabase.WarningSiteCount.ToString()));
                contentDatabaseAttributes.Add(new AttributeValuePair("MaximumSiteCount", contentDatabase.MaximumSiteCount.ToString()));
                if (contentDatabase.Username != null)
                {
                    contentDatabaseAttributes.Add(new AttributeValuePair("Username", contentDatabase.Username));
                    contentDatabaseAttributes.Add(new AttributeValuePair("Password", contentDatabase.Password));
                }
            }
            catch (Exception e)
            {
                contentDatabaseAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return contentDatabaseAttributes;
        }

        public static List<AttributeValuePair> GetSPWebApplicationAttributes(SPWebApplication webApp)
        {
            List<AttributeValuePair> webAppAttributes = new List<AttributeValuePair>();
            try
            {
                webAppAttributes.Add(new AttributeValuePair("Name", webApp.Name));
                webAppAttributes.Add(new AttributeValuePair("Id", webApp.Id.ToString()));
            }
            catch (Exception e)
            {
                webAppAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return webAppAttributes;
        }
        public static List<AttributeValuePair> GetSPSolutionAttributes(SPSolution solution)
        {
            List<AttributeValuePair> solutionAttributes = new List<AttributeValuePair>();
            try
            {
                solutionAttributes.Add(new AttributeValuePair("Name", solution.Name));
                solutionAttributes.Add(new AttributeValuePair("Id", solution.Id.ToString()));
                solutionAttributes.Add(new AttributeValuePair("Status", solution.Status.ToString()));
                solutionAttributes.Add(new AttributeValuePair("Deployed", solution.Deployed.ToString()));
                solutionAttributes.Add(new AttributeValuePair("LastOperationEndTime", solution.LastOperationEndTime.ToLongDateString() + " " + solution.LastOperationEndTime.ToLongTimeString()));
                solutionAttributes.Add(new AttributeValuePair("DeploymentState", solution.DeploymentState.ToString()));
                solutionAttributes.Add(new AttributeValuePair("LastOperationDetails", solution.LastOperationDetails));
            }
            catch (Exception e)
            {
                solutionAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return solutionAttributes;
        }

        public static List<AttributeValuePair> GetSPApplicationPoolAttributes(SPApplicationPool appPool)
        {
            List<AttributeValuePair> appPoolAttributes = new List<AttributeValuePair>();
            try
            {
                appPoolAttributes.Add(new AttributeValuePair("Name", appPool.Name));
                appPoolAttributes.Add(new AttributeValuePair("UserName", appPool.Username));
                appPoolAttributes.Add(new AttributeValuePair("Password", appPool.Password));
                appPoolAttributes.Add(new AttributeValuePair("Id", appPool.Id.ToString()));
            }
            catch (Exception e)
            {
                appPoolAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return appPoolAttributes;
        }
        public static List<AttributeValuePair> GetSPJobDefinitionAttributes(SPJobDefinition jobDef)
        {
            List<AttributeValuePair> jobDefAttributes = new List<AttributeValuePair>();
            try
            {
                jobDefAttributes.Add(new AttributeValuePair("Title", jobDef.Title));
                jobDefAttributes.Add(new AttributeValuePair("DisplayName", jobDef.DisplayName));
                jobDefAttributes.Add(new AttributeValuePair("Name", jobDef.Name));
                jobDefAttributes.Add(new AttributeValuePair("Id", jobDef.ToString()));
                jobDefAttributes.Add(new AttributeValuePair("IsDisabled", jobDef.IsDisabled.ToString()));
                jobDefAttributes.Add(new AttributeValuePair("LastTimeRun", jobDef.LastRunTime.ToLongDateString() + " " + jobDef.LastRunTime.ToLongTimeString()));
                jobDefAttributes.Add(new AttributeValuePair("LockType", jobDef.LockType.ToString()));
                jobDefAttributes.Add(new AttributeValuePair("Retry", jobDef.Retry.ToString()));
                jobDefAttributes.Add(new AttributeValuePair("Status", jobDef.Status.ToString()));
                jobDefAttributes.Add(new AttributeValuePair("Schedule", jobDef.Schedule.ToString()));
            }
            catch (Exception e)
            {
                jobDefAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return jobDefAttributes;
        }
        public static List<AttributeValuePair> GetSPFeatureDefinitionAttributes(SPFeatureDefinition featureDef)
        {
            List<AttributeValuePair> featureDefAttributes = new List<AttributeValuePair>();
            try
            {
                featureDefAttributes.Add(new AttributeValuePair("DisplayName", featureDef.DisplayName));
                featureDefAttributes.Add(new AttributeValuePair("Id", featureDef.Id.ToString()));
                featureDefAttributes.Add(new AttributeValuePair("Scope", featureDef.Scope.ToString()));
                featureDefAttributes.Add(new AttributeValuePair("Hidden", featureDef.Hidden.ToString()));
                featureDefAttributes.Add(new AttributeValuePair("Name", featureDef.Name.ToString()));
                // suppress full feature path
                featureDefAttributes.Add(new AttributeValuePair("RootDirectory", featureDef.RootDirectory.Replace(SPUtility.GetGenericSetupPath(@"Template"), "")));
                featureDefAttributes.Add(new AttributeValuePair("Version", featureDef.Version.ToString()));
                featureDefAttributes.Add(new AttributeValuePair("Status", featureDef.Status.ToString()));
                featureDefAttributes.Add(new AttributeValuePair("SolutionId", featureDef.SolutionId.ToString()));
                if (featureDef.ReceiverAssembly != null)
                {
                    featureDefAttributes.Add(new AttributeValuePair("ReceiverAssembly", featureDef.ReceiverAssembly.ToString()));
                    featureDefAttributes.Add(new AttributeValuePair("ReceiverClass", featureDef.ReceiverClass.ToString()));
                }
            }
            catch (Exception e)
            {
                featureDefAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return featureDefAttributes;
        }
        public static List<AttributeValuePair> GetSPUserAttributes(SPUser user)
        {
            List<AttributeValuePair> userAttributes = new List<AttributeValuePair>();
            try
            {
                userAttributes.Add(new AttributeValuePair("Name", user.Name));
                userAttributes.Add(new AttributeValuePair("LoginName", user.LoginName));
                if (!string.IsNullOrEmpty(user.Email))
                {
                    userAttributes.Add(new AttributeValuePair("Email", user.Email));
                }
            }
            catch (Exception e)
            {
                userAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return userAttributes;
        }
        public static List<AttributeValuePair> GetSPSiteAttributes(SPSite site)
        {
            List<AttributeValuePair> siteAttributes = new List<AttributeValuePair>();
            try
            {
                siteAttributes.Add(new AttributeValuePair("Url", site.Url));
                siteAttributes.Add(new AttributeValuePair("WebApplication", site.WebApplication.Name));
                siteAttributes.Add(new AttributeValuePair("Id", site.ID.ToString()));
                siteAttributes.Add(new AttributeValuePair("WebCount", site.AllWebs.Count.ToString()));
            }
            catch (Exception e)
            {
                siteAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            } 
            return siteAttributes;
        }
        public static List<AttributeValuePair> GetSPSiteAdministrationAttributes(SPSiteAdministration siteAdmin)
        {
            List<AttributeValuePair> siteAdminAttributes = new List<AttributeValuePair>();
            try
            {
                siteAdminAttributes.Add(new AttributeValuePair("UsersCount", siteAdmin.UsersCount.ToString()));
                siteAdminAttributes.Add(new AttributeValuePair("Title", siteAdmin.Title));
                siteAdminAttributes.Add(new AttributeValuePair("Description", siteAdmin.Description));
                siteAdminAttributes.Add(new AttributeValuePair("Owner", siteAdmin.OwnerLoginName));
                siteAdminAttributes.Add(new AttributeValuePair("OwnerEmail", siteAdmin.OwnerEmail));
                siteAdminAttributes.Add(new AttributeValuePair("DiskUsed", (siteAdmin.DiskUsed / 1024 / 1024).ToString() + " MB"));
                siteAdminAttributes.Add(new AttributeValuePair("SecondaryContact", siteAdmin.SecondaryContactLoginName));
            }
            catch (Exception e)
            {
                siteAdminAttributes.Add(new AttributeValuePair("Exception", e.ToString()));
            }
            return siteAdminAttributes;
        }
    }
}
