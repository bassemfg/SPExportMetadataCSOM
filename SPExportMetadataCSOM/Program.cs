using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.SharePoint.Client;

namespace SPExportMetadataCSOM
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            //pass SharePoint Online credentials to get ClientContext object
            System.Security.SecureString securePassword = new System.Security.SecureString();// SecureString $PassWord

            #region constants
            string userName = "boxmigration@originmaterials.com";
            string password = "BoxShuttle718";
            string txtSiteURL =

        @"https://micromidasincorporated.sharepoint.com/sites/TestGroupAlex
        https://micromidasincorporated.sharepoint.com/sites/RyanandAndyShafer
        https://micromidasincorporated.sharepoint.com/sites/MariahsTasks
        https://micromidasincorporated.sharepoint.com/sites/Legal
        https://micromidasincorporated.sharepoint.com/sites/test2
        https://micromidasincorporated.sharepoint.com/sites/genadmin
        https://micromidasincorporated.sharepoint.com/sites/HTCRDTeam
        https://micromidasincorporated.sharepoint.com/sites/sharepoint
        https://micromidasincorporated.sharepoint.com/sites/Test
        https://micromidasincorporated.sharepoint.com/sites/InformationTechnology
        https://micromidasincorporated.sharepoint.com/sites/Chemistry10
        https://micromidasincorporated.sharepoint.com/sites/Chemistry85
        https://micromidasincorporated.sharepoint.com/sites/TeamSiteUsabilityUpdate
        https://micromidasincorporated.sharepoint.com/sites/FDCA
        https://micromidasincorporated.sharepoint.com/sites/Chemistry56
        https://micromidasincorporated.sharepoint.com/sites/accounting
        https://micromidasincorporated.sharepoint.com/sites/humanresources
        https://micromidasincorporated.sharepoint.com/sites/BDTeam
        https://micromidasincorporated.sharepoint.com/sites/engineering
        https://micromidasincorporated.sharepoint.com/sites/BusinessDevelopment-CanadianStrategies
        https://micromidasincorporated.sharepoint.com/sites/BICSharedSite
        https://micromidasincorporated.sharepoint.com/sites/Chemistry17
        https://micromidasincorporated.sharepoint.com/sites/ETInternshipFiles
        https://micromidasincorporated.sharepoint.com/sites/HTCFertilizerPellet
        https://micromidasincorporated.sharepoint.com/sites/accouting
        https://micromidasincorporated.sharepoint.com/sites/productdevelopment
        https://micromidasincorporated.sharepoint.com/sites/analytics
        https://micromidasincorporated.sharepoint.com/sites/TestGroup-Alex
        https://micromidasincorporated.sharepoint.com/sites/leads
        https://micromidasincorporated.sharepoint.com/sites/Surfactants
        https://micromidasincorporated.sharepoint.com/sites/WoodProducts
        https://micromidasincorporated.sharepoint.com/sites/ehs
        https://micromidasincorporated.sharepoint.com/sites/ActivatedCarbon
        https://micromidasincorporated.sharepoint.com/sites/pilot
        https://micromidasincorporated.sharepoint.com/sites/ExperimentVideos
        https://micromidasincorporated.sharepoint.com/sites/SellCarbonBlack
        https://micromidasincorporated.sharepoint.com/sites/chemistry
        https://micromidasincorporated.sharepoint.com/sites/HTCAg
        https://micromidasincorporated.sharepoint.com/sites/management
        https://micromidasincorporated.sharepoint.com/sites/EmailResponses
        https://micromidasincorporated.sharepoint.com/sites/construction
        https://micromidasincorporated.sharepoint.com/sites/CriticalProjects
        https://micromidasincorporated.sharepoint.com/sites/admins
        http://fundamental_donotdelete_7ab7c2c7-fd7d-44c5-83e4-ec9e804eea38/
        https://micromidasincorporated.sharepoint.com/sites/DoransCommunicationSiteTestBlank1
        https://micromidasincorporated-my.sharepoint.com/
        https://micromidasincorporated.sharepoint.com/portals/personal/rsmith
        https://micromidasincorporated.sharepoint.com/sites/CompliancePolicyCenter
        https://micromidasincorporated-admin.sharepoint.com/
        https://micromidasincorporated.sharepoint.com/search
        https://micromidasincorporated.sharepoint.com/portals/Community
        https://micromidasincorporated.sharepoint.com/sites/pwa
        https://micromidasincorporated.sharepoint.com/portals/hub
        https://micromidasincorporated.sharepoint.com/
        https://micromidasincorporated.sharepoint.com/sites/BoardRoom
        https://micromidasincorporated.sharepoint.com/sites/contentTypeHub";
            #endregion

            foreach (char c in password.ToCharArray())
                securePassword.AppendChar(c);
            SharePointOnlineCredentials spoCred = new SharePointOnlineCredentials(userName, securePassword);

            ClientContext ctx;

            foreach (string url in txtSiteURL.Split('\n'))
            {
                try
                {

                    ctx = new ClientContext(url.Replace('\r', '/').Trim());

                    ctx.Credentials = spoCred;
                    Web web = ctx.Web;
                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    //call the function that does the inventory of the site collection    
                    List<string> processedSites = new List<string>();
                    GetSPOSites(web, ctx, processedSites);
                }
                catch { }
            }
        }

        //CamlQuery camlQuery = new CamlQuery();

        private static List<ListItem> GetAllItems(ClientContext Context, List list, List<ListItem> ListItems, CamlQuery camlQuery)
        {
            //Create a CAML Query object
            //You can pass an undefined CamlQuery object to return all items from the list, or use the ViewXml property to define a CAML query and return items that meet specific criteria - https://msdn.microsoft.com/en-us/library/office/ee534956(v=office.14).aspx#sectionSection0
            //In this script an undefined CamlQuery object is passed, to get all list items 

            camlQuery.ViewXml =
    @"< View Scope = 'RecursiveAll'>
    < Query >
        <Where>
       </Where> <Gt><FieldRef Name='Modified'/><Value IncludeTimeValue='False' Type='DateTime'>" + new DateTime(2018, 9, 27).ToShortDateString() + @"</Value></Gt>

       <OrderBy>
            <FieldRef Name='Modified' />
        </OrderBy>
    </ Query >
</ View >";


            ListItemCollection AllItems = list.GetItems(camlQuery);
            Context.Load(AllItems);
            Context.ExecuteQuery();
            foreach (ListItem item in AllItems)
            {
                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    if (!ListItems.Contains(item))
                        ListItems.Add(item);
                }


                if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    camlQuery.FolderServerRelativeUrl = item.FieldValues["FileRef"].ToString();
                    GetAllItems(Context, list, ListItems, camlQuery);
                }


            }

            return ListItems;
        }

        private static void GetSPOSites(Web RootWeb, ClientContext Context, List<string> ProcessedSites)
        {
            string connStr = @"Data Source = SP2010VM\SHAREPOINT;Initial Catalog = OriginMaterials; Connect Timeout = 30; Encrypt=True;TrustServerCertificate=True;Authentication='Active Directory Integrated';ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

            string RootSiteCollections = "https://micromidasincorporated.sharepoint.com/";
            //int i = 0;
            int j = 0;
            StringBuilder sbFields = new StringBuilder();
            StringBuilder sbVals = new StringBuilder();


            //Create array variable to store data
            try
            {
                //get all webs under root web
                WebCollection Webs = RootWeb.Webs;

                Context.Load(Webs);
                Context.ExecuteQuery();
                //
                // loop through the webs
                List<Web> allWebs = Webs.ToList();
                allWebs.Insert(0, RootWeb);
                if (allWebs.Count > 0)
                    foreach (Web sWeb in allWebs)
                    {
                        if (!ProcessedSites.Contains(sWeb.Url))
                        {
                            try
                            {
                                //Console.WriteLine(sWeb.Url);
                                Console.WriteLine("Processing site: " + sWeb.Url);


                                var siteUrl = sWeb.Url;
                                ProcessedSites.Add(siteUrl);
                                //get all lists in web
                                ListCollection AllLists = sWeb.Lists;
                                Context.Load(AllLists);
                                Context.ExecuteQuery();
                                //loop through all lists in web
                                foreach (List list in AllLists)
                                {
                                    try
                                    {
                                        //Console.WriteLine("List: " + list.Title);
                                        Console.WriteLine("Processing list: " + list.Title);

                                        //get list title
                                        string listTitle = list.Title;

                                        if (list.BaseType == BaseType.DocumentLibrary && list.Hidden == false)
                                        {
                                            List<ListItem> AllItems = GetAllItems(Context, list, new List<ListItem>(), new CamlQuery());

                                            if (AllItems.Count > 0)
                                            {

                                                j = 0;

                                                foreach (ListItem item in AllItems)
                                                {
                                                    try
                                                    {
                                                        Context.Load(item);
                                                        Context.ExecuteQuery();

                                                        //Console.WriteLine("Processing item " + item.Id);
                                                        Console.WriteLine("Processing file: " + j + " of " + AllItems.Count + " in current list: " + list.Title);

                                                        //try catch block in the sql command to continue inserting all rows if one of the rows already exists
                                                        var text =
     @"BEGIN TRY 
INSERT INTO OriginMaterialsMetadata ([#FIELDS#]) VALUES('#VALUES#')
END TRY
BEGIN CATCH
END CATCH";

                                                        sbFields.Clear();
                                                        sbVals.Clear();
                                                        for (int ctrFields = 0; ctrFields < item.FieldValues.Count; ctrFields++)
                                                        {
                                                            try
                                                            {
                                                                string fieldKey = item.FieldValues.ElementAt(ctrFields).Key;
                                                                string fieldValue;
                                                                try { fieldValue = item.FieldValues.ElementAt(ctrFields).Value.ToString(); } catch { fieldValue = " "; }

                                                                //if (field.Hidden == false && field.Sealed == false)
                                                                {

                                                                    {
                                                                        sbFields.Append(fieldKey.ToString());
                                                                        sbFields.Append("],[");
                                                                    }

                                                                    if (!string.IsNullOrEmpty(fieldValue) && !fieldValue.Contains("/>") && !fieldValue.Contains("</"))
                                                                        sbVals.Append(fieldValue.Replace(",", "$^*").Replace(@"
", ""));
                                                                    else
                                                                        sbVals.Append(" ");

                                                                    sbVals.Append("','");

                                                                }
                                                            }
                                                            catch (Exception e) { System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue); }
                                                        }
                                                        //remove last ','
                                                        if (sbFields.Length > 0)
                                                            sbFields.Length = sbFields.Length - 3;
                                                        if (sbVals.Length > 0)
                                                            sbVals.Length = sbVals.Length - 3;

                                                        text = text.Replace("#FIELDS#", sbFields.ToString()).Replace("#VALUES#", sbVals.ToString());

                                                        j++;

                                                        using (System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection(connStr))
                                                        {
                                                            conn.Open();

                                                            using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(text, conn))
                                                            {
                                                                // Execute the command and log the # rows affected.
                                                                var rows = cmd.ExecuteNonQuery();
                                                                //log.Info($"{rows} rows were inserted");


                                                            }
                                                        }
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        Console.WriteLine("SPExport Exception Create", e.Message + 
                                                            "Trace" + e.StackTrace +"--inner "+ e.InnerException);

                                                    }                                                    }
                                                }

                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue);
                                    }
                                }
                            }
                            catch (Exception e) { System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue); }


                            GetSPOSites(sWeb, Context, ProcessedSites);
                        }
                    }
            }
            catch (Exception e) { System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue); }

        }
    }
}
