using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;


namespace SPExportMetadataCSOM
{
    public partial class Form1 : System.Windows.Forms.Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        string siteUrl;

        private void btnRun_Click(object sender, EventArgs e)

        {
            this.Cursor = Cursors.WaitCursor;
            //pass SharePoint Online credentials to get ClientContext object
            System.Security.SecureString securePassword = new System.Security.SecureString();// SecureString $PassWord
            string password = txtPwd.Text;
            string userName = txtUID.Text;
            foreach (char c in password.ToCharArray())
                securePassword.AppendChar(c);
            SharePointOnlineCredentials spoCred = new SharePointOnlineCredentials(userName, securePassword);
            
            ClientContext ctx = new ClientContext(txtSiteURL.Text);
            ctx.Credentials = spoCred;
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQuery();
            //call the function that does the inventory of the site collection    
            List<string> processedSites = new List<string>();
            GetSPOSites(web, ctx, processedSites);
            this.Cursor = Cursors.Default;
        }

        CamlQuery camlQuery = new CamlQuery();

        private List<ListItem> GetAllItems( ClientContext Context, List list, List<ListItem>  ListItems)
        {
            //Create a CAML Query object
            //You can pass an undefined CamlQuery object to return all items from the list, or use the ViewXml property to define a CAML query and return items that meet specific criteria - https://msdn.microsoft.com/en-us/library/office/ee534956(v=office.14).aspx#sectionSection0
            //In this script an undefined CamlQuery object is passed, to get all list items 
            
            camlQuery.ViewXml =
    @"< View Scope = 'RecursiveAll'>
    < Query >
        <Where>
<Gt><FieldRef Name='Modified'/><Value IncludeTimeValue='False' Type='DateTime'>" + new DateTime(2018, 8, 1).ToShortDateString() + @"</Value></Gt>
        </Where>
       
    </ Query >
</ View >";


            ListItemCollection AllItems = list.GetItems(camlQuery);
            Context.Load(AllItems);
            Context.ExecuteQuery();
            foreach (ListItem item in AllItems)
            {
                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    if(!ListItems.Contains(item))
                        ListItems.Add(item);
                }


                if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    camlQuery.FolderServerRelativeUrl = item.FieldValues["FileRef"].ToString();
                    GetAllItems(Context, list, ListItems);
                }


            }

            return ListItems;
        }

        private void GetSPOSites(Web RootWeb, ClientContext Context, List<string> ProcessedSites)
        {
            string RootSiteCollections = System.Configuration.ConfigurationSettings.AppSettings["RootSiteCollection"];
            //int i = 0;
            int j = 0;
            StringBuilder sbFields = new StringBuilder();
            StringBuilder sbVals = new StringBuilder();

            StreamWriter sw = null;

            //Create array variable to store data
            string itemTitle;
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
                                txtOutput.Text+="Processing site: " + sWeb.Url + @"
";

                                txtOutput.SelectionStart = txtOutput.Text.Length;
                                txtOutput.ScrollToCaret();

                                siteUrl = sWeb.Url;
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
                                        txtOutput.Text += "Processing list: " + list.Title + @"
";

                                        txtOutput.SelectionStart = txtOutput.Text.Length;
                                        txtOutput.ScrollToCaret();

                                        //get list title
                                        string listTitle = list.Title;

                                        if (list.BaseType == BaseType.DocumentLibrary && list.Hidden == false)
                                        {
                                            List<ListItem> AllItems = GetAllItems(Context,list,new List<ListItem>() );

                                            if (AllItems.Count > 0)
                                            {

                                                j = 0;

                                                foreach (ListItem item in AllItems)
                                                {
                                                    Context.Load(item);
                                                    Context.ExecuteQuery();

                                                    //Console.WriteLine("Processing item " + item.Id);
                                                    txtOutput.Text += "Processing file: " + j + " of " + AllItems.Count + " in current list: " + list.Title + @"
";
                                                    txtOutput.SelectionStart = txtOutput.Text.Length;
                                                    txtOutput.ScrollToCaret();
                                                 
                                                    for (int ctrFields = 0; ctrFields < item.FieldValues.Count; ctrFields++)
                                                    {
                                                        try
                                                        {
                                                            string fieldKey = item.FieldValues.ElementAt(ctrFields).Key;
                                                            string fieldValue;
                                                            try { fieldValue = item.FieldValues.ElementAt(ctrFields).Value.ToString(); } catch { fieldValue = " "; }

                                                            //if (field.Hidden == false && field.Sealed == false)
                                                            {
                                                                if (j == 0)
                                                                {
                                                                    sbFields.Append(fieldKey.ToString());
                                                                    sbFields.Append(',');
                                                                }

                                                                if (!string.IsNullOrEmpty(fieldValue) && !fieldValue.Contains("/>") && !fieldValue.Contains("</"))
                                                                    sbVals.Append(fieldValue.Replace(',', ' ').Replace(@"
", ""));
                                                                else
                                                                    sbVals.Append(" ");

                                                                sbVals.Append(',');

                                                            }
                                                        }
                                                        catch (Exception e) { System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue); }
                                                    }
                                                    //remove last ','
                                                    if (sbFields.Length > 0)
                                                        sbFields.Length = sbFields.Length - 1;
                                                    if (sbVals.Length > 0)
                                                        sbVals.Length = sbVals.Length - 1;
                                                    // add new lines
                                                    sbFields.Append(@"
"); sbVals.Append(@"
");
                                                    j++;

                                                }
                                            }
                                            sbFields.Append(sbVals.ToString());
                                            sbFields.Append(@"
");
                                            sbVals.Clear();
                                        }
                                    }
                                    catch (Exception e) {
                                        System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue);
                                    }
                                }
                            }
                            catch (Exception e) { System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue); }

                            sw = new StreamWriter(@"c:\test\metadata_" + sWeb.Url.Substring(sWeb.Url.LastIndexOf(@"/") + 1) + @".csv");

                            sw.Write(sbFields.ToString());
                            sw.Write(sbVals.ToString());
                            sw.Flush();
                            sw.Close();
                            sbFields.Clear();
                            sbVals.Clear();
                            GetSPOSites(sWeb, Context, ProcessedSites);
                        }
                    }
            }
            catch (Exception e) { System.Diagnostics.EventLog.WriteEntry("SPExport Exception Create", e.Message + "Trace" + e.StackTrace, System.Diagnostics.EventLogEntryType.Error, 121, short.MaxValue); }
        
        }
    }
}

