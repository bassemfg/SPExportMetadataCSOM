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
            this.Cursor= Cursors.WaitCursor;
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
            GetSPOSites(web, ctx);
            this.Cursor = Cursors.Default;
        }

        private void GetSPOSites(Web RootWeb, ClientContext Context)
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
                if (Webs.Count > 0)
                    // loop through the webs
                    foreach (Web sWeb in Webs)
                    {
                        Console.WriteLine(sWeb.Url);
                        siteUrl = sWeb.Url;
                        //get all lists in web
                        ListCollection AllLists = sWeb.Lists;
                        Context.Load(AllLists);
                        Context.ExecuteQuery();
                        //loop through all lists in web
                        foreach (List list in AllLists)
                        {
                            Console.WriteLine("List: " + list.Title);
                            //get list title
                            string listTitle = list.Title;

                            if (list.BaseType == BaseType.DocumentLibrary && list.Hidden == false)
                            {

                                //Create a CAML Query object
                                //You can pass an undefined CamlQuery object to return all items from the list, or use the ViewXml property to define a CAML query and return items that meet specific criteria - https://msdn.microsoft.com/en-us/library/office/ee534956(v=office.14).aspx#sectionSection0
                                //In this script an undefined CamlQuery object is passed, to get all list items 
                                CamlQuery camlQuery = new CamlQuery();
                                ListItemCollection AllItems = list.GetItems(camlQuery);
                                Context.Load(AllItems);
                                Context.ExecuteQuery();
                                if (AllItems.Count > 0)
                                {

                                    j = 0;

                                    foreach (ListItem item in AllItems)
                                    {
                                        Context.Load(item);
                                        Context.ExecuteQuery();

                                        Console.WriteLine("Processing item " + item.Id);
                                        /*
                                    if (j == 0)
                                    {
                                        sbFields.Append("SourcePath");
                                        sbFields.Append(',');
                                        sbFields.Append("UniqueId");
                                        sbFields.Append(',');
                                        sbFields.Append("SiteURL");
                                        sbFields.Append(',');

                                    }

                                    sbVals.Append(item["URL Path"].ToString().Replace(',', ' '));
                                    sbVals.Append(',');

                                    sbVals.Append(item["UniqueId"].ToString());
                                    sbVals.Append(',');


                                    sbVals.Append(sWeb.Url);
                                    sbVals.Append(',');
                                    */
                                        for (int ctrFields = 0; ctrFields < item.FieldValues.Count; ctrFields++)
                                        {
                                            try
                                            {
                                                string fieldKey = item.FieldValues.ElementAt(ctrFields).Key;
                                                string fieldValue = item.FieldValues.ElementAt(ctrFields).Value.ToString();

                                                //if (field.Hidden == false && field.Sealed == false)
                                                {
                                                    if (j == 0)
                                                    {
                                                        sbFields.Append(fieldKey.ToString());
                                                        sbFields.Append(',');
                                                    }

                                                    if (!string.IsNullOrEmpty(fieldValue) && !fieldValue.Contains("/>"))
                                                        sbVals.Append(fieldValue.Replace(',', ' ').Replace(@"
",""));
                                                    else
                                                        sbVals.Append(" ");

                                                    sbVals.Append(',');

                                                }
                                            }
                                            catch { }
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
                        sw = new StreamWriter(@"c:\test\metadata_" + sWeb.Url.Substring(sWeb.Url.LastIndexOf(@"/") + 1) + @".csv");

                        sw.Write(sbFields.ToString());
                        sw.Write(sbVals.ToString());
                        sw.Flush();
                        sw.Close();
                        sbFields.Clear();
                        sbVals.Clear();
                        GetSPOSites(sWeb, Context);

                    }
            }
            catch (Exception e){ }

        }
    }
}

