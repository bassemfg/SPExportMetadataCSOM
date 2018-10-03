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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)

        {
            //pass SharePoint Online credentials to get ClientContext object
            string securePassword = "";// SecureString $PassWord
            SharePointOnlineCredentials spoCred = new SharePointOnlineCredentials(UserName, securePassword);
            ClientContext ctx = new ClientContext(siteUrl);
            ctx.Credentials = spoCred;
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQuery();
            //call the function that does the inventory of the site collection    
            GetSPOSites(web, ctx);

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

            //get all webs under root web
           WebCollection Webs = RootWeb.Webs;
            Context.Load(Webs);
            Context.ExecuteQuery();
// loop through the webs
        foreach(Web sWeb in Webs)
        {
            Write - Host $sWeb.url
            $siteUrl = $sWeb.Url;
            #get all lists in web
            $AllLists = $sWeb.Lists
            $Context.Load($AllLists)
            $Context.ExecuteQuery()
            #loop through all lists in web
            ForEach($list in $AllLists){
                Write - Host List: $list.Title
              #get list title
              $listTitle = $list.Title;

# Do not inventory the following lists -> User Information List, Workflow History, Images, Site Assets, Composed Looks, Microfeed, Workflow Tasks, Access Requests, Master Page Gallery, Web Part Gallery, Style Library, List Template Library
# NOTE: Add to (or remove from) the list below, as needed.

                If($listTitle - ne 'User Information List' - and `
                 $listTitle - ne 'Workflow History' - and `
                 $listTitle - ne 'Images' - and `
                 $listTitle - ne 'Site Assets' - and `
                 $listTitle - ne 'Composed Looks' - and `
                 $listTitle - ne 'Microfeed' - and `
                 $listTitle - ne 'Workflow Tasks' - and `
                 $listTitle - ne 'Access Requests' - and `
                 $listTitle - ne 'Master Page Gallery' - and  `
                 $listTitle - ne 'Web Part Gallery' - and `
                 $listTitle - ne 'Style Library' - and `
                 $listTitle - ne 'List Template Library') {
                
                 #Create a CAML Query object
                 #You can pass an undefined CamlQuery object to return all items from the list, or use the ViewXml property to define a CAML query and return items that meet specific criteria - https://msdn.microsoft.com/en-us/library/office/ee534956(v=office.14).aspx#sectionSection0
                 #In this script an undefined CamlQuery object is passed, to get all list items 
                 $camlQuery = New - Object Microsoft.SharePoint.Client.CamlQuery
                  $AllItems = $list.GetItems($camlQuery)
                  $Context.Load($AllItems)
                  $Context.ExecuteQuery()
                  If($AllItems.Count - gt 0) {
                        ForEach($item in $AllItems){
                                                   
                          $listType = $list.BaseTemplate
                                $listUrl = $item["FileDirRef"]
                                #set item title based on the type of list
                          switch ($listType) 
                           {
                                101 { $itemTitle = $item["FileLeafRef"] }    #Document Library
                                103 { $itemTitle = $item["FileLeafRef"] }    #Links List
                                109 { $itemTitle = $item["FileLeafRef"] }    #Picture Library
                                119 { $itemTitle = $item["FileLeafRef"] }    #Site Pages
                                851 { $itemTitle = $item["FileLeafRef"] }    #Media
                                default { $itemTitle = $item["Title"] }
                            }
                            Write - Host Item Name: $itemTitle
                               
                          #retrieve item values
                          $itemType = $item.FileSystemObjectType
                                $itemurl = $item["FileRef"]
                                $itemCreatedBy = $item["Author"].LookupValue
                                $itemCreated = $item["Created"]
                                $itemModifiedBy = $item["Editor"].LookupValue
                                $itemModified = $item["Modified"]
                                #store the item values
                          #earlier versions (v2) of PowerShell do not support [Ordered]. If so, remove [Ordered]. Columns will be randomly ordered, but can be rearranged manually in the CSV file
                          $props = [Ordered]@{
                                'Site' = $siteUrl;
                                'List Title' = $listTitle;
                                'List URL' = $listUrl;
                                'List Type' = $listType;
                                'Item Title' = $itemTitle;
                                'Item URL' = $itemUrl;
                                'Item Type' = $itemType;
                                'Created By' = $itemCreatedBy;
                                'Created' = $itemCreated;
                                'Modified By' = $itemModifiedBy;
                                'Modified' = $itemModified};                      
                          #append the values to the existing array object 
                          $siteitemarray = New - Object - TypeName PSObject - Property $props; $siteitems += $siteitemarray
                         } #end loop for all items in list
                 } #check if item count is > 0
            } #check if it is a 'do not inventory' list          
          
          } #end loop for all lists in site 
          
          Get - SPOSites - RootWeb $sWeb - Context $Context #recursive call          
        } #end loop for all sites in site collection     
    
    #Output site collection inventory to CSV 
    $siteitems | Export-Csv $OutputFile -Append
    }

        }
    }
}

