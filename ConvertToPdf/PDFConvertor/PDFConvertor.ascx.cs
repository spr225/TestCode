using Microsoft.SharePoint;
using System;
using System.ComponentModel;
using System.IO;
using System.Web.UI.WebControls.WebParts;
using Microsoft.Office.Word.Server.Conversions;
//using Microsoft.Office.Server.PowerPoint.Conversion;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ConvertToPdf.Common;

namespace ConvertToPdf.PDFConvertor
{
    [ToolboxItemAttribute(false)]
    public partial class PDFConvertor : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling on a farm solution
        // using the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public PDFConvertor()
        {
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            InitializeControl();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
        }

        public static string WORD_AUTOMATION_SERVICE = "WAS";

        protected void btnWordConvertToPDF_Click(object sender, EventArgs e)
        {
            var site  = SPContext.Current.Site;
           
            SPList library = CommonUtilities.isValidLibrary(txtLibrary.Text);

            if(library == null)
            {
                ltResult.Text = "Please ensure you enter a name of a valid Document Library."; 
                return;
            }

           ltResult.Text = WordDocsToConvertToPdf(library);


        }

        private String WordDocsToConvertToPdf(SPList library)
        {
            StringBuilder logResult = new StringBuilder();
            SPQuery query = new SPQuery();
            query.Folder = library.RootFolder;
            //query.ViewAttributes = "Scope=\"Recursive\"";
            query.ViewXml = @"<View Scope='Recursive'>
                                <Query>
                                   <Where>
                                        <Or>
										    <Contains>
											    <FieldRef Name='File_x0020_Type'/>
											    <Value Type='Text'>doc</Value>
										    </Contains>
										    <Contains>
											    <FieldRef Name='File_x0020_Type'/>
											    <Value Type='Text'>docx</Value>
										    </Contains>
									    </Or>
                                    </Where>
                                </Query>
                            </View>";

            SPListItemCollection listItems = library.GetItems(query);

            if (listItems.Count > 0)
            {
                foreach (SPListItem li in listItems)
                {
                    using (MemoryStream destinationStream = new MemoryStream())
                    {
                        SyncConverter sc = new SyncConverter(WORD_AUTOMATION_SERVICE);
                        sc.UserToken = SPContext.Current.Site.UserToken;
                        sc.Settings.UpdateFields = true;
                        sc.Settings.OutputFormat = SaveFormat.PDF;

                        //Convert to PDF
                        ConversionItemInfo info = sc.Convert(li.File.OpenBinaryStream(), destinationStream);
                        var filename = Path.GetFileNameWithoutExtension(li.File.Name) + ".pdf";
                        if (info.Succeeded)
                        {
                            
                            SPFile newfile = library.RootFolder.Files.Add(filename, destinationStream, true);
                            logResult.AppendLine("Successfully converted " + li.File.Name);
                        }
                        else if (info.Failed)
                        {
                            //http://www.ilikesharepoint.de/2014/07/sharepoint-word-automation-service-does-not-work-file-may-be-corrupted/
                            logResult.AppendLine("Error converted file: " + li.File.Name + info.ErrorMessage);
                        }
                     }
                }
            }

            return logResult.ToString();
        }

        
        private SPFolder GetFolder(SPFolder parentFolder, string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentException("folderName");
            }

            var folderCollection = parentFolder.SubFolders;
            var folder = GetFolderImplementation(folderCollection, folderName);
            return folder;
        }

        private SPFolder GetFolderImplementation(SPFolderCollection folderCollection, string folderName)
        {
            if (folderCollection == null)
            {
                throw new ArgumentNullException("folderCollection");
            }
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentException("folderName");
            }

            foreach (SPFolder f in folderCollection)
            {
                if (f.Name.Equals(folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return f;
                }
            }
            return null;
        }

        private bool DoesFolderExist(SPFolder parentFolder, string folderName)
        {
           if(string.IsNullOrEmpty(folderName))
           {
               throw new ArgumentException("folderName");
           }

           var folderCollection = parentFolder.SubFolders;
           var exists = FolderExistsImplementation(folderCollection, folderName);
           return exists;
        }

        private bool FolderExistsImplementation(SPFolderCollection folderCollection, string folderName)
        {
           if(folderCollection == null)
           {
               throw new ArgumentNullException("folderCollection");
           }
           if(string.IsNullOrEmpty(folderName))
           {
               throw new ArgumentException("folderName");
           }

            foreach(SPFolder f in folderCollection)
            {
                if(f.Name.Equals(folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        
    }
}
