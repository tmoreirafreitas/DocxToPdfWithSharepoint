using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using Microsoft.Office.Interop.Word;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace SPDocxToPdf.DocxToPdf
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class DocxToPdf : SPItemEventReceiver
    {
        public Document wordDocument { get; set; }
        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
        }

        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    var spItem = properties.ListItem;
                    if (bool.Parse(spItem["Converter"].ToString()))
                    {
                        var spFile = spItem.File;
                        var pathSave = @"C:\Temp";
                        Download(spFile, pathSave);
                        ConvertToPDF(pathSave);
                        var filePdf = GetAllFile(pathSave).Single(f => Path.GetExtension(f).ToLower().Equals(".pdf") &&
                            f.Replace(Path.GetExtension(f), "").ToLower() == spFile.Name.Replace(Path.GetExtension(spFile.Name), "").ToLower());
                        Upload(filePdf, properties.List.Title, properties.OpenSite());
                    }
                });
            }
            catch (SPException ex)
            {
                throw ex;
            }
        }

        private void Download(SPFile spFile, string pathSave)
        {
            using (var result = spFile.OpenBinaryStream())
            {
                using (var ms = new MemoryStream())
                {
                    result.CopyTo(ms);
                    //Check if the directory exists
                    if (!Directory.Exists(pathSave))
                    {
                        Directory.CreateDirectory(pathSave);
                    }
                    pathSave = string.Format(@"{0}\{1}", pathSave.TrimEnd('\\'), spFile.Name);
                    using (FileStream fStream = new FileStream(pathSave, FileMode.Create))
                        ms.WriteTo(fStream);
                }
            }
        }

        private void Upload(string fullPath, string documentLibraryName, SPSite spSite)
        {
            using (spSite)
            {
                using (var spWeb = spSite.OpenWeb())
                {
                    if (!File.Exists(fullPath))
                    {
                        throw new FileNotFoundException("File not found.", fullPath);
                    }

                    SPFolder myLibrary = spWeb.Folders[documentLibraryName];
                    bool replaceExistingFiles = true;
                    string fileName = Path.GetFileName(fullPath);
                    FileStream fileStream = File.OpenRead(fullPath);
                    SPFile spfile = myLibrary.Files.Add(fileName, fileStream, replaceExistingFiles);
                    myLibrary.Update();
                }
            }
        }

        private List<string> GetAllFile(string targetDirectory)
        {
            List<string> AllFile = new List<string>();
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
            {
                AllFile.Add(Path.GetFullPath(fileName));
            }

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);

            return AllFile;
        }

        private void ConvertToPDF(string path)
        {
            var missing = System.Reflection.Missing.Value;
            Application appWord = new Application();
            wordDocument = appWord.Documents.Open(path);
            var diretorio = Path.GetDirectoryName(path);
            var fileName = Path.GetFileName(path);
            wordDocument.ExportAsFixedFormat(string.Format(@"{0}\{1}.pdf", diretorio, fileName), WdExportFormat.wdExportFormatPDF,
                false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument);
            wordDocument.Close();
            appWord.Quit();
        }
    }
}