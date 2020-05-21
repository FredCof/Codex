using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Web;

namespace Codex.Word.Net.Base
{
    public class BasePackage
    {
        #region Packaging Method

        /// <summary>
        /// Add a folder along with its subfolders to a Package
        /// </summary>
        /// <param name="folderName">The folder to add</param>
        /// <param name="compressedFileName">The package to create</param>
        /// <param name="overrideExisting">Override exsisitng files</param>
        /// <returns></returns>
        public static bool PackageFolder(string folderName, string compressedFileName, bool overrideExisting)
        {
            if (folderName.EndsWith(@"\"))
                folderName = folderName.Remove(folderName.Length - 1);
            bool result = false;
            if (!Directory.Exists(folderName))
            {
                return result;
            }

            if (!overrideExisting && File.Exists(compressedFileName))
            {
                return result;
            }
            try
            {
                using (Package package = Package.Open(compressedFileName, FileMode.Create))
                {
                    var fileList = Directory.EnumerateFiles(folderName, "*", SearchOption.AllDirectories);
                    foreach (string fileName in fileList)
                    {
                       
                        //The path in the package is all of the subfolders after folderName
                        string pathInPackage;
                        pathInPackage = Path.GetDirectoryName(fileName).Replace(folderName, string.Empty) + "/" + Path.GetFileName(fileName);

                        Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri(pathInPackage, UriKind.Relative));
                        PackagePart packagePartDocument = package.CreatePart(partUriDocument,"", CompressionOption.Maximum);
                        using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                        {
                            fileStream.CopyTo(packagePartDocument.GetStream());
                        }
                    }
                }
                result = true;
            }
            catch (Exception e)
            {
                throw new Exception("Error zipping folder " + folderName, e);
            }
           
            return result;
        }
        
        /// <summary>
        /// Compress a file into a ZIP archive as the container store
        /// </summary>
        /// <param name="fileName">The file to compress</param>
        /// <param name="compressedFileName">The archive file</param>
        /// <param name="overrideExisting">override existing file</param>
        /// <returns></returns>
        public static bool PackageFile(string fileName, string compressedFileName, bool overrideExisting)
        {
            bool result = false;

            if (!File.Exists(fileName))
            {
                return result;
            }

            if (!overrideExisting && File.Exists(compressedFileName))
            {
                return result;
            }

            try
            {
                Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri(Path.GetFileName(fileName), UriKind.Relative));
               
                using (Package package = Package.Open(compressedFileName, FileMode.OpenOrCreate))
                {
                    if (package.PartExists(partUriDocument))
                    {
                        package.DeletePart(partUriDocument);
                    } 

                    PackagePart packagePartDocument = package.CreatePart(partUriDocument, "", CompressionOption.Maximum);
                    using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                    {
                        fileStream.CopyTo(packagePartDocument.GetStream());
                    }
                }
                result = true;
            }
            catch (Exception e)
            {
                throw new Exception("Error zipping file " + fileName, e);
            }
            
            return result;
        }
        
        /// <summary>
        /// Extract a container Zip. NOTE: container must be created as Open Packaging Conventions (OPC) specification
        /// </summary>
        /// <param name="folderName">The folder to extract the package to</param>
        /// <param name="compressedFileName">The package file</param>
        /// <param name="overrideExisting">override existing files</param>
        /// <returns></returns>
        public static bool UncompressFile(string folderName, string compressedFileName, bool overrideExisting)
        {
            bool result = false;
            try
            {
                if (!File.Exists(compressedFileName))
                {
                    return result;
                }

                DirectoryInfo directoryInfo = new DirectoryInfo(folderName);
                if (!directoryInfo.Exists)
                    directoryInfo.Create();

                using (Package package = Package.Open(compressedFileName, FileMode.Open, FileAccess.Read))
                {
                    foreach (PackagePart packagePart in package.GetParts())
                    {
                        ExtractPart(packagePart, folderName, overrideExisting);
                    }
                }

                result = true;
            }
            catch (Exception e)
            {
                throw new Exception("Error unzipping file " + compressedFileName, e);
            }
            
            return result;
        }

        static void ExtractPart(PackagePart packagePart, string targetDirectory, bool overrideExisting)
        {
            string stringPart = targetDirectory + HttpUtility.UrlDecode(packagePart.Uri.ToString()).Replace('\\', '/');

            if (!Directory.Exists(Path.GetDirectoryName(stringPart)))
                Directory.CreateDirectory(Path.GetDirectoryName(stringPart));

            if (!overrideExisting && File.Exists(stringPart))
                return;
            using (FileStream fileStream = new FileStream(stringPart, FileMode.Create))
            {
                packagePart.GetStream().CopyTo(fileStream);
            }
        }

        #endregion

        #region Compression Method

        /// <summary>
        /// Using System.IO.Compress to Pack files.
        /// </summary>
        /// <param name="startPath">Compress files from startPath, which will contains its all sub fold and files</param>
        /// <param name="zipPath">The compress to target</param>
        /// <returns>If compress successful</returns>
        public static bool Compress(string startPath, string zipPath)
        {
            try
            {
                ZipFile.CreateFromDirectory(startPath, zipPath);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Using System.IO.Compress to Unpack Zip file
        /// </summary>
        /// <param name="zipPath">Zip file to decompress.</param>
        /// <param name="extractPath">The path to decompress file.</param>
        /// <returns></returns>
        public static bool Decompress(string zipPath, string extractPath)
        {
            try
            {
                ZipFile.ExtractToDirectory(zipPath, extractPath);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        #endregion
    }
}