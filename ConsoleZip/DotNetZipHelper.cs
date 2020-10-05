using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleZip
{
    public class DotNetZipHelper
    {
        /// <summary>
        /// 壓縮單個檔案成zip檔
        /// </summary>
        /// <param name="zipFilePathList">檔案路徑</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void ZipSingleFile(string zipFilePath, string zipSavaPath, string password = null, Encoding enc = null)
        {
            if (File.Exists(zipSavaPath))
                File.Delete(zipSavaPath);

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddFile(zipFilePath);

                zip.Save(zipSavaPath);
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮單個檔案成zip檔(Stream)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFilePath">檔案路徑</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static Stream ZipSingleFileStream(string zipFilePath, string password = null, Encoding enc = null)
        {
            MemoryStream ms = new MemoryStream();

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddFile(zipFilePath);

                zip.Save(ms);
            }
            // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// 壓縮單個檔案成zip檔
        /// 要壓縮的檔案不落地(來源)
        /// </summary>
        /// <param name="zipFile">壓縮檔案Byte</param>
        /// <param name="fileName">檔案名稱(含副檔名)</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void ZipSingleFile(byte[] zipFile, string fileName, string zipSavaPath, string password = null, Encoding enc = null)
        {
            if (File.Exists(zipSavaPath))
                File.Delete(zipSavaPath);

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddEntry(fileName, zipFile);

                zip.Save(zipSavaPath);
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮單個檔案成zip檔(Stream)
        /// 要壓縮的檔案不落地(來源)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFile">壓縮檔案Byte</param>
        /// <param name="fileName">檔案名稱(含副檔名)</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static Stream ZipSingleFileStream(byte[] zipFile, string fileName, string password = null, Encoding enc = null)
        {
            MemoryStream ms = new MemoryStream();

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddEntry(fileName, zipFile);

                zip.Save(ms);
            }
            // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// 壓縮多個檔案成zip檔
        /// </summary>
        /// <param name="zipFilePathList">要壓縮的檔案資料集</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void ZipMultipleFile(IEnumerable<string> zipFilePathList, string zipSavaPath, string password = null, Encoding enc = null)
        {
            if (File.Exists(zipSavaPath))
                File.Delete(zipSavaPath);

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddFiles(zipFilePathList);

                zip.Save(zipSavaPath);
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮多個檔案成zip檔(Stream)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFilePathList">要壓縮的檔案資料集</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static Stream ZipMultipleFileStream(IEnumerable<string> zipFilePathList, string password = null, Encoding enc = null)
        {
            MemoryStream ms = new MemoryStream();

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddFiles(zipFilePathList);

                zip.Save(ms);
            }
            // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// 壓縮多個檔案成zip檔
        /// 要壓縮的檔案不落地(來源)
        /// </summary>
        /// <param name="zipFileList">壓縮檔案Byte</param>
        /// <param name="fileName">檔案名稱(含副檔名)</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void ZipMultipleFile(List<byte[]> zipFileList,List<string> fileName, string zipSavaPath, string password = null, Encoding enc = null)
        {
            if (File.Exists(zipSavaPath))
                File.Delete(zipSavaPath);

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                for (int i = 0; i < zipFileList.Count; i++)
                    zip.AddEntry(fileName[i], fileName[i]);

                zip.Save(zipSavaPath);
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮多個檔案成zip檔(Stream)
        /// 要壓縮的檔案不落地(來源)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFileList"></param>
        /// <param name="fileName"></param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static Stream ZipMultipleFileStream(List<byte[]> zipFileList, List<string> fileName, string password = null, Encoding enc = null)
        {
            MemoryStream ms = new MemoryStream();

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                for (int i = 0; i < zipFileList.Count; i++)
                    zip.AddEntry(fileName[i], fileName[i]);

                zip.Save(ms);
            }

            // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// 壓縮"目錄"成zip檔
        /// </summary>
        /// <param name="zipFolderPath">資料夾路徑</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void ZipFromDirectory(string zipFolderPath, string zipSavaPath, string password = null, Encoding enc = null)
        {
            if (File.Exists(zipSavaPath))
                File.Delete(zipSavaPath);

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddDirectory(zipFolderPath);

                zip.Save(zipSavaPath);
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮"目錄"成zip檔(Stream)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFolderPath">資料夾路徑</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static Stream ZipFromDirectoryStream(string zipFolderPath, string password = null, Encoding enc = null)
        {
            MemoryStream ms = new MemoryStream();

            using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
            {
                zip.Password = password;

                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                zip.AddDirectory(zipFolderPath);

                zip.Save(ms);
            }

            // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// 解壓縮zip檔到目錄
        /// </summary>
        /// <param name="zipFile">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipToFolder">解壓縮資料夾位置</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void UnZipToDirectory(string zipFile, string zipToFolder, string password = null, Encoding enc = null)
        {
            if (!Directory.Exists(zipToFolder))
                Directory.CreateDirectory(zipToFolder);

            ReadOptions opt = new ReadOptions();

            opt.Encoding = enc ?? Encoding.Default; // 設為localhost的OS的語係。用以解決檔名中文編碼的問題。

            ZipFile zip = ZipFile.Read(zipFile, opt);

            foreach (ZipEntry entry in zip.Entries)
            {
                // if(entry.FileName == "target_file_name") // 與目的檔名相同才解壓縮
                entry.ExtractExistingFile = ExtractExistingFileAction.OverwriteSilently;
                entry.ExtractWithPassword(zipToFolder, password);
            }
        }

        /// <summary>
        /// 解壓縮zip檔到目錄(Stream)
        /// </summary>
        /// <param name="zipStream">讀取壓縮檔的Stream</param>
        /// <param name="zipToFolder">解壓縮資料夾位置</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static void UnZipToDirectory(Stream zipStream, string zipToFolder, string password = null, Encoding enc = null)
        {
            if (!Directory.Exists(zipToFolder))
                Directory.CreateDirectory(zipToFolder);

            ReadOptions opt = new ReadOptions();

            opt.Encoding = enc ?? Encoding.Default; // 設為localhost的OS的語係。用以解決檔名中文編碼的問題。

            using (ZipFile zip = ZipFile.Read(zipStream, opt))
            {
                foreach (ZipEntry entry in zip.Entries)
                {
                    // if(entry.FileName == "target_file_name") // 與目的檔名相同才解壓縮
                    entry.ExtractExistingFile = ExtractExistingFileAction.OverwriteSilently;
                    entry.ExtractWithPassword(zipToFolder, password);
                }
            }
        }
    }
}
