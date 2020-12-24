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
        #region 單檔壓縮

        /// <summary>
        /// 壓縮單個檔案成zip檔
        /// </summary>
        /// <param name="zipFilePath">檔案路徑</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static ZipExecuteResult ZipSingleFile(string zipFilePath, string zipSavaPath, string zipFolderName = "", string password = null, Encoding enc = null)
        {
            try
            {
                if (!File.Exists(zipFilePath))
                    return ZipExecuteResult.Fail(string.Format("來源路徑:{0} 找不到該檔案。", zipFilePath));

                //原本壓縮檔存在的話就將其移除
                if (File.Exists(zipSavaPath))
                    File.Delete(zipSavaPath);

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddFile(zipFilePath, zipFolderName);

                    zip.Save(zipSavaPath);

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }

        /// <summary>
        /// 加入檔案至壓縮檔裡面
        /// </summary>
        /// <param name="zipFilePath">檔案路徑</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">針對新增的檔案加入解壓縮密碼</param>
        public static ZipExecuteResult ZipSingleFileToZip(string zipFilePath, string zipSavaPath, string zipFolderName = "", string password = null)
        {
            try
            {
                if(!File.Exists(zipFilePath))
                    return ZipExecuteResult.Fail(string.Format("來源路徑:{0} 找不到該檔案。", zipFilePath));

                //原本壓縮檔存在的話就將其移除
                if (!File.Exists(zipSavaPath))
                    return ZipExecuteResult.Fail("來源壓縮檔案路徑找不到該檔案。");

                using (ZipFile zip = ZipFile.Read(zipSavaPath))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    foreach (var item in zip.EntryFileNames)
                        if (Path.Combine(item.Split('/')) == Path.Combine(zipFolderName, Path.GetFileName(zipFilePath)))
                            return ZipExecuteResult.Fail("檔案已存在");
                    
                    zip.AddFile(zipFilePath, zipFolderName);

                    zip.Save(zipSavaPath);

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }


        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮單個檔案成zip檔(Stream)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFilePath">檔案路徑</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static ZipExecuteResult<Stream> ZipSingleFileStream(string zipFilePath, string zipFolderName = "", string password = null, Encoding enc = null)
        {
            try
            {
                if (!File.Exists(zipFilePath))
                    return ZipExecuteResult<Stream>.Fail(string.Format("來源路徑:{0} 找不到該檔案。", zipFilePath));

                MemoryStream ms = new MemoryStream();

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddFile(zipFilePath, zipFolderName);

                    zip.Save(ms);
                }
                // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
                ms.Seek(0, SeekOrigin.Begin);

                return ZipExecuteResult<Stream>.Ok(ms);
            }
            catch (Exception ex)
            {
                return ZipExecuteResult<Stream>.Fail(ex.ToString());
            }
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
        public static ZipExecuteResult ZipSingleFile(byte[] zipFile, string fileName, string zipSavaPath, string password = null, Encoding enc = null)
        {
            try
            {
                //原本壓縮檔存在的話就將其移除
                if (File.Exists(zipSavaPath))
                    File.Delete(zipSavaPath);

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddEntry(fileName, zipFile);

                    zip.Save(zipSavaPath);

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
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
        public static ZipExecuteResult<Stream> ZipSingleFileStream(byte[] zipFile, string fileName, string password = null, Encoding enc = null)
        {
            try
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

                return ZipExecuteResult<Stream>.Ok(ms);
            }
            catch (Exception ex)
            {
                return ZipExecuteResult<Stream>.Fail(ex.ToString());
            }
        }

        #endregion

        #region 多檔案壓縮

        /// <summary>
        /// 壓縮多個檔案成zip檔
        /// </summary>
        /// <param name="zipFilePathList">要壓縮的檔案資料集</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static ZipExecuteResult ZipMultipleFile(IEnumerable<string> zipFilePathList, string zipSavaPath, string zipFolderName = "", string password = null, Encoding enc = null)
        {
            try
            {
                foreach (var item in zipFilePathList)
                    if (!File.Exists(item))
                        return ZipExecuteResult.Fail(string.Format("來源路徑:{0} 找不到該檔案。", item));

                //原本壓縮檔存在的話就將其移除
                if (File.Exists(zipSavaPath))
                    File.Delete(zipSavaPath);

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddFiles(zipFilePathList, zipFolderName);

                    zip.Save(zipSavaPath);

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }

        /// <summary>
        /// 加入多個檔案至壓縮檔裡面
        /// </summary>
        /// <param name="zipFilePathList">要壓縮的檔案資料集</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <returns></returns>
        public static ZipExecuteResult ZipMultipleFileToZip(IEnumerable<string> zipFilePathList, string zipSavaPath, string zipFolderName = "", string password = null)
        {
            try
            {
                foreach (var item in zipFilePathList)
                    if (!File.Exists(item))
                        return ZipExecuteResult.Fail(string.Format("來源路徑:{0} 找不到該檔案。", item));

                //原本壓縮檔存在的話就將其移除
                if (!File.Exists(zipSavaPath))
                    return ZipExecuteResult.Fail("來源壓縮檔案路徑找不到該檔案。");

                using (ZipFile zip = ZipFile.Read(zipSavaPath))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    foreach (var item in zip.EntryFileNames)
                    {
                        foreach (var fileNameItem in zipFilePathList)
                            if (Path.Combine(item.Split('/')) == Path.Combine(fileNameItem.Split('/')))
                                return ZipExecuteResult.Fail("檔案已存在");
                    }

                    zip.AddFiles(zipFilePathList, zipFolderName);

                    zip.Save(zipSavaPath);

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮多個檔案成zip檔(Stream)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFilePathList">要壓縮的檔案資料集</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static ZipExecuteResult<Stream> ZipMultipleFileStream(IEnumerable<string> zipFilePathList, string zipFolderName = "", string password = null, Encoding enc = null)
        {
            try
            {
                foreach (var item in zipFilePathList)
                    if (!File.Exists(item))
                        return ZipExecuteResult<Stream>.Fail(string.Format("來源路徑:{0} 找不到該檔案。", item));

                MemoryStream ms = new MemoryStream();

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddFiles(zipFilePathList, zipFolderName);

                    zip.Save(ms);
                }
                // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
                ms.Seek(0, SeekOrigin.Begin);

                return ZipExecuteResult<Stream>.Ok(ms);
            }
            catch (Exception ex)
            {
                return ZipExecuteResult<Stream>.Fail(ex.ToString());
            }
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
        public static ZipExecuteResult ZipMultipleFile(List<byte[]> zipFileList,List<string> fileName, string zipSavaPath, string password = null, Encoding enc = null)
        {
            try
            {
                //原本壓縮檔存在的話就將其移除
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

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult<Stream>.Fail(ex.ToString());
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
        public static ZipExecuteResult<Stream> ZipMultipleFileStream(List<byte[]> zipFileList, List<string> fileName, string password = null, Encoding enc = null)
        {
            try
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

                return ZipExecuteResult<Stream>.Ok(ms);
            }
            catch (Exception ex)
            {
                return ZipExecuteResult<Stream>.Fail(ex.ToString());
            }
        }

        /// <summary>
        /// 壓縮"目錄"成zip檔
        /// </summary>
        /// <param name="zipFolderPath">資料夾路徑</param>
        /// <param name="zipSavaPath">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static ZipExecuteResult ZipFromDirectory(string zipFolderPath, string zipSavaPath, string zipFolderName = "", string password = null, Encoding enc = null)
        {
            try
            {
                if (!Directory.Exists(zipFolderPath))
                    return ZipExecuteResult.Fail(string.Format("來源路徑:{0} 的資料夾不存在。", zipFolderPath));

                //原本壓縮檔存在的話就將其移除
                if (File.Exists(zipSavaPath))
                    File.Delete(zipSavaPath);

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddDirectory(zipFolderPath, zipFolderName);

                    zip.Save(zipSavaPath);

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }

        /// <summary>
        /// 檔案不可以太大(1GB以內)
        /// 壓縮"目錄"成zip檔(Stream)
        /// 壓縮檔不落地(目標)
        /// </summary>
        /// <param name="zipFolderPath">資料夾路徑</param>
        /// <param name="zipFolderName">壓縮後裡面的資料夾名稱，可以傳空字串</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static ZipExecuteResult<Stream> ZipFromDirectoryStream(string zipFolderPath, string zipFolderName = "", string password = null, Encoding enc = null)
        {
            try
            {
                if (!Directory.Exists(zipFolderPath))
                    return ZipExecuteResult<Stream>.Fail(string.Format("來源路徑:{0} 的資料夾不存在。", zipFolderPath));

                MemoryStream ms = new MemoryStream();

                using (ZipFile zip = new ZipFile(enc ?? Encoding.Default))
                {
                    zip.Password = password;

                    zip.UseZip64WhenSaving = Zip64Option.AsNecessary;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                    zip.AddDirectory(zipFolderPath, zipFolderName);

                    zip.Save(ms);
                }

                // 把檔案讀寫游標移到第一位，以利於後面續輸出成實體檔
                ms.Seek(0, SeekOrigin.Begin);

                return ZipExecuteResult<Stream>.Ok(ms);
            }
            catch (Exception ex)
            {
                return ZipExecuteResult<Stream>.Fail(ex.ToString());
            }
        }

        #endregion

        #region 解壓縮

        /// <summary>
        /// 解壓縮zip檔到目錄
        /// </summary>
        /// <param name="zipFile">壓縮檔路徑含檔名、副檔名</param>
        /// <param name="zipToFolder">解壓縮資料夾位置</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static ZipExecuteResult UnZipToDirectory(string zipFile, string zipToFolder, string password = null, Encoding enc = null)
        {
            try
            {
                if (!File.Exists(zipFile))
                    return ZipExecuteResult.Fail(string.Format("來源路徑:{0} 找不到該壓縮檔案。", zipFile));

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

                return ZipExecuteResult.Ok();
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }

        /// <summary>
        /// 解壓縮zip檔到目錄(Stream)
        /// </summary>
        /// <param name="zipStream">讀取壓縮檔的Stream</param>
        /// <param name="zipToFolder">解壓縮資料夾位置</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        public static ZipExecuteResult UnZipToDirectory(Stream zipStream, string zipToFolder, string password = null, Encoding enc = null)
        {
            try
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

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }  
        }

        /// <summary>
        ///  解壓縮zip檔到目錄(byte[])
        /// </summary>
        /// <param name="zipFile">讀取壓縮檔的Byte[]</param>
        /// <param name="zipToFolder">解壓縮資料夾位置</param>
        /// <param name="password">解壓縮密碼</param>
        /// <param name="enc">編碼</param>
        /// <returns></returns>
        public static ZipExecuteResult UnZipToDirectory(byte[] zipFile, string zipToFolder, string password = null, Encoding enc = null)
        {
            try
            {
                if (!Directory.Exists(zipToFolder))
                    Directory.CreateDirectory(zipToFolder);

                Stream zipStream = new MemoryStream(zipFile);

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

                    return ZipExecuteResult.Ok();
                }
            }
            catch (Exception ex)
            {
                return ZipExecuteResult.Fail(ex.ToString());
            }
        }

        #endregion
    }
}
