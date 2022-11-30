using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Function
{
    public class Function_Zip
    {
        /// <summary>
        /// ZIP解壓縮
        /// <param name="zipFile"></param>
        /// <param name="outDir"></param>
        /// <param name="overwrite"></param>
        /// 若已存在的檔案不想被覆蓋，overwrite 可設為 false
        /// </summary>
        public static void UnzipToDirectory(string zipFile, string outDir, bool overwrite = true)
        {

            ZipArchive archive = ZipFile.OpenRead(zipFile);

            foreach (ZipArchiveEntry file in archive.Entries)
            {
                // file.Name == "" 表示 file 為目錄

                if (file.Name == "")
                {
                    string desPath = Path.Combine(outDir, file.FullName);
                    // 目錄不存在就要建立
                    if (!Directory.Exists(desPath))
                    {
                        Directory.CreateDirectory(desPath);
                    }
                }
                else
                {
                    // 目錄不存在就要建立
                    if (!Directory.Exists(outDir))
                    {
                        Directory.CreateDirectory(outDir);
                    }
                    // file 為檔案
                    string desFileName = Path.Combine(outDir, file.FullName);
                    // 可覆寫就直接解壓縮
                    if (overwrite)
                    {
                        file.ExtractToFile(desFileName, true);
                    }
                    else
                    {
                        // 不可覆寫就要先判斷檔案是否存在，不存在才解壓縮
                        if (!File.Exists(desFileName))
                        {
                            file.ExtractToFile(desFileName);
                        }
                    }
                }
            }
        }
    }
}
