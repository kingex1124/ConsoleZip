using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleZip
{
    class Program
    {
        static void Main(string[] args)
        {
            //DotNetZipHelper.ZipSingleFile(@"D:\hana\dpagent_windows.zip", @"D:\123.zip");

            var result = DotNetZipHelper.ZipSingleFileStream(@"D:\hana\dpagent_windows.zip");

        }

    }
}
