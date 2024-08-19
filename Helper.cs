using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Org.BouncyCastle.Asn1;


namespace TimeIngest
{
    public static class Helper
    {
        
        public static string GetDLLPath()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312\python312.dll");

        }
        public static string GetPythonRoot()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312");
        }

        public static string GetModulePath()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312\Lib\site-packages");
        }

        public static string GetExecutionPath()
        {
            string? currentPath = Path.GetDirectoryName(System.AppContext.BaseDirectory);
            if (currentPath != null) {
                return  Path.GetFullPath(Path.Combine(currentPath, "."));
            }
            else{
                return "";
            }
            

        }
        public static string GetJsonPath()
        {
            return Helper.GetExecutionPath() + @"\timeentries.json";
        }

          public static string GetClientDataFileName()
        {
            return @"G:\clientdata.csv";
        }
        public static string GetExampleFileName()
        {
            return @"G:\timeexamples.csv";
        }

        
        
        public static string GetEmailFolderPath()
        {
            return @"G:\Data\Email";
        }

        public static string ConvertDate(string msgdate)
        {
            DateTime dt = DateTime.Parse(msgdate);
            string dtstr = dt.ToString("yyyyMMdd");
            

            return dtstr;
;
        }
        public static string ConvertTime(string msgtime)
        {
            DateTime dt = DateTime.Parse(msgtime);
            string dtstr = dt.ToString("H:mm:ss");
            

            return dtstr;
;
        }

    }
}