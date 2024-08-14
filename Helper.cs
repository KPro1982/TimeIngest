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

        public static DateTime Convert2DateTime(string sentdate)
        {

            // "Thu, 25 Jul 2024 15:09:49 -0700"

            DateTime dt = DateTime.Parse(sentdate);
            
            return dt;
        }

        public static string Convert2String(DateTime dt)
        {
            string dtToString = dt.ToString("yyyyMMdd");
            return dtToString;
           
        }

    }
}