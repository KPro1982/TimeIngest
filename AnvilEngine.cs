using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Python.Runtime;


namespace TimeIngest
{
    public static class AnvilEngine
    {
        
        public static void Run(string scriptName)
        {
            
            var appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var PythonPath =  @"C:\Users\danie\AppData\Local\Programs\Python\Python312\;C:\Users\danie\AppData\Local\Programs\Python\Python312\Lib\site-packages\";
            var PythonDLL = Path.Join(Path.GetDirectoryName(appdata), @"\Local\Programs\Python\Python312\python312.dll");
            Runtime.PythonDLL = PythonDLL;
            
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;

            using(Py.GIL())
            {

                dynamic anvil = Py.Import("anvil.server");
                anvil.connect("server_7UDURRX3SBYJOCU3HKUJCCUV-CWQSJE54273JU7WL");
                var text = anvil.call("ChangeName", "clientdanny");
                Console.WriteLine(text);
                
            }


        }
    }
}