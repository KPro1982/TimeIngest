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
            Runtime.PythonDLL = "C:\\Users\\danie\\AppData\\Local\\Programs\\Python\\Python312\\python312.dll";
            PythonEngine.Initialize();
            using(Py.GIL())
            {
                var pythonScript = Py.Import(scriptName);
                var result = pythonScript.InvokeMethod("say_hello");
                Console.WriteLine(result);

            }


        }
    }
}