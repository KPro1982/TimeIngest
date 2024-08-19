using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Python.Runtime;

namespace TimeIngest
{
    public class GetMessageData
    {

        public GetMessageData()
        {
            var PythonPath =  Helper.GetPythonRoot() + ";" + Helper.GetModulePath();
            Runtime.PythonDLL = Helper.GetDLLPath();
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;
        }

        public void Get(string filename, out string msgdatestr, out string messageidstr)
        {


		
            dynamic extract_msg = Py.Import("extract_msg");
            dynamic sys = Py.Import("sys");
            sys.path.append(Path.Combine(@"G:\projects\TimeIngest\TimeIngest\"));

            var msg = extract_msg.openMsg(filename);
            msgdatestr = msg.date.ToString();
            messageidstr = msg.messageId.ToString();
            msg.close();

          //  PythonEngine.Shutdown();

        }
    }
        

    
}