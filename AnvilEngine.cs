using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Python.Runtime;
using System.Text.Json;
using Org.BouncyCastle.Asn1.Cms;
using Newtonsoft.Json;
using RtfPipe.Tokens;
using Org.BouncyCastle.Crypto.Modes;
using System.Timers;
using Microsoft.VisualBasic;




// anvil.connect("server_7UDURRX3SBYJOCU3HKUJCCUV-CWQSJE54273JU7WL");
// var text = anvil.call("Process_Msg", file);

namespace TimeIngest
{
    public static class AnvilEngine
    {
        public static int entryIndex = 0;
        
        public static void Process()
        {
            var PythonPath =  Helper.GetPythonRoot() + ";" + Helper.GetModulePath();
            Runtime.PythonDLL = Helper.GetDLLPath();
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;


            var lastfilename = "";
            foreach (var file in 
            Directory.EnumerateFiles(Helper.GetExecutionPath(), "*.msg"))
            {


                
                dynamic extract_msg = Py.Import("extract_msg");
                dynamic sys = Py.Import("sys");
                sys.path.append(Path.Combine(@"G:\projects\TimeIngest\TimeIngest\"));
                var Generate = Py.Import("generate");
                dynamic os = Py.Import("os");


         
                var msg = extract_msg.openMsg(file);
                TimeEntry timeEntry = new TimeEntry();
                timeEntry.bcc = msg.bcc;
                timeEntry.cc = msg.cc;
                timeEntry.body = msg.body;
              
                timeEntry.filename = msg.filename;
                timeEntry.messageId = msg.messageId;
                timeEntry.sender    = msg.sender;
                timeEntry.subject = msg.subject;
                timeEntry.recipients = msg.to;

                

 
                string msgJson = msg.getJson();
                var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(msgJson);
                try
                {
                    timeEntry.sentdate = values["date"];
                    var subject = new PyString(timeEntry.subject);
                    var sender = new PyString(timeEntry.sender);
                    var recipient = new PyString(timeEntry.recipients);
                    var body = new PyString(timeEntry.body);

                   // var narrative = Generate.InvokeMethod("Narrative", new PyObject[] {recipient, sender, body, subject});
                    
                    var clientmatter = Generate.InvokeMethod("ClientMatter", new PyObject[] {subject});
                    
                    timeEntry.client = clientmatter.ToString();
                    // timeEntry.narrative = narrative.ToString();
                    // Console.WriteLine("Narrative:" + timeEntry.narrative);
                    
                }
                catch {}
  
                string msgdate = msg.date.ToString();
                timeEntry.date = ConvertDate(msgdate);

                
                AppendJson(timeEntry);
                MakeXL(timeEntry);
                
               

               
            }
                
               

            
          


        }

        public static string ConvertDate(string msgdate)
        {
                DateTime dt = DateTime.Parse(msgdate);
                string dtstr = dt.ToString("yyyyMMdd");
                

                return dtstr;
;
        }
        public static void AppendJson(TimeEntry entry)
        {
            Console.WriteLine(entry.subject);
            TimeEntry[] timeEntries = new TimeEntry[1000];
            string jsonFilename = Helper.GetJsonPath();
            try{
                
                 timeEntries = JsonConvert.DeserializeObject<TimeEntry[]>(File.ReadAllText(jsonFilename));
            }
            catch { }
            

                timeEntries[entryIndex++] = entry;
                
               

                
            
            

            
           
            // serialize JSON to a string and then write string to a file
            File.WriteAllText(jsonFilename, JsonConvert.SerializeObject(timeEntries, Formatting.Indented));
        }

        public static bool CheckJson(TimeEntry entry, TimeEntry[] timeEntries)  
        {
            foreach (var e in timeEntries)
            {
                if (e.messageId == entry.messageId)
                {
                    return false;
                }
            }

            return true;

        }

        public static void MakeXL(TimeEntry e)
        {

            var PythonPath =  Helper.GetPythonRoot() + ";" + Helper.GetModulePath();
            //Runtime.PythonDLL = Helper.GetDLLPath();
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;
            dynamic Pyxl = Py.Import("openpyxl");

            var msgWorkbook = Pyxl.load_workbook(@"G:\cravens timeentry.xlsx");

            var sheet = msgWorkbook["Sheet1"];
            int x = sheet.max_row + 1;
            


            sheet["a"+ x.ToString()].value = e.userId;
            sheet["b" + x.ToString()].value = e.sentdate; 
            sheet["c" + x.ToString()].value = e.timekeepr;
            sheet["d" + x.ToString()].value = e.client;
            sheet["e" + x.ToString()].value = e.matter;
            sheet["f" + x.ToString()].value = "";
            sheet["g" + x.ToString()].value = "";
            sheet["h" + x.ToString()].value = e.billable;
            sheet["i" + x.ToString()].value = e.hoursWorked;
            sheet["j" + x.ToString()].value = e.hoursWorked;  // always same as hoursworked
            sheet["k" + x.ToString()].value = "";
            sheet["l" + x.ToString()].value = "";
            sheet["m" + x.ToString()].value = "";
            sheet["n" + x.ToString()].value = "";
            sheet["o" + x.ToString()].value = "";
            sheet["p" + x.ToString()].value = "";
            sheet["q" + x.ToString()].value = "";
            sheet["r" + x.ToString()].value = "";
            sheet["s" + x.ToString()].value = e.narrative;
            sheet["t" + x.ToString()].value = e.body;
            msgWorkbook.save(@"G:\cravens timeentry.xlsx");

        }
        public static void Shutdown()
        {
            try {
                PythonEngine.Shutdown();
            }
            catch {}
        }
   

        
    }
}