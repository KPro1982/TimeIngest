using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Python.Runtime;
using System.Text.Json;
using Org.BouncyCastle.Asn1.Cms;
using Newtonsoft.Json;




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

            foreach (var file in 
            Directory.EnumerateFiles(Helper.GetExecutionPath(), "*.msg"))
            {
                dynamic extract_msg = Py.Import("extract_msg");

                
         
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
                    //string default_format = @"%a %b %d %Y %H:%M:%S ";
                    // Fri, 09 Aug 2024 10:30:47 -0700

                   // DateTime dt = DateTime.strptime(timeEntry.sentdate, default_format);
                   // Console.WriteLine("Converted Date: " + dt.ToString(""))
                }
                catch {}
                
                AppendJson(timeEntry);
                
               
            }
                
               
                
            
         


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
            
   //         if(CheckJson(entry, timeEntries))
   //         {
                timeEntries[entryIndex++] = entry;
                
               
   //         }
                
            
            

            
           
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

        public static void MakeXL()
        {

            var PythonPath =  Helper.GetPythonRoot() + ";" + Helper.GetModulePath();
            Runtime.PythonDLL = Helper.GetDLLPath();
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;
            dynamic extract_msg = Py.Import("extract_msg");

                  


        }
   

        
    }
}