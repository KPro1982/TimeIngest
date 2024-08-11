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
        
        public static void Process()
        {
            var PythonPath =  @"C:\Users\danie\AppData\Local\Programs\Python\Python312\;C:\Users\danie\AppData\Local\Programs\Python\Python312\Lib\site-packages\";
            var PythonDLL = @"C:\Users\danie\AppData\Local\Programs\Python\Python312\python312.dll";
            Runtime.PythonDLL = PythonDLL;
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;

            foreach (var file in 
            Directory.EnumerateFiles(@"G:\projects\TimeIngest\TimeIngest", "*.msg"))
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
                    timeEntry.date = values["date"];
                }
                catch {}
                
                AppendJson(timeEntry);
                
               
            }
                
               
                
            
         


        }
        public static void AppendJson(TimeEntry entry)
        {
            Console.WriteLine(entry.subject);
            TimeEntries timeEntries = new TimeEntries();
            try{
                 timeEntries = JsonConvert.DeserializeObject<TimeEntries>(File.ReadAllText(@"G:\timeentries.json"));
            }
            catch { }
            
            if(CheckJson(entry, timeEntries))
            {
                timeEntries.entries.Add(entry);
            }
                
            
            

            
           
            // serialize JSON to a string and then write string to a file
            File.WriteAllText(@"G:\timeentries.json", JsonConvert.SerializeObject(timeEntries, Formatting.Indented));
        }

        public static bool CheckJson(TimeEntry entry, TimeEntries timeEntries)  
        {
            foreach (var e in timeEntries.entries)
            {
                if (e.messageId == entry.messageId)
                {
                    return false;
                }
            }

            return true;

        }
   

        
    }
}