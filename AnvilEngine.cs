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
using System.Text.RegularExpressions;
using System.IO.Pipes;






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

            string clientstr = "";
            string matterstr = "";
            int successfulParse = 0;
            int unsuccessfulParse = 0;


            foreach (var file in 
            Directory.EnumerateFiles(GetEmailFolderPath(), "*.msg"))
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
                timeEntry.sender    = "SENDER OF EMAIL" + msg.sender;
                timeEntry.subject = "SUBJECT OF EMAIL" + msg.subject;
                timeEntry.recipients = "RECIPIENTS OF EMAIL" + msg.to;

                

 
                string msgJson = msg.getJson();
                var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(msgJson);
                var clientmatter = "";


                timeEntry.sentdate = values["date"];
                var api_key = new PyString(GetAPIKey());
                var subject = new PyString(timeEntry.subject);
                var sender = new PyString(timeEntry.sender);
                var recipient = new PyString(timeEntry.recipients);
                var body = new PyString(timeEntry.body);
                var aliasList = new PyString(GetAliasList());
                var narrativeexamples = new PyString(GetNarrativeExamples());

                var narrative = Generate.InvokeMethod("Narrative", new PyObject[] {api_key, recipient, sender, body, subject, narrativeexamples});
                
                var cmgenerated = Generate.InvokeMethod("ClientMatter", new PyObject[] {subject, api_key, aliasList});

                

                timeEntry.narrative = narrative.ToString();

                    


                if(GetCM(cmgenerated.ToString(), out clientstr, out matterstr))
                {
                    timeEntry.client = clientstr;
                    timeEntry.matter = matterstr;
                    successfulParse += 1;
                }
                else{
                    unsuccessfulParse += 1;

                }
                string msgdate = msg.date.ToString();
                timeEntry.date = ConvertDate(msgdate);
                
                float fSuccessPercent =  successfulParse / (successfulParse + unsuccessfulParse) * 100;
                Console.WriteLine("Success::Failure (" + successfulParse + "::" + unsuccessfulParse + ") [" + fSuccessPercent.ToString() + "]" );    
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
            // Console.WriteLine(entry.subject);
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
            sheet["u" + x.ToString()].value = e.subject;
            
            msgWorkbook.save(@"G:\cravens timeentry.xlsx");

        }

        public static Dictionary<string, string?>  GetMatterDictionary()
        {
            Dictionary<string, string?> dict = new Dictionary<string, string?>();
            dict.Add("None","0000-00000");
            foreach (string line in File.ReadLines(GetClientDataFileName()))
            {
                string removed = RemoveCommas(line);
                string[] fields  = ParseCM(removed);
                dict.Add(fields[0], fields[1]);               
                
                
                
                

            }
            return dict;
            
        }
        public static string[] ParseCM(string rawline)
        {
            
            string cleaned = CleanCM(rawline);
            string[] fields  = rawline.Split(",");
            try
            {
                fields[0] = fields[0].Replace("\"","");
                fields[1] = fields[1].Replace(".","-");
               
         
        
            }
            catch (Exception e) when (e is IndexOutOfRangeException) {};
            return fields;

        }

        public static string GetAliasList()
        {
            string aliaslist = "None;";
            foreach (string line in File.ReadLines(GetClientDataFileName()))
            {
                string removed = RemoveCommas(line); // remove internal commas
                string[] parsed = ParseCM(removed);
                aliaslist += parsed[0] + ";";
                                 
            }
            return aliaslist;
            
        }

         public static string GetNarrativeExamples()
        {
            string examples = "";
            foreach (string line in File.ReadLines(GetExampleFileName()))
            {
                examples += line;
                                 
            }
            return examples;
            
        }
        public static string RemoveCommas(string rawline)
        {
                string pattern = "(?<=\\\"[^\\\"]*),(?=[^\\\"]*\\\")";
                rawline = Regex.Replace(rawline, pattern, "");
                return rawline;
        }
        public static string CleanCM(string rawline)
        {


               // rawline = rawline.Replace(" ", "");
                string pattern = @"\?s";
                rawline = Regex.Replace(rawline, pattern, "s");
   //             pattern = "\"";
   //             rawline = Regex.Replace(rawline, pattern, ""); 
                rawline = rawline.TrimEnd(',');
                rawline = rawline.TrimEnd(' ');
             //   Console.WriteLine("Cleaned: " + rawline);

                return rawline;
               
        }
        public static string GetAPIKey()
        {
            var dict = File.ReadLines(@"G:\secret.txt").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line[1]);
            
            if(dict.TryGetValue("API_Key", out var apikey))
            {
                return apikey;
            }

            return "";
    

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
        
        public static bool GetCM(string cmstring, out string clientstr, out string matterstr)
        {
            RemoveCommas(cmstring);
            
           
            if(GetMatterDictionary().TryGetValue(cmstring, out string matter))
            {
                if(matter == "") matter = "0000-00001";
                string[] cmarr = matter.Split("-");
                clientstr = cmarr[0];
                matterstr = cmarr[1];
                return true;
            }
            else
            {
                Console.WriteLine("Processing: " + cmstring);
                foreach (var k in GetMatterDictionary())
                {
                    Console.WriteLine(k.Key + "::" + cmstring);
                    if(k.Key.Contains(cmstring))
                    {
                        try 
                        {
                            string[] cmarr = cmstring.Split("-");
                            clientstr = cmarr[0];
                            matterstr = cmarr[1];
                        }
                        catch 
                        {
                            clientstr = "";
                            matterstr = "";

                        }
                        
                        return true;
                    }
                }

                Console.WriteLine("MATTER CLIENT LOOKUP -- FAILURE!!!");
                clientstr = "0000";
                matterstr = "00000";
                return false;


            }
            
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