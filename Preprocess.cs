using System.Linq;
using System.IO;
using System.Diagnostics;
using TimeIngest;
using Python.Runtime;

public static class PreProcess 
{
	
	public static void RenameEmail()
	{

		GetMessageData Msg_Extract = new GetMessageData();



		foreach (var old_file in 
            Directory.EnumerateFiles(Helper.GetEmailFolderPath(), "*.test"))
            {


				if(!old_file.StartsWith("_-_"))
				{
					string msgidstr = "";
					string msgdatestr = "";
					
					Msg_Extract.Get(old_file, out msgdatestr, out msgidstr);
					msgidstr = msgidstr.TrimEnd('>').Replace("<","").Replace("@","-at-").Replace(".","-dot-");
					string msgid = msgidstr;
					string msgtime = Helper.ConvertTime(msgdatestr);
					msgtime.Replace(":","-");
                	string datestr = Helper.ConvertDate(msgdatestr);
					FileInfo fileInfo = new FileInfo(old_file);
					string pathstr = fileInfo.Directory.FullName;
					string namestr = fileInfo.Name.ToString();
					string new_name = "_-_" + datestr + "_" + msgtime + "_" + msgid;
					string new_file = pathstr + "\\" + new_name;
	
					
					try {
						if (File.Exists(old_file) && !File.Exists(new_file)) {
							File.Move(old_file, new_file);
						} else {
							Console.WriteLine("Rename error: " + old_file + " exists!");
						}
					} catch(Exception ex) {
						Console.WriteLine("Rename error: Cannot rename " + old_file + " to " + new_file);
					}
				}



			}
		
	}

}








