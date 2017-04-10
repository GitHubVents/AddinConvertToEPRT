using System.IO;

namespace AddinConvertTo.Classes
{
    static class Logger
    {        
        static string myDir = "C:\\LogPDM";
        static string path = myDir + "\\LogEPRT.txt";   //static string path = @"\\" + "192.168.14.11" + @"\SolidWorks Admin\Bat\Log.txt";               

        public static void Add(string message)
        {
            if (!Directory.Exists(myDir))
            {
                Directory.CreateDirectory(myDir);
            }
            else
            {                
                using (StreamWriter writetext = File.AppendText(path))
                {
                    writetext.WriteLine(message);
                }
            }
        }        
    }
}