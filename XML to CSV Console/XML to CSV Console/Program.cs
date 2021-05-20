using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace XML_to_CSV_Console
{
    class Program
    {
        public static List<string> Headers = new List<string>();

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Enter data folder directory or directory to single xml file");
            string filepath = Console.ReadLine();
            string savedir = SetupSaveDirectory();

            if(savedir != null)
            {
                if(File.Exists(filepath))
                {
                    XmlReader reader = CreateReader(filepath);
                    Headers = GetHeaders(reader);
                    List<string> csv = GetValueByHeaders(reader);
                    SaveCSV(savedir, filepath, csv);
                }
                else
                {
                    foreach (string dir in Directory.GetFiles(filepath, "*.xml", SearchOption.AllDirectories))
                    {
                        XmlReader reader = CreateReader(dir);
                        Headers = GetHeaders(reader);
                        List<string> csv = GetValueByHeaders(reader);
                        SaveCSV(savedir, dir, csv);
                    }
                }
            }
            else
            {
                Console.WriteLine("Either something went wrong or you did not pick a save directory");
                Console.WriteLine("Press any key to exit");
            }
        }

        private static XmlReader CreateReader(string filepath)
        {
            try
            {
                return XmlReader.Create(filepath);
            }
            catch (FileNotFoundException exc)
            {
                Console.WriteLine(exc);
            }
            catch (PathTooLongException exc)
            {
                Console.WriteLine(exc);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc);
            }

            return null;
        }

        private static List<string> GetHeaders(XmlReader reader)
        {
            Console.WriteLine("Getting headers...");
            List<string> headers = new List<string>();
            reader.MoveToContent();

            string firstheader = "";

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if(firstheader.Length < 1)
                    {
                        firstheader = reader.LocalName;
                    }
                    if (reader.LocalName == firstheader && headers.Count > 0) break;
                    if (!headers.Contains(reader.LocalName)) headers.Add(reader.LocalName);
                }
            }

            return headers;
        }

        private static List<string> GetValueByHeaders(XmlReader reader)
        {
            Console.WriteLine("Getting values...");
            List<string> csv = new List<string>();
            string[] cache = new string[Headers.Count];

            string lastHeader = "";

            reader.MoveToContent();
            using (reader)
            {
                string[] line = new string[Headers.Count];

                while (reader.Read())                                                   //While there are values to read...
                {
                    switch (reader.NodeType)                                            //Create switch for type check
                    {
                        case XmlNodeType.Element:                                       //If case is a headeer
                            lastHeader = reader.LocalName;
                            break;
                        case XmlNodeType.Text:                                          //If case is a value
                            if (lastHeader != "" && Headers.Contains(lastHeader))      //If the header shows up in checked headers
                            {
                                line[Headers.IndexOf(lastHeader)] = reader.Value;     //Set new value to index of header column
                            }
                            else if(lastHeader != "")
                            {
                                Headers.Add(lastHeader);
                                string[] cacheline = new string[line.Length + 1];
                                string[] newvalue = new string[] { reader.Value };
                                line.CopyTo(cacheline, 0);
                                newvalue.CopyTo(cacheline, line.Length);
                                line = cacheline;
                                line[Headers.IndexOf(lastHeader)] = reader.Value;     //Set new value to index of new header column
                            }
                            break;
                        case XmlNodeType.Whitespace:
                            string row = "";
                            bool start = true;
                            foreach (string s in line)                               //for each value in the line string...
                            {
                                if (start)
                                {
                                    row += s;
                                    start = false;
                                }
                                else row += "," + s; ;
                            }
                            csv.Add(row);                                         //Add line to csv value
                            line = new string[Headers.Count];                     //Clear cache line value
                            break;
                    }
                }
            }

            return csv;
        }

        private static void SaveCSV(string savepath, string filepath, List<string> csv)
        {
            try
            {
                Console.WriteLine("Saving csv...");
                string csvFilePath = savepath + "/" + Path.GetFileNameWithoutExtension(filepath) + ".csv"; //Join selected path with opened xml file name to create save directory

                if(File.Exists(csvFilePath))
                {
                    File.Delete(csvFilePath);
                }

                using (StreamWriter file = new StreamWriter(csvFilePath, true))
                {
                    string csvheaders = "";
                    foreach (string s in Headers)
                    {
                        if (csvheaders.Length > 0) csvheaders += "," + s;                 //If the cache header string its length is bigger then 0 add a comma in front of the string      
                        else csvheaders = s;                                           //Else set the cache header to the new found header
                    }
                    file.WriteLine(csvheaders);

                    foreach (string s in csv)                                     //For each string in lines
                    {
                        file.WriteLine(s);                                          //Write CSV file for each CSV data in parameter
                    }
                }

                Console.WriteLine("Done!");
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Another process is probably using the document provided");
            }
        }

        private static string SetupSaveDirectory()
        {
            using(FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                Console.WriteLine("Pick a directory to save the CSV file in");

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.SelectedPath;
                }
            }

            return null;
        }
    }
}
