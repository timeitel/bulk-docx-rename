using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace bulk_docx_rename
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("reading directory");
            string dir = args[0];
            string[] files = Directory.GetFiles(dir);

            foreach (string file in files)
            {
                string currentFileName = file.Split("\\").Last();
                string fileExtension = currentFileName.Split(".").Last();
                string pattern = @"[A-Z]{2,3}[-][A-Z]{2,5}[-][A-Z]{2,4}[-]\d{3}";
                string title = "";
                string match;

                if (fileExtension == "docx")
                {
                    try
                    {
                        using (WordprocessingDocument doc = WordprocessingDocument.Open(file, true))
                        {
                            foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                            {
                                title = headerPart.RootElement.InnerText;

                                break;
                            }

                            Match m = Regex.Match(title, pattern);
                            if (m.Success)
                            {
                                Console.WriteLine("Found '{0}' at position {1}.", m.Value, m.Index);
                                match = m.Value;
                            }
                            else
                            {
                                match = "!EDIT";
                            }
                        }

                        File.Move(file, $"{dir}\\{match} {currentFileName}");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }

                }
                else
                {
                    File.Move(file, $"{dir}\\!EDIT {currentFileName}");
                }

            }

        }
    }
}
