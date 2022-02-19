using System;
using System.Xml.Linq;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using ClosedXML.Excel;

namespace PDFJuice
{
    internal class Program
    {
        private const string KAP1 = "Kap 00";
        private const string KAP2 = "Kap 000";
        private const string KAP3 = "KAP 000.0";
        private const string TEXT1 = "Text 00";
        private const string TEXT2 = "Text 000";
        private const string TEXT3 = "Text 000.0";
        private const string TEXT = "Text";
        private const string FONT = "Font";
        private const string UNDERLINED = "Underlined";
        private const string MENGE = "Menge";
        private const string ME = "ME";
        public static readonly string[] meTypes = { "Stk.", "m", "kpl.", "Stk" };

        static void Main(string[] args)
        {
            List<TextModel> textModelList = new List<TextModel>();
            TextModel textModelPrev = null;
            string workingDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string pdfFilePath = workingDir + "\\Beispiel_1_Devi.pdf";
            
            if (args.Length == 0)
            {
                ShowHelp();
            }
            else if (args.Length == 1)
            {
                ShowHelp();
            }
            else if (args.Length == 2)
            {
                if (args[0].ToLower() == "extract")
                {
                    if (File.Exists(args[1]))
                    {
                        pdfFilePath = args[1];
                        GenerateSVGFiles(pdfFilePath);
                        // Need to loop to all converted svg files
                        DirectoryInfo dirInfo = new DirectoryInfo(workingDir);
                        FileSystemInfo[] svgFiles = dirInfo.GetFileSystemInfos("*.svg");
                        var orderSvgFiles = svgFiles.OrderBy(f => f.CreationTime);

                        int count = 0;
                        foreach (var svgFile in orderSvgFiles)
                        {
                            count++;
                            Console.WriteLine("Parsing Page : {0}", count);
                            Parse(textModelList, textModelPrev, dirInfo.FullName + "\\" + svgFile.Name);
                            PostProcessing(textModelList);
                            textModelPrev = new TextModel(textModelList.LastOrDefault());
                            textModelPrev.Text = "";
                            textModelPrev.Font = "";
                            textModelPrev.Underlined = "";
                            textModelPrev.Me = "";
                            textModelPrev.Menge = "";
                        }

                        Console.WriteLine("");
                        Display(textModelList);
                        DeleteSVGFiles(workingDir);

                        PopulateExcelDoc(textModelList);
                    }
                    else 
                    {
                        Console.WriteLine("No pdf file found");
                    }
                }
                else 
                {
                    ShowHelp();
                }
            }
            
        }

        static void Parse(List<TextModel> textModelList, TextModel textModelPrev, string filePath)
        {
            XDocument doc = XDocument.Load(filePath);
            XElement rootElements = doc.Root;

            IEnumerable<XElement> nodes = from element1 in rootElements.Elements("{http://www.w3.org/2000/svg}g") select element1;

            foreach (var node in nodes)
            {
                //IEnumerable<XElement> childNodes = from element2 in node.Elements("{http://www.w3.org/2000/svg}text") select element2;
                IEnumerable<XElement> childNodes = node.Descendants().Where(x => x.Name.LocalName == "text"
                                                                            || x.Name.LocalName == "path").Skip(1);

                foreach (var childNod in childNodes)
                {
                    //Get child of <g>, ract tag
                    string txtLocalName = childNod.Name.LocalName;

                    if (txtLocalName.Equals("text"))
                    {
                        if (childNod.Value.Trim().Equals("TextMengeMEPreisBetrag"))
                        {
                            //skip this particular element
                            continue;
                        }

                        //Get Attribute values like "style", "width", "height", etc..
                        string fontSize = childNod.Attributes().Where(e => e.Name.LocalName.Equals("font-size")).FirstOrDefault()?.Value;
                        string fontFamily = childNod.Attributes().Where(e => e.Name.LocalName.Equals("font-family")).FirstOrDefault()?.Value;
                        string fontWeight = childNod.Attributes().Where(e => e.Name.LocalName.Equals("font-weight")).FirstOrDefault()?.Value;
                        bool bIsBold = fontWeight != null && fontWeight.Equals("bold");

                        

                        if (childNod.HasElements)
                        {
                            IEnumerable<XElement> spans = childNod.Elements();

                            foreach (var span in spans)
                            {
                                TextModel textModel = null;
                                bool isMeTypes = false;
                                string input = span.Value.Trim();

                                if (!String.IsNullOrEmpty(input))
                                {
                                    //Console.WriteLine("Text: " + input);

                                    if (bIsBold)
                                    {
                                        string output = new string(input.TakeWhile(char.IsDigit).ToArray());

                                        if (!String.IsNullOrEmpty(output))
                                        {
                                            int dotIndex = input.IndexOf('.');
                                            if (dotIndex > -1 && dotIndex < input.Length)
                                            {
                                                string[] decimalPart = input.Split('.');
                                                if (decimalPart.Length == 2)
                                                {
                                                    int decimalPlace = decimalPart[1].TakeWhile(char.IsDigit).Count();
                                                    output = output + input.Substring(dotIndex, decimalPlace + 1);
                                                }
                                            }

                                            if (textModelList.Count == 0)
                                            {
                                                // L1
                                                textModel = new TextModel();
                                                textModel.Kap1 = output;
                                                textModel.Text1 = input.Substring(output.Length);
                                                //l1Kp1 = output;
                                                //l1Text1 = input;
                                            }

                                            else if (output.Length == 3)
                                            {
                                                // L2
                                                textModel = new TextModel(textModelPrev);
                                                textModel.Kap2 = output;
                                                textModel.Text2 = input.Substring(output.Length);
                                            }
                                            else if (output.Length == 5)
                                            {
                                                // L3

                                                textModel = new TextModel(textModelPrev);
                                                textModel.Kap3 = output;
                                                textModel.Text3 = input.Substring(output.Length);
                                                //l1Kp3 = output;
                                                //l1Text3 = input;
                                            }
                                        }
                                        else
                                        {
                                            // Text with bold
                                            textModel = new TextModel(textModelPrev);
                                            textModel.Text = input;
                                        }

                                        textModel.Font = fontFamily + " " + fontSize;
                                    }
                                    else
                                    {
                                        
                                        textModel = new TextModel(textModelPrev);

                                        textModel.Font = fontFamily + " " + fontSize;

                                        if (input.Contains("..................................................."))
                                        {
                                            input = input.Replace("...................................................", "");
                                        }

                                        string num = new string(input.TakeWhile(char.IsDigit).ToArray());
                                        Decimal dummy;
                                        bool isAllNumber = Decimal.TryParse(input, out dummy);

                                        if (!num.Equals(string.Empty) && !isAllNumber)
                                        {

                                            if (meTypes.Any(input.Contains) && !(input.Contains("/") || input.Contains('"')))
                                            {
                                                
                                                var lastModel = textModelList.LastOrDefault();
                                                if (lastModel != null)
                                                {
                                                    string tempMe = input.Substring(num.Length);
                                                    if (tempMe.Length > 5)
                                                    {
                                                        // This is not for Me. This is for Text.
                                                        textModel.Text = input;
                                                    }
                                                    else 
                                                    {
                                                        isMeTypes = true;
                                                        textModelList.LastOrDefault().Menge = num;
                                                        textModelList.LastOrDefault().Me = input.Substring(num.Length);
                                                    }
                                                }

                                            }
                                            else 
                                            {
                                                string tempInput = "";
                                                foreach (var me in meTypes) 
                                                {
                                                    if (input.Contains(me))
                                                    {
                                                        tempInput = input.Replace(me, "");
                                                        tempInput = tempInput.Substring(num.Length).Trim();
                                                        if (tempInput.Length > 5)
                                                        {
                                                            textModel.Text = input;
                                                        }
                                                        else 
                                                        {
                                                            textModel.Me = me;
                                                            textModel.Menge = num;
                                                            textModel.Text = tempInput;
                                                        }
                                                        break;
                                                    }                                                    
                                                }

                                                if (string.IsNullOrEmpty(tempInput)) 
                                                {
                                                    // It is a text only.
                                                    textModel.Text = input;
                                                }

                                            }
                                        }
                                        else
                                        {
                                            textModel.Text = input;
                                        }
                                    }
                                }

                                if (textModel != null && !isMeTypes)
                                {
                                    textModel.Underlined = ""; // reset underlined property

                                    if (spans.Count() == 1 && bIsBold)
                                    {
                                        if (string.IsNullOrEmpty(textModel.Text))
                                        {

                                            textModelPrev = new TextModel(textModel);
                                        }
                                        else
                                        {
                                            textModelList.Add(textModel);
                                        }
                                    }
                                    else
                                    {
                                        textModelList.Add(textModel);
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        bool isTransformed = childNod.Attributes().Where(e => e.Name.LocalName.Equals("transform")).Any();
                        if (isTransformed)
                        {
                            int attributeCounts = childNod.Attributes().Count();
                            if (attributeCounts > 2)
                            {
                                // This is the lines comprises the table.
                            }
                            else
                            {
                                // Now we are we are sure that the previous text/heading is underlined.
                                if (textModelPrev != null)
                                {
                                    textModelPrev.Underlined = "Underlined";
                                    textModelList.Add(textModelPrev);
                                }
                            }
                        }
                        else
                        {
                            // Just an empty text let's ignore this.
                        }
                    }

                }
            }

            Console.WriteLine("Parsing Done");
        }

        static void PostProcessing(List<TextModel> textModelList)
        {
            if (textModelList.Count > 0)
            {
                textModelList.RemoveAll(e => e.Text == "............................");
                textModelList.RemoveAll(e => e.Text != null && e.Text.Contains("............................"));
                textModelList.RemoveAll(e => string.IsNullOrEmpty(e.Font));
                textModelList.RemoveAll(e => e.Text == "TextMengeMEPreisBetrag");

                int totalIndex = textModelList.FindIndex(e => e.Text == "Total");
                int indexToCombine = textModelList.Count() - totalIndex;
                if (totalIndex > 0)
                {
                    // This will be the list to combined and later be removed.
                    var newList = textModelList.Skip(Math.Max(0, textModelList.Count() - indexToCombine));
                    TextModel combinedModel = null;

                    if (newList != null)
                    {
                        combinedModel = newList.FirstOrDefault();
                        string text = "";
                        bool isUnderlined = false;

                        foreach (var item in newList)
                        {
                            TextModel textModel = item as TextModel;
                            if (item.Text != null)
                            {
                                text += item.Text + " ";
                                if (!isUnderlined) 
                                {
                                    isUnderlined = !string.IsNullOrEmpty(item.Underlined) || item.Text.Equals("Total");
                                }
                            }
                        }
                        combinedModel.Text = text.Trim();
                        combinedModel.Underlined = "Underlined";
                    }

                    textModelList.RemoveRange(totalIndex, indexToCombine);
                    textModelList.Add(combinedModel);
                }
                
            }
        }

        static void Display(List<TextModel> textModelList) 
        {
            Console.WriteLine("Display Table...");
            foreach (var textModel in textModelList)
            {
                //string underLined = textModel.IsUnderlined ? "Underlined" : "";

                Console.WriteLine(textModel.Kap1 + " - " + textModel.Text1 + " - " +
                                  textModel.Kap2 + " - " + textModel.Text2 + " - " +
                                  textModel.Kap3 + " - " + textModel.Text3 + " - " +
                                  textModel.Text + " - " + textModel.Font + " - " +
                                  textModel.Text + " - " + textModel.Font + " - " +
                                  textModel.Underlined + " - " + textModel.Menge + " - " +
                                  textModel.Me + " - ");
            }
        }

        static void GenerateSVGFiles(string pdfFilePath) 
        {
            Process process = new Process();
            // Configure the process using the StartInfo properties.
            process.StartInfo.FileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\thirdparty\\mutool.exe";
            string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string args = string.Format("convert -F svg -O text=text -o {0}\\your_pdf_pg.svg {1}", executableLocation,pdfFilePath);
            process.StartInfo.Arguments = args;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.Start();
            process.WaitForExit();// Waits here for the process to exit.
        }

        static void DeleteSVGFiles(string workingDir) 
        {
            Directory.GetFiles(workingDir,"*.svg").ToList().ForEach(File.Delete);
        }

        static void ShowHelp() 
        {
            Console.WriteLine(
                "Parse PDF text and styles/formatting." + "\n" +
                     "Extract:" + "\n" +
                     "\t" + "extract pdf_path" + "\n" +
                     "\t" + "example : extract C:\\temp\\mypdf.pdf" + "\n" +
                     "Help:" + "\t" + "h or help"
                );
        }

        static void PopulateExcelDoc(List<TextModel> textModel) 
        {
            var wb = new XLWorkbook(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Devi_XLS_Template.xlsx");
            var worksheet1 = wb.Worksheet("Tabelle1");
            var textModelEnum = textModel.AsEnumerable();

            var rangeWithData = worksheet1.Cell(2, 2).InsertData(textModelEnum);
            worksheet1.Columns().AdjustToContents();
            wb.SaveAs("InsertingData.xlsx");
        }
    }
}
