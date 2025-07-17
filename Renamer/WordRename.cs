using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace Renamer
{
    public class WordRename : IRename
    {
        public void DoRename(string directory, List<(string searchName, string replaceName)> renameList)
        {
            List<string> files = new List<string>();
            files.AddRange(Directory.GetFiles(directory, "*.docx").Where(f => !Path.GetFileName(f).StartsWith("~$")));
            files.AddRange(Directory.GetFiles(directory, "*.docm").Where(f => !Path.GetFileName(f).StartsWith("~$")));

            var Check = MainWindow.CheckWildcards;
            Helpers helpers = new Helpers();

            foreach (var file in files)
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file, true))
                {
                    // Переименовывавние полей
                    var customPropsPart = wordDoc.CustomFilePropertiesPart;
                    
                    if (customPropsPart != null)
                    {
                        foreach (var customProp in customPropsPart.Properties.Elements<CustomDocumentProperty>())
                        {
                            if (customProp.VTLPWSTR != null)
                            {
                                foreach (var (searchName, replaceName) in renameList)
                                {
                                    if (Check == false && helpers.Contains(customProp.VTLPWSTR.Text, searchName))
                                    {
                                        customProp.VTLPWSTR.Text = helpers.Replace(customProp.VTLPWSTR.Text, searchName, replaceName);
                                    }
                                    else if (helpers.Comparator(customProp.VTLPWSTR.Text, searchName))
                                    {
                                        customProp.VTLPWSTR.Text = helpers.Replacer(customProp.VTLPWSTR.Text, searchName, replaceName);
                                    }
                                }
                            }
                        }
                        customPropsPart.Properties.Save();
                    }
                
                    // Переименовывавние основного текста и таблиц
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    /*foreach (var table in body.Descendants<Table>())
                    {
                        foreach (var row in table.Descendants<TableRow>())
                        {
                            foreach (var cell in row.Descendants<TableCell>())
                            {
                                foreach (var run in cell.Descendants<Run>())
                                {
                                    foreach (var text in run.Descendants<Text>())
                                    {
                                        foreach (var (searchName, replaceName) in renameList)
                                        {
                                            if (text.Text.Contains(searchName))
                                            {
                                                text.Text = text.Text.Replace(searchName, replaceName);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }*/

                    /*foreach (var table in body.Descendants<Table>())
                    {
                        foreach (var row in table.Descendants<TableRow>())
                        {
                            foreach (var cell in row.Descendants<TableCell>())
                            {
                                foreach (var paragraph in cell.Descendants<Paragraph>())
                                {
                                    string cellPText = paragraph.InnerText;
                                    Console.WriteLine(cellPText);

                                    foreach (var (searchName, replaceName) in renameList)
                                    {
                                        if (cellPText.Contains(searchName))
                                        {
                                            cellPText = cellPText.Replace(searchName, replaceName);
                                            
                                            var i = 0;
                                            foreach (var text in paragraph.Descendants<Text>())
                                            {
                                                if (i == 0)
                                                {
                                                    text.Text = cellPText;
                                                }
                                                else
                                                {
                                                    text.Remove();
                                                }
                                                i++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }*/

                    foreach (var table in body.Descendants<Table>())
                    {
                        foreach (var row in table.Descendants<TableRow>())
                        {
                            foreach (var cell in row.Descendants<TableCell>())
                            {
                                foreach (var paragraph in cell.Descendants<Paragraph>())
                                {
                                    StringBuilder stringBuilder = new StringBuilder();
                                    foreach (var run in paragraph.Descendants<Run>())
                                    {
                                        foreach (var text in run.Elements<Text>())
                                        {
                                            stringBuilder.Append(text.Text);
                                        }
                                    }

                                    string cellPText = stringBuilder.ToString();
                                    
                                    foreach (var (searchName, replaceName) in renameList)
                                    {
                                        if (Check == false && helpers.Contains(cellPText, searchName))
                                        {
                                            cellPText = helpers.Replace(cellPText, searchName, replaceName);

                                            var i = 0;
                                            foreach (var run in paragraph.Descendants<Run>())
                                            {
                                                foreach (var text in run.Elements<Text>())
                                                {
                                                    if (i == 0)
                                                    {
                                                        text.Text = cellPText;
                                                    }
                                                    else
                                                    {
                                                        text.Remove();
                                                    }
                                                    i++;
                                                }
                                            }
                                        }
                                        else if (helpers.Comparator(cellPText, searchName))
                                        {
                                            cellPText = helpers.Replacer(cellPText, searchName, replaceName);

                                            var i = 0;
                                            foreach (var run in paragraph.Descendants<Run>())
                                            {
                                                foreach (var text in run.Elements<Text>())
                                                {
                                                    if (i == 0)
                                                    {
                                                        text.Text = cellPText;
                                                    }
                                                    else
                                                    {
                                                        text.Remove();
                                                    }
                                                    i++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (var text in body.Descendants<Text>())
                    {
                        foreach (var (searchName, replaceName) in renameList)
                        {
                            if (Check == false && helpers.Contains(text.Text, searchName))
                            {
                                text.Text = helpers.Replace(text.Text, searchName, replaceName);
                            }
                            else if (helpers.Comparator(text.Text, searchName))
                            {
                                text.Text = helpers.Replacer(text.Text, searchName, replaceName);
                            }
                        }
                    }

                    // Переименовывавние верхнего колонитутла
                    foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                    {
                        var par = headerPart.RootElement.Descendants<Text>();
                        foreach (var text in par)
                        {
                            foreach (var (searchName, replaceName) in renameList)
                            {
                                if (Check == false && helpers.Contains(text.Text, searchName))
                                {
                                    text.Text = helpers.Replace(text.Text, searchName, replaceName);
                                }
                                else if (helpers.Comparator(text.Text, searchName))
                                {
                                    text.Text = helpers.Replacer(text.Text, searchName, replaceName);
                                }
                            }
                        }
                        headerPart.RootElement.Save();
                    }

                    // Переименовывавние нижнего колонитутла
                    foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                    {
                        var par = footerPart.RootElement.Descendants<Text>();
                        foreach (var text in par)
                        {
                            foreach (var (searchName, replaceName) in renameList)
                            {
                                if (Check == false && helpers.Contains(text.Text, searchName))
                                {
                                    text.Text = helpers.Replace(text.Text, searchName, replaceName);
                                }
                                else if (helpers.Comparator(text.Text, searchName))
                                {
                                    text.Text = helpers.Replacer(text.Text, searchName, replaceName);
                                }
                            }
                        }
                        footerPart.RootElement.Save();
                    }

                    wordDoc.MainDocumentPart.Document.Save();
                }
            }

            // Рекурсия для прохождения всех папок внутри выбранной директории
            foreach (var subDirectory in Directory.GetDirectories(directory))
            {
                DoRename(subDirectory, renameList);
            }
        }
    }
}
