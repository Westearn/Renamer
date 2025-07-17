using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;


namespace Renamer
{
    public class WordRenameFields : IRename
    {
        public void DoRename(string directory, List<(string searchName, string replaceName)> renameList)
        {
            List<string> files = new List<string>();
            files.AddRange(Directory.GetFiles(directory, "*.docx").Where(f => !Path.GetFileName(f).StartsWith("~$")));
            files.AddRange(Directory.GetFiles(directory, "*.docm").Where(f => !Path.GetFileName(f).StartsWith("~$")));

            var Check = MainWindow.CheckWildcards;
            Helpers helpers = new Helpers();

            Application wordApp = new Application();
            Document doc = null;
            wordApp.Visible = true;

            try
            {
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
                    }
                    doc = wordApp.Documents.Open(file);
                    doc.Fields.Update();

                    foreach (Section section in doc.Sections)
                    {
                        foreach (HeaderFooter header in section.Headers)
                        {
                            header.Range.Fields.Update();
                        }
                        foreach (HeaderFooter footer in section.Footers)
                        {
                            footer.Range.Fields.Update();
                        }
                    }
                    doc.Save();
                }
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                }
                wordApp.Quit();
            }
            
            // Рекурсия для прохождения всех папок внутри выбранной директории
            foreach (var subDirectory in Directory.GetDirectories(directory))
            {
                DoRename(subDirectory, renameList);
            }
        }
    }
}
