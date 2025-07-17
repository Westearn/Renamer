using ClosedXML.Excel;
using System.Linq;
using System.Collections.Generic;
using System.IO;

namespace Renamer
{
    public class ExcelRename : IRename
    {
        public void DoRename(string directory, List<(string searchName, string replaceName)> renameList)
        {
            List<string> files = new List<string>();
            files.AddRange(Directory.GetFiles(directory, "*.xlsx").Where(f => !Path.GetFileName(f).StartsWith("~$")));
            var Check = MainWindow.CheckWildcards;
            Helpers helpers = new Helpers();

            foreach (var file in files)
            {
                using (var workbook = new XLWorkbook(file))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var cell in worksheet.CellsUsed())
                        {
                            string cellValue = cell.GetValue<string>();

                            foreach (var (searchName, replaceName) in renameList)
                            {
                                if (!string.IsNullOrEmpty(cellValue) && Check == false && helpers.Contains(cellValue, searchName))
                                {
                                    cellValue = helpers.Replace(cellValue, searchName, replaceName);
                                    cell.Value = cellValue;
                                }
                                else if (!string.IsNullOrEmpty(cellValue) && helpers.Comparator(cellValue, searchName))
                                {
                                    string newCellValue = helpers.Replacer(cellValue, searchName, replaceName);
                                    cell.Value = newCellValue;
                                }
                            }
                        }
                    }

                    workbook.Save();
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
