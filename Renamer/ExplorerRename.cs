using System.Collections.Generic;
using System.IO;
using System;

namespace Renamer
{
    public class ExplorerRename : IRename
    {
        public void DoRename(string directory, List<(string searchName, string replaceName)> renameList)
        {
            var Check = MainWindow.CheckWildcards;
            Helpers helpers = new Helpers();

            foreach (var file in Directory.GetFiles(directory))
            {
                var fileName = Path.GetFileName(file);

                foreach (var (searchName, replaceName) in renameList)
                {
                    if (Check == false && helpers.Contains(fileName, searchName))
                    {
                        fileName = helpers.Replace(fileName, searchName, replaceName);
                    }
                    else if (helpers.Comparator(fileName, searchName))
                    {
                        fileName = helpers.Replacer(fileName, searchName, replaceName);
                    }
                }

                var newFilePath = Path.Combine(directory, fileName);
                if (file != newFilePath)
                {
                    File.Move(file, newFilePath);
                }
            }

            foreach (var subDirectory in Directory.GetDirectories(directory))
            {
                var folderName = Path.GetFileName(subDirectory);

                foreach (var (searchName, replaceName) in renameList)
                {
                    
                    if (Check == false && helpers.Contains(folderName, searchName))
                    {
                        folderName = helpers.Replace(folderName, searchName, replaceName);
                    }
                    else if (helpers.Comparator(folderName, searchName))
                    {
                        Console.WriteLine(folderName);
                        Console.WriteLine(helpers.Replacer(folderName, searchName, replaceName));
                        folderName = helpers.Replacer(folderName, searchName, replaceName);
                        Console.WriteLine(folderName);
                    }
                }

                var newFolderPath = Path.Combine(Path.GetDirectoryName(subDirectory), folderName);
                if (subDirectory != newFolderPath)
                {
                    Directory.Move(subDirectory, newFolderPath);
                }

                DoRename(newFolderPath, renameList);
            }
        }
    }
}
