using System;
using System.IO;


namespace Renamer
{
    public class Loader
    {
        public string DirectoryLoader()
        {
            string selectedPath = null;

            using (var folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Выберите папку";

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    selectedPath = folderBrowserDialog.SelectedPath;
                    if (selectedPath.Contains(@"tnn\pir") || selectedPath.Contains("projects") || (Path.GetPathRoot(selectedPath) ?? "").StartsWith("\\\\"))
                    {
                        return null;
                    }
                    return selectedPath;
                }
                else
                {
                    throw new Exception("Выбор папки отменен");
                }
            }
        }
    }
}
