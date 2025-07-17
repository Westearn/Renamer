using System;
using System.IO;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;

using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;

namespace Renamer
{
    public class LoaderDLL : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                if (new Uri(assemblyPath).IsUnc || (Path.GetPathRoot(assemblyPath) ?? "").StartsWith("\\\\"))
                {
                    throw new InvalidOperationException("Загрузка DLL с сетевого диска запрещена");
                }

                Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
                editor.WriteMessage("Плагин Renamer успешно загружен");
            }
            catch (System.Exception ex)
            {
                Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
                editor.WriteMessage($"{ex.Message}");
                Application.ShowAlertDialog($"{ex.Message}");
                throw;
            }
        }
        public void Terminate()
        {
        }
    }
    public class CommandManager
    {
        /// <summary>
        /// Создание команды "RENAMER"
        /// </summary>
        [CommandMethod("RENAMER")]
        public void Renamer()
        {
            try
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.ShowDialog();
                
                string selectedPath = mainWindow.selectedPath;

                List<(string searchName, string replaceName)> renameList = mainWindow.renameList;

                if (mainWindow.DialogResult == true)
                {
                    switch (mainWindow.SelectedAction)
                    {
                        case "Action1":
                            RenameFields(selectedPath, renameList);
                            break;
                        case "Action2":
                            RenameAll(selectedPath, renameList);
                            break;
                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
                editor.WriteMessage($"{ex.Message}");
                Application.ShowAlertDialog($"{ex.Message}");
            }
            
        }

        public void RenameFields(string selectedPath, List<(string searchName, string replaceName)> renameList)
        {
            ExplorerRename explorerRename = new ExplorerRename();
            explorerRename.DoRename(selectedPath, renameList);

            WordRenameFields wordRenameFields = new WordRenameFields();
            wordRenameFields.DoRename(selectedPath, renameList);

            ExcelRename excelRename = new ExcelRename();
            excelRename.DoRename(selectedPath, renameList);

            try
            {
                CADRename cADRename = new CADRename();
                cADRename.DoRename(selectedPath, renameList);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"{ex.Message}");
            }

            MessageBox.Show("Файлы и поля переименованы");
        }

        public void RenameAll(string selectedPath, List<(string searchName, string replaceName)> renameList)
        {
            ExplorerRename explorerRename = new ExplorerRename();
            explorerRename.DoRename(selectedPath, renameList);

            WordRename wordRename = new WordRename();
            wordRename.DoRename(selectedPath, renameList);

            ExcelRename excelRename = new ExcelRename();
            excelRename.DoRename(selectedPath, renameList);

            try
            {
                CADRename cADRename = new CADRename();
                cADRename.DoRename(selectedPath, renameList);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"{ex.Message}");
            }

            MessageBox.Show("Файлы, поля и текст заменены");
        }
    }
}
