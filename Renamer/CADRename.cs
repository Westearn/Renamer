using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;

using Autodesk.AutoCAD.DatabaseServices;


namespace Renamer
{
    public class CADRename :IRename
    {
        public void DoRename(string directory, List<(string searchName, string replaceName)> renameList)
        {
            List<string> files = new List<string>();
            files.AddRange(Directory.GetFiles(directory, "*.dwg").Where(f => !Path.GetFileName(f).StartsWith("~$")));
            var Check = MainWindow.CheckWildcards;
            Helpers helpers = new Helpers();

            foreach (var file in files)
            {
                Database db = new Database(false, true);
                try
                {
                    db.ReadDwgFile(file, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                    db.CloseInput(true);

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                        ProcessBlockTableRecord(bt[BlockTableRecord.ModelSpace], tr, renameList);

                        DBDictionary layouts = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        foreach (DBDictionaryEntry layoutEntry in layouts)
                        {
                            Layout layout = tr.GetObject(layoutEntry.Value, OpenMode.ForRead) as Layout;
                            if (layout != null && layout.BlockTableRecordId.IsValid)
                            {
                                ProcessBlockTableRecord(layout.BlockTableRecordId, tr, renameList);
                            }
                        }

                        tr.Commit();
                    }

                    db.SaveAs(file, DwgVersion.Current);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    /*editor.WriteMessage($"Ошибка при обработке файла {file}. Ошибка: {ex.Message}");*/
                    throw new InvalidOperationException($"Ошибка при обработке файла {file}. Ошибка: {ex.Message}");
                }
                catch (System.Exception ex)
                {
                    /*editor.WriteMessage($"Ошибка при обработке файла {file}. Ошибка: {ex.Message}");*/
                    throw new InvalidOperationException($"Ошибка при обработке файла {file}. Ошибка: {ex.Message}");
                }
                finally
                {
                    db.Dispose();
                }
            }

            foreach (var subDirectory in Directory.GetDirectories(directory))
            {
                DoRename(subDirectory, renameList);
            }
        }

        private void ProcessBlockTableRecord(ObjectId btrId, Transaction tr, List<(string searchName, string replaceName)> renameList)
        {
            BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForWrite) as BlockTableRecord;
            var Check = MainWindow.CheckWildcards;
            Helpers helpers = new Helpers();

            foreach (ObjectId objId in btr)
            {
                DBObject dbObj = tr.GetObject(objId, OpenMode.ForRead);

                if (dbObj is DBText text)
                {
                    foreach (var (searchName, replaceName) in renameList)
                    {
                        if (Check == false && helpers.Contains(text.TextString, searchName))
                        {
                            text.UpgradeOpen();
                            text.TextString = helpers.Replace(text.TextString, searchName, replaceName);
                        }
                        else if (helpers.Comparator(text.TextString, searchName))
                        {
                            text.UpgradeOpen();
                            text.TextString = helpers.Replacer(text.TextString, searchName, replaceName);
                        }
                    }
                }
                else if (dbObj is MText mtext)
                {
                    foreach (var (searchName, replaceName) in renameList)
                    {
                        if (Check == false && helpers.Contains(mtext.Contents, searchName))
                        {
                            mtext.UpgradeOpen();
                            mtext.Contents = helpers.Replace(mtext.Contents, searchName, replaceName);
                        }
                        else if (helpers.Comparator(mtext.Contents, searchName))
                        {
                            mtext.UpgradeOpen();
                            mtext.Contents = helpers.Replacer(mtext.Contents, searchName, replaceName);
                        }
                    }
                }
                else if (dbObj is BlockReference blockRef)
                {
                    if (blockRef.AttributeCollection.Count > 0)
                    {
                        foreach (ObjectId attrId in blockRef.AttributeCollection)
                        {
                            AttributeReference attrRef = tr.GetObject(attrId, OpenMode.ForRead) as AttributeReference;

                            foreach (var (searchName, replaceName) in renameList)
                            {
                                if (attrRef != null && Check == false && helpers.Contains(attrRef.TextString, searchName))
                                {
                                    attrRef.UpgradeOpen();
                                    attrRef.TextString = helpers.Replace(attrRef.TextString, searchName, replaceName);
                                }
                                else if (attrRef != null && helpers.Comparator(attrRef.TextString, searchName))
                                {
                                    attrRef.UpgradeOpen();
                                    attrRef.TextString = helpers.Replacer(attrRef.TextString, searchName, replaceName);
                                }
                            }
                        }
                    }

                    BlockTableRecord blockDef = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                    if (blockDef == null || blockDef.IsLayout) continue;

                    foreach (ObjectId objIdB in blockDef)
                    {
                        DBObject dbObjB = tr.GetObject(objIdB, OpenMode.ForRead);

                        if (dbObjB is DBText textB)
                        {
                            foreach (var (searchName, replaceName) in renameList)
                            {
                                if (Check == false && helpers.Contains(textB.TextString, searchName))
                                {
                                    textB.UpgradeOpen();
                                    textB.TextString = helpers.Replace(textB.TextString, searchName, replaceName);
                                }
                                else if (helpers.Comparator(textB.TextString, searchName))
                                {
                                    textB.UpgradeOpen();
                                    textB.TextString = helpers.Replacer(textB.TextString, searchName, replaceName);
                                }
                            }
                        }
                        else if (dbObjB is MText mtextB)
                        {
                            foreach (var (searchName, replaceName) in renameList)
                            {
                                if (Check == false && helpers.Contains(mtextB.Contents, searchName))
                                {
                                    mtextB.UpgradeOpen();
                                    mtextB.Contents = helpers.Replace(mtextB.Contents, searchName, replaceName);
                                }
                                else if (helpers.Comparator(mtextB.Contents, searchName))
                                {
                                    mtextB.UpgradeOpen();
                                    mtextB.Contents = helpers.Replacer(mtextB.Contents, searchName, replaceName);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
