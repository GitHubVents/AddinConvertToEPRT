
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExportPartData;
using EPDM.Interop.epdm;

namespace AddinConvertTo.Classes
{
    public class Batch
    {

        #region Variables Batch
        static IEdmBatchGet batchGetter;
        static IEdmFile5 aFile;
        static IEdmFolder5 aFolder;
        static IEdmPos5 aPos;
        static IEdmSelectionList6 fileList = null;
        static IEdmBatchUnlock batchUnlocker;
        static IEdmBatchAdd poAdder;
        static EdmSelItem[] ppoSelection;
        static string fileNameErr = "";
              

        #endregion

        public static List<Files.Info> UpdateFileInfo(IEdmVault7 vault, List<Files.Info> listType)
        {
            var Fulllist = new List<Files.Info>();

            foreach (var fileVar in listType)
            {
                var filePath = fileVar.FolderPath + "\\" + fileVar.ConvertFile;
                var rev = GetRevision(vault, fileVar.TaskType, filePath);

                string message;

                var udatedFile = new Files.Info
                {
                    CurrentVersion = fileVar.CurrentVersion,
                    FileName = fileVar.FileName,
                    FolderPath = fileVar.FolderPath,
                    FullFilePath = fileVar.FullFilePath,
                    FolderID = fileVar.FolderID,
                    IdPDM = fileVar.IdPDM,
                    ConvertFile = fileVar.ConvertFile,
                    Revision = rev,

                    ExistEdrawing = fileVar.CurrentVersion.ToString() == rev,
                    ExistCutList = ExportXmlSql.ExistXml(fileVar.FileName.ToUpper().Replace(".SLDPRT", ""), fileVar.CurrentVersion, out message),
                    ExistDXF = CheckDxfExistance(vault, fileVar.IdPDM, fileVar.CurrentVersion, fileVar.FileName)
                };
                
                Fulllist.Add(udatedFile);
                //Logger.Add($"FileName: {udatedFile.FileName} Edrw: {udatedFile.ExistEdrawing} Dxf: {udatedFile.ExistDXF} XML: {udatedFile.ExistCutList} Message: {message}");
            }

            return Fulllist;
        }

        static string GetRevision(IEdmVault7 vault, int TaskType, string filePath)
        {
            var variable = "";
            try
            {
                var aFolder = default(IEdmFolder5);
                aFile = vault.GetFileFromPath(filePath, out aFolder);
                object oVarRevision;

                if (aFile == null)
                {
                    variable = "0";
                }
                else
                {
                    var pEnumVar = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable();
                    pEnumVar.GetVar("Revision", "", out oVarRevision);
                    if (oVarRevision == null)
                    {
                        variable = "0";
                    }
                    else
                    {
                        variable = oVarRevision.ToString();
                    }
                }                
            }
            catch (COMException ex)
            {
                Logger.Add("Batch.Get HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);                
            }
            catch (Exception ex)
            {
                Logger.Add("Batch.Get Error " + ex.Message);
            }

            return variable;
        }


        public static bool CheckDxfExistance(IEdmVault7 vault, int IdPDM, int CurrentVersion, string fileName)
        {
            var result = true;
            //Get configurations
            var fileEdm = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, IdPDM);
            EdmStrLst5 cfgList = default(EdmStrLst5);
            cfgList = fileEdm.GetConfigurations();
            string cfgName = null;
            IEdmPos5 pos = default(IEdmPos5);
            pos = cfgList.GetHeadPosition();
            while (!pos.IsNull)
            {
                Exception exc;
                cfgName = cfgList.GetNext(pos);
                if (cfgName != "@")
                {
                    var existDxf = Dxf.ExistDxf(IdPDM, CurrentVersion, cfgName, out exc);
                    //Logger.Add("Проверка на DXF; \n IdPDM: " + IdPDM + "\n Name: " + fileName + "\n Version: " + CurrentVersion + "\n ConfigName: " + cfgName + "\n Exist: " + existDxf);
                    if (!existDxf)
                    {
                        //return 
                            result = false;
                    }
                }
            }
            return result;
        }

        public static void ClearLocalCache(IEdmVault7 vault, List<Files.Info> listType)
        {
            try
            {
                var ClearLocalCache = (IEdmClearLocalCache3)vault.CreateUtility(EdmUtility.EdmUtil_ClearLocalCache);
                ClearLocalCache.IgnoreToolboxFiles = true;

                //Declare and create the IEdmBatchListing object
                var BatchListing = (IEdmBatchListing)vault.CreateUtility(EdmUtility.EdmUtil_BatchList);

                foreach (var item in listType)
                {
                    ClearLocalCache.AddFileByPath(item.FullFilePath);
                    //((IEdmBatchListing2)BatchListing).AddFileCfg(KvPair.Key, DateTime.Now, (Convert.ToInt32(KvPair.Value)), "@", Convert.ToInt32(EdmListFileFlags.EdmList_Nothing));
                }

                //Clear the local cache of the reference files
                ClearLocalCache.CommitClear();
            }
            catch (COMException ex)
            {
                Logger.Add("ERROR ClearLocalCache файл: " + fileNameErr + " HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);//, 10001);
            }
        }

        public static bool Get(IEdmVault7 vault, List<Files.Info> files)
        {
            var status = false;
            Logger.Add($"Получение {files?.Count} файлов из PDM");
            try
            {
                batchGetter = (IEdmBatchGet)vault.CreateUtility(EdmUtility.EdmUtil_BatchGet);
                foreach (var file in files)
                {
                    try
                    {
                        batchGetter.AddSelectionEx((EdmVault5)vault, file.IdPDM, file.FolderID, file.CurrentVersion);
                    }
                    catch (Exception ex)
                    {
                        Logger.Add($"Ошибка при получении:{ex.Message} Path: {file.FullFilePath} IdPDM: {file.IdPDM} FolderID: {file.FolderID} Version: {file.CurrentVersion}");
                    }
                }
                if ((batchGetter != null))
                {                 
                    batchGetter.CreateTree(0, (int)EdmGetCmdFlags.Egcf_SkipUnlockedWritable);
                    batchGetter.GetFiles(0, null);
                }
            }
            catch (COMException ex)
            {
                Logger.Add("ERROR BatchGet HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Add("ERROR BatchGet HRESULT = 0x" + ex.StackTrace + "\n" + ex.Message);
            }
            return status;
        }

        public static void Delete(IEdmVault7 vault, List<Files.Info> files)
        {
            Logger.Add($"Удаление файлов из PDM");

            EdmBatchDelErrInfo[] ppoDelErrors = null;
            try
            {
                var batchDeleter = (IEdmBatchDelete3)vault.CreateUtility(EdmUtility.EdmUtil_BatchDelete);

                foreach (var file in files)
                {
                    if (!file.ExistEdrawing)
                    {
                        var filePath = file.FolderPath + "\\" + file.ConvertFile;

                        // Add selected file to the batch

                        try
                        {
                            batchDeleter.AddFileByPath(filePath);
                        }
                        catch (Exception ex)
                        {
                            Logger.Add($"Ошибка при удалении: {ex.Message} Path: {file.FullFilePath} IdPDM: {file.IdPDM} FolderID: {file.FolderID} Version: {file.CurrentVersion}");
                        }
                    }
                }
                batchDeleter.ComputePermissions(true, null);
                var retVal = batchDeleter.CommitDelete(0, null);
                if ((!retVal))
                {
                    batchDeleter.GetCommitErrors(ppoDelErrors);
                }
            }
            catch (COMException ex)
            {
                Logger.Add("ERROR BatchDelete - " + ppoDelErrors + " - HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);//, 10001);
            }

        }


        public static void AddFiles(IEdmVault7 vault, List<Files.Info> files)
        {
            Logger.Add($"Добавление в PDM {files?.Count} файлов");
            string msg = "";

            try
            {
                var result = default(bool);
                poAdder = (IEdmBatchAdd)vault.CreateUtility(EdmUtility.EdmUtil_BatchAdd);
                foreach (var file in files)
                {
                    fileNameErr = file.FolderPath + "\\" + file.FileName;
                    try
                    {
                        poAdder.AddFileFromPathToPath(file.FolderPath + "\\" + file.ConvertFile, file.FolderPath, 0, "", 0);
                    }
                    catch (Exception ex)
                    {
                        Logger.Add($"Ошибка при добавлении: {ex.Message} Path: {file.FullFilePath} IdPDM: {file.IdPDM} FolderID: {file.FolderID} Version: {file.CurrentVersion}");
                    }
                }

                var edmFileInfo = new EdmFileInfo[files.Count];

                result = Convert.ToBoolean(poAdder.CommitAdd(0, edmFileInfo, 0));
                var idx = edmFileInfo.GetLowerBound(0);

                while (idx <= edmFileInfo.GetUpperBound(0))
                {
                    string row = null;
                    row = "(" + edmFileInfo[idx].mbsPath + ") arg = " + Convert.ToString(edmFileInfo[idx].mlArg);

                    if (edmFileInfo[idx].mhResult == 0)
                    {
                        row = row + " status = OK " + edmFileInfo[idx].mbsPath;
                    }
                    else
                    {
                        string oErrName = "";
                        string oErrDesc = "";

                        vault.GetErrorString(edmFileInfo[idx].mhResult, out oErrName, out oErrDesc);
                        row = row + " status = " + oErrName;
                    }

                    idx = idx + 1;
                    msg = msg + "\n" + row;
                }

                //  statusChank = true;

            }
            catch (COMException ex)
            {
                Logger.Add("ERROR BatchAddFiles " + msg + ", file: " + fileNameErr + " HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
                //  statusChank = false;
            }

            #region To delete

            //try
            //{
            //    var result = default(bool);
            //    poAdder = (IEdmBatchAdd)vault.CreateUtility(EdmUtility.EdmUtil_BatchAdd);
            //    foreach (var fileName in listType)
            //    {
            //        fileNameErr = fileName.FolderPath + "\\" + fileName.FileName;
            //        poAdder.AddFileFromPathToPath(fileName.FolderPath + "\\" + fileName.ConvertFile, fileName.FolderPath, 0, "", 0);
            //    }
            //    var edmFileInfo = new EdmFileInfo[listType.Count];

            //    result = Convert.ToBoolean(poAdder.CommitAdd(0, edmFileInfo, 0));
            //    var idx = edmFileInfo.GetLowerBound(0);
            //    while (idx <= edmFileInfo.GetUpperBound(0))
            //    {
            //        string row = null;
            //        row = "(" + edmFileInfo[idx].mbsPath + ") arg = " + Convert.ToString(edmFileInfo[idx].mlArg);

            //        if (edmFileInfo[idx].mhResult == 0)
            //        {
            //            row = row + " status = OK " + edmFileInfo[idx].mbsPath;
            //        }
            //        else
            //        {
            //            string oErrName = "";
            //            string oErrDesc = "";

            //            vault.GetErrorString(edmFileInfo[idx].mhResult, out oErrName, out oErrDesc);
            //            row = row + " status = " + oErrName;
            //        }

            //        idx = idx + 1;
            //        msg = msg + "\n" + row;
            //    }
            //    statusChank = true;
            //}
            //catch (COMException ex)
            //{
            //    Logger.ToLog("ERROR BatchAddFiles " + msg + ", file: " + fileNameErr + " HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message, 10001);
            //    statusChank = false;
            //}

            #endregion

        }

        public static void SetVariable(IEdmVault7 vault, List<Files.Info> files)
        {
            try
            {
                foreach (var fileVar in files)
                {
                    var filePath = fileVar.FolderPath + "\\" + fileVar.ConvertFile;
                    fileNameErr = filePath;
                    IEdmFolder5 folder;
                    aFile = vault.GetFileFromPath(filePath, out folder);

                    var pEnumVar = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable(); ;
                    pEnumVar.SetVar("Revision", "", fileVar.CurrentVersion);
                }                
            }
            catch (COMException ex)
            {                
                Logger.Add("ERROR BatchSetVariable файл: " + fileNameErr + ", " + ex.Message);
            }                  
        }


        public static void UnLock(IEdmVault7 vault, List<Files.Info> files)
        {
            Logger.Add($"Начало регистрации {files?.Count} файлов.");

            try
            {
                ppoSelection = new EdmSelItem[files.Count];
                batchUnlocker = (IEdmBatchUnlock)vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);
                var i = 0;
                foreach (var file in files)
                {
                    try
                    {
                        var filePath = file.FolderPath + "\\" + file.ConvertFile;
                        fileNameErr = filePath;
                        IEdmFolder5 folder = default(IEdmFolder5);
                        aFile = vault.GetFileFromPath(filePath, out folder);
                        aPos = aFile.GetFirstFolderPosition();
                        aFolder = aFile.GetNextFolder(aPos);

                        ppoSelection[i] = new EdmSelItem();
                        ppoSelection[i].mlDocID = aFile.ID;
                        ppoSelection[i].mlProjID = aFolder.ID;

                        i = i + 1;
                    }
                    catch (Exception ex)
                    {
                        Logger.Add($"Ошибка при добавлении файла на регистрацию: {ex.Message} Path: {file.FullFilePath} IdPDM: {file.IdPDM} FolderID: {file.FolderID} Version: {file.CurrentVersion}");
                    }
                }

                // Add selections to the batch of files to check in
                batchUnlocker.AddSelection((EdmVault5)vault, ppoSelection);
                if ((batchUnlocker != null))
                {
                    batchUnlocker.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock);
                    fileList = (IEdmSelectionList6)batchUnlocker.GetFileList((int)EdmUnlockFileListFlag.Euflf_GetUnlocked + (int)EdmUnlockFileListFlag.Euflf_GetUndoLocked + (int)EdmUnlockFileListFlag.Euflf_GetUnprocessed);
                    batchUnlocker.UnlockFiles(0, null);
                }
            }
            catch (COMException ex)
            {
                Logger.Add("ERROR BatchUnLock файл: '" + fileNameErr + "', " + ex.StackTrace + " " + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Add("ERROR BatchUnLock: '" + fileNameErr + "', " + ex.StackTrace + " " + ex.Message);
            }

            Logger.Add($"Завершена регистрации {files?.Count} файлов.");
        }
    }

}