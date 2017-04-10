using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace AddinConvertTo.Classes
{

    public class Files
    {
        public class Info
        {
            public string FileName { get; set; }
            public string FolderPath { get; set; }
            public string FullFilePath { get; set; }

            public int FolderID { get; set; }
            public int IdPDM { get; set; }
            public int CurrentVersion { get; set; }            

            public int TaskType { get; set; }
            public string ConvertFile { get; set; }
            public string Revision { get; set; }

            public bool ExistDXF { get; set; }
            public bool ExistCutList { get; set; }
            public bool ExistEdrawing { get; set; }
        }

        static public List<Info> FilesPdm(IEdmVault5 vault, ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            var list = new List<Info>();
            try
            {
                foreach (EdmCmdData edmCmdData in ppoData)
                {
                    try
                    {
                        var fileId = edmCmdData.mlObjectID1;
                        var parentFolderId = edmCmdData.mlObjectID2; // for Task
                                                                     //var parentFolderId = edmCmdData.mlObjectID3; // for Addin
                        var fileEdm = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, fileId);
                        var folder = (IEdmFolder5)vault.GetObject(EdmObjectType.EdmObject_Folder, parentFolderId);

                        var taskParam = new Info();
                        taskParam.IdPDM = fileEdm.ID;
                        taskParam.CurrentVersion = fileEdm.CurrentVersion;
                        taskParam.FileName = fileEdm.Name;
                        taskParam.FolderPath = folder.LocalPath;
                        taskParam.FolderID = folder.ID;
                        taskParam.FullFilePath = folder.LocalPath + "\\" + fileEdm.Name;
                        list.Add(taskParam);
                    }
                    catch (Exception exeption)
                    {
                        Logger.Add("FilesPdm - " + exeption.Message);
                    }                   
                }
            }
            catch (COMException ex)
            {
                Logger.Add("OnTaskSetup HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Add(ex.Message + "; " + ex.StackTrace);
            }
            return list;
        }

        static public List<Info> GetFilesToConvert(IEdmVault5 vault, ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            var list = new List<Info>();

            foreach (var item in FilesPdm(vault, ref poCmd, ref ppoData))
            {
                var taskParam = new Info();
                taskParam.IdPDM = item.IdPDM;
                var extension = Path.GetExtension(item.FileName);
                switch (extension.ToUpper())
                {
                    case ".SLDPRT":
                        taskParam.CurrentVersion = item.CurrentVersion;
                        taskParam.FileName = item.FileName;
                        taskParam.FolderPath = item.FolderPath;
                        taskParam.FolderID = item.FolderID;
                        taskParam.FullFilePath = item.FullFilePath;
                        taskParam.ConvertFile = item.FileName.ToUpper().Replace(".SLDPRT", ".EPRT");
                        taskParam.TaskType = 1;
                        list.Add(taskParam);
                        break;
                }
            }
            return list;
        }
    }


}