using System;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices; 
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using AddinConvertTo.Classes;
using System.Collections.Generic;
 using ExportPartData;
using EPDM.Interop.epdm;

namespace NewAddinToPDM
{
    public class ClassHome : IEdmAddIn5
    {
        private bool ThisIsTask = true;

        #region Variables

        private SldWorks swApp;
        private Process[] processes;
        private IEdmTaskInstance edmTaskInstance;

        private int currentVer;
        private int filesToProceed;
        private Exception ex;

        private List<Files.Info> filesPdm { get; set; }

        private List<Files.Info> FilesList { get; set; }

        #endregion

        #region EdmAddIn

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            // Version
            const int ver = 14;

            try
            {
                if (ThisIsTask)
                {
                    #region Task 

                    poInfo.mbsAddInName = $"Make .eprt files Task Add-In ver." + ver;
                    poInfo.mbsCompany = "Vents";
                    poInfo.mbsDescription = "Создание и сохранение файлов .eprt для листовых деталей";
                    poInfo.mlAddInVersion = ver;
                    currentVer = poInfo.mlAddInVersion;

                    //Minimum SolidWorks Enterprise PDM version
                    //needed for C# Task Add-Ins is 10.0
                    poInfo.mlRequiredVersionMajor = 10;
                    poInfo.mlRequiredVersionMinor = 0;

                    //Register this add-in as a task add-in
                    poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
                    //Register this add-in to be called when
                    //selected as a task in the Administration tool
                    poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
                    poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton);
                    poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskDetails);

                    #endregion
                }
                else
                {
                    #region Addin 

                    // Addin version
                    poInfo.mbsAddInName = "C# Task Add-In";
                    poInfo.mbsCompany = "Vents";
                    poInfo.mbsDescription = "PDF & EPRT";
                    poInfo.mlAddInVersion = 1;
                    poInfo.mlRequiredVersionMajor = 6;
                    poInfo.mlRequiredVersionMinor = 4;
                    poCmdMgr.AddCmd(1, "Convert", (int)EdmMenuFlags.EdmMenu_Nothing, "", "", 0, 0);

                    #endregion
                }
            }
            catch (COMException ex)
            {
                Logger.Add("GetAddInInfo HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Add(ex.Message + "; \n" + ex.StackTrace);
            }
        }

        /// <summary>
        /// Begins the task.
        /// </summary>
        /// <param name="poCmd"></param>
        /// <param name="ppoData"></param>
        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {

                #region Task Addin version

                switch (poCmd.meCmdType)
                {
                    case EdmCmdType.EdmCmd_TaskRun:
                        OnTaskRun(ref poCmd, ref ppoData);
                        break;
                    case EdmCmdType.EdmCmd_TaskSetup:
                        OnTaskSetup(ref poCmd, ref ppoData);
                        break;
                    case EdmCmdType.EdmCmd_TaskDetails:
                        break;
                    case EdmCmdType.EdmCmd_TaskSetupButton:
                        break;
                }

                #endregion

                #region Addin version

                //vault = poCmd.mpoVault as IEdmVault7;
                //switch (poCmd.meCmdType)
                //        {
                //            case EdmCmdType.EdmCmd_Menu:
                //                OnTaskRun(ref poCmd, ref ppoData);
                //                break;
                //        }

                #endregion

            }
            catch (COMException ex)
            {
                Logger.Add("OnCmd HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Add("OnCmd = " + ex.Message + "; " + ex.StackTrace);
            }
        }

        #endregion                

        #region Task Addin

        static void OnTaskSetup(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                //Get the property interface used to
                //access the framework
                var edmTaskProperties = (IEdmTaskProperties)poCmd.mpoExtra;
                //Set the property flag that says you want a
                //menu item for the user to launch the task
                //and a flag to support scheduling
                edmTaskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitForm + (int)EdmTaskFlag.EdmTask_SupportsDetails + (int)EdmTaskFlag.EdmTask_SupportsChangeState;
                //edmTaskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsChangeState + (int)EdmTaskFlag.EdmTask_SupportsInitExec;
                //Set up the menu commands to launch this task
                var edmTaskMenuCmds = new EdmTaskMenuCmd[1];
                edmTaskMenuCmds[0].mbsMenuString = "Выгрузить eDrawing";
                edmTaskMenuCmds[0].mbsStatusBarHelp = "Выгрузить eDrawing";
                edmTaskMenuCmds[0].mlCmdID = 1;
                edmTaskMenuCmds[0].mlEdmMenuFlags = (int)EdmMenuFlags.EdmMenu_Nothing;
                edmTaskProperties.SetMenuCmds(edmTaskMenuCmds);
            }
            catch (COMException ex)
            {
                Logger.Add("OnTaskSetup HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Add("OnTaskSetup Error" + ex.Message + ppoData);
            }
        }

        private void InitializeErrorList()
        {
            ListWithConvertErrors = null;
            ListWithConvertErrors = new List<Files.Info>();
        }

        List<Files.Info> ListWithConvertErrors;

        void OnTaskRun(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            var statusTask = true;

            edmTaskInstance = (IEdmTaskInstance)poCmd.mpoExtra;
            edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_Running);
            List<Files.Info> filesToConvert = null;

            try
            {

                var vault = poCmd.mpoVault as IEdmVault7;

                InitializeErrorList();

                filesToProceed = 1;

                filesPdm = Files.GetFilesToConvert(vault, ref poCmd, ref ppoData);
                FilesList = Batch.UpdateFileInfo(vault, filesPdm);

                filesToConvert = FilesList.Where(x => !x.ExistCutList | !x.ExistDXF /*| !x.ExistEdrawing*/).ToList();

                Logger.Add($@"\n\n{new string('-', 500)}\nTime {DateTime.Now}\nЗадача '{edmTaskInstance.TaskName}'
                            ID: {edmTaskInstance.ID}, для {filesToConvert.Count} элемента(ов), OnTaskRun ver.{currentVer}");

                edmTaskInstance.SetProgressRange(filesToConvert.Count, 0, $"Запуск задачи в {DateTime.Now}");

                GetAndDeleteUnnecesseryFromPdm(vault, filesToConvert);

                // Run SldWorks

                if (filesToConvert.Count != 0)
                {
                    KillSwAndStartNewProccess();

                    // Разбиваем лист на группы
                    var nChunks = 1;
                    var totalLength = filesToConvert.Count();
                    var chunkLength = (int)Math.Ceiling(totalLength / (double)nChunks);
                    var partsToList = Enumerable.Range(0, chunkLength).Select(i => filesToConvert.Skip(i * nChunks).Take(nChunks).ToList()).ToList();

                    for (var i = 0; i < partsToList.Count; i++)
                    {
                        if (!ConvertChunk(vault, partsToList[i]))
                        {
                            statusTask = false;
                        }
                    }
                }
            }
            catch (COMException ex)
            {
                statusTask = false;
                this.ex = ex;
                Logger.Add("OnTaskRun HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.StackTrace + "; \n" + ex.Message);
            }
            catch (Exception ex)
            {
                statusTask = false;
                this.ex = ex;
                Logger.Add(ex.Message + ";\n" + ex.StackTrace);
            }
            finally
            {
                if (ListWithConvertErrors?.Count > 0)
                {
                    statusTask = false;
                    try
                    {
                        var message = "";
                        foreach (var item in ListWithConvertErrors)
                        {
                            message = message + "\n" + $"FileName - {item.FileName} FolderID - {item.FolderID} IdPDM - {item.IdPDM} Revision {item.Revision}";
                        }
                        Logger.Add(message);

                    }
                    catch (Exception)
                    {
                        Logger.Add("Ошибки при выгрузке данных об ошибке");
                    }
                }

                edmTaskInstance.SetProgressPos(filesToConvert.Count, "Выполнено, ID = " + edmTaskInstance.ID);
                edmTaskInstance.SetStatus(statusTask ? EdmTaskStatus.EdmTaskStat_DoneOK : EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode());
            }
        }

        private void GetAndDeleteUnnecesseryFromPdm(IEdmVault7 vault, List<Files.Info> files)
        {
            Batch.Get(vault, files);
            Batch.Delete(vault, files);
        }

        private void KillSwAndStartNewProccess()
        {
            processes = Process.GetProcessesByName("SLDWORKS");
            foreach (var process in processes)
            {
                process.Kill();
            }
            swApp = new SldWorks() { Visible = true };
        }

        #endregion

        public bool ConvertChunk(IEdmVault7 vault, List<Files.Info> filesInChunk)
        {
            Logger.Add("");
            try
            {
                AddFilesToPdm(vault, Convert(swApp, filesInChunk));
                return true;
            }
            catch (Exception ex)
            {
                Logger.Add("Convert: " + ex.Message + "; " + ex.StackTrace);
                return false;
            }
        }

        private static void AddFilesToPdm(IEdmVault7 vault, List<Files.Info> ForBatchAdd)
        {
            Batch.AddFiles(vault, ForBatchAdd);
            Batch.SetVariable(vault, ForBatchAdd);
            Batch.UnLock(vault, ForBatchAdd);
        }

        public List<Files.Info> Convert(SldWorks swApp, List<Files.Info> filesToConvert)
        {
            var ListForBatchAdd = new List<Files.Info>();
            var fileNameErr = "";

            foreach (var item in filesToConvert)
            {
                fileNameErr = item.FullFilePath;
                edmTaskInstance.SetProgressPos(filesToProceed++, item.FullFilePath);

                Logger.Add($"Open in Solidworks : {item.FileName}");

                ModelDoc2 swModel = swApp.OpenDoc6(item.FullFilePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);

                // Запуск конвертации только для литовых деталей

                if (IsSheetMetalPart(swModel))
                {
                    Exception exOut = null;

                    var confArray = (object[])swModel.GetConfigurationNames();

                    bool isFix = false;
                    ExportDataToXmlSql(swApp, item, swModel, ref exOut, confArray, out isFix);
                    ConvertToDxf(swApp, item, swModel, ref exOut, confArray, isFix);

                    //   ConvertToEprt(swApp, ListForBatchAdd, item, swModel);
                }

                Logger.Add($"CloseDoc: {item.FileName}");
                swApp.CloseDoc(item.FullFilePath);

            }

            return ListForBatchAdd;
        }

        private static void ExportDataToXmlSql(SldWorks swApp, Files.Info item, ModelDoc2 swModel, ref Exception exOut, object[] confArray , out bool isFix)
        {
            if (!item.ExistCutList)
            {
                try
                {
                    List<ExportXmlSql.DataToExport> listCutList = new List<ExportXmlSql.DataToExport>();                                                       // using export part data  

                    // Проходимся по всем конфигурация для Fix, CutList
                    foreach (var confName in confArray)
                    {
                        Configuration swConf = swModel.GetConfigurationByName(confName.ToString());
                        if (swConf.IsDerived()) continue; // только детали верхнего уровня
                        swModel.ShowConfiguration2(confName.ToString());

                        // Проверка на CutList: false - XML отсутствует, true - XML есть
                        if (!item.ExistCutList)
                        { 
                            var list = new List<Bends.SolidWorksFixPattern.PartBendInfo>();                                                                   // using export part data  
                            Bends.Fix(swApp, out list, false); // разгиб 
                            
                            List<ExportXmlSql.DataToExport> listCutListConf;                                                                                    // using export part data  
                            ExportXmlSql.GetCurrentConfigPartData(swApp, item.CurrentVersion, item.IdPDM, false, false, out listCutListConf, out exOut);
                            listCutList.AddRange(listCutListConf);
                        }
                        if (exOut != null)
                        {
                            Logger.Add(exOut.Message);
                        }
                    }

                    isFix = true;

                    //CutList To Sql
                    if (listCutList != null)
                    {
                        //Logger.Add("=============================== Сохранение Cut list data через  ExportXmlSql.ExportDataToXmlSql метод =====================================");
                        //System.Text.StringBuilder messageBuilder = new System.Text.StringBuilder();
                        //foreach (var EachCut in listCutList)
                        //{
                        //    messageBuilder.Append("id " + EachCut.IdPdm+", ");

                        //    messageBuilder.Append("name " + EachCut.FileName + ", ");

                        //    messageBuilder.Append("ver. " + EachCut.Version + ", ");

                        //    messageBuilder.Append("conf. " + EachCut.Config + ", ");

                        //    messageBuilder.Append("Mat-l Id " + EachCut.MaterialId + ", ");

                        //    messageBuilder.Append("PaintX " + EachCut.PaintX + ", ");
                        //    messageBuilder.Append("PaintY " + EachCut.PaintY + ", ");
                        //    messageBuilder.Append("PaintZ " + EachCut.PaintZ + ", ");

                        //    messageBuilder.Append("ДлинаГраничнойРамки " + EachCut.ДлинаГраничнойРамки + ", ");
                        //    messageBuilder.Append("КодМатериала " + EachCut.КодМатериала + ", ");
                        //    messageBuilder.Append("Материал " + EachCut.Материал + ", ");
                        //    messageBuilder.Append("ПлощадьПокрытия " + EachCut.ПлощадьПокрытия + ", ");
                        //    messageBuilder.Append("Сгибы " + EachCut.Сгибы + ", ");

                        //    messageBuilder.Append("ТолщинаЛистовогоМеталла " + EachCut.ТолщинаЛистовогоМеталла + ", ");
                        //    messageBuilder.Append("ШиринаГраничнойРамки " + EachCut.ШиринаГраничнойРамки + ", ");
                        //    messageBuilder.Append("\n============================================================================================\n");
                        //}

                        //Logger.Add(messageBuilder.ToString());

                        ExportXmlSql.ExportDataToXmlSql(item.FileName.ToUpper().Replace(".SLDPRT", ""), listCutList, out exOut);                                            // using export part data  
                        if (exOut != null)
                        {
                            Logger.Add(exOut.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Add("!item.ExistCutList ------ " + ex.Message + "; " + ex.StackTrace);
                }
            }
            isFix = false;
        }

        private void ConvertToDxf(SldWorks swApp, Files.Info item, ModelDoc2 swModel, ref Exception exOut, object[] confArray, bool IsBendsFixed)
        {
            if (!item.ExistDXF)
            {
                // Проходимся по всем конфигурациям для Fix, DXF
                foreach (var confName in confArray)
                {
                    // Проверка на Dxf
                    if (!item.ExistDXF)
                    {
                        Exception exc;
                        if (!Dxf.ExistDxf(item.IdPDM, item.CurrentVersion, confName.ToString(), out exc))                                                       // using export part data  
                        {
                            swModel.ShowConfiguration2(confName.ToString());
                            if (!IsBendsFixed)
                            {
                                var list = new List<Bends.SolidWorksFixPattern.PartBendInfo>();
                                Bends.Fix(swApp, out list, false);                                                                                              // using export part data  
                            } 
                            var listDxf = new List<Dxf.DxfFile>();
                            if (Dxf.Save(swApp, out exOut, item.IdPDM, item.CurrentVersion, out listDxf, false, false, confName.ToString()))                    // using export part data  
                            {
                                if (exOut != null)
                                {
                                    Logger.Add(exOut.Message + "; " + exOut.StackTrace);
                                }
                                var exOutList = new List<Dxf.ResultList>();
                                Dxf.AddToSql(listDxf, true, out exOutList);
                                if (exOutList != null)
                                {
                                    foreach (var itemEx in exOutList)
                                    {
                                        Logger.Add($"DXF AddToSql err: {itemEx.dxfFile} Exception {itemEx.exc}");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ConvertToEprt(SldWorks swApp, List<Files.Info> ListForBatchAdd, Files.Info item, ModelDoc2 swModel)
        {
            if (!item.ExistEdrawing)
            {
                try
                {
                    if (ConvertToErpt(swApp, swModel, item.FullFilePath))
                    {
                        ListForBatchAdd.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Add(ex.Message + "; " + ex.StackTrace);
                }
            }
        }

        static bool IsSheetMetalPart(ModelDoc2 swModel)
        {
            var isSheet = false;
            try
            {
                var swPart = (PartDoc)swModel;
                if (swPart != null)
                {
                    var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
                    foreach (Body2 vBody in vBodies)
                    {
                        var isSheetMetal = vBody.IsSheetMetal();
                        if (!isSheetMetal) continue;
                        isSheet = true;
                    }
                }
            }
            catch (Exception)
            {
                isSheet = false;
                Logger.Add($"Part is not sheet metal");
            }
            return isSheet;
        }

        bool ConvertToErpt(SldWorks swApp, ModelDoc2 swModel, string filePath)
        {
            var result = false;
            try
            {
                ChangesTheVisibilityOfItems(swModel);
                swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption, (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                var fileParthEprt = filePath.ToUpper().Replace("SLDPRT", "EPRT");
                swModel.Extension.SaveAs(fileParthEprt, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                result = true;
            }
            catch (Exception ex)
            {
                Logger.Add("ERROR CONVERT TO ERPT + " + ex.Message + "; " + ex.StackTrace);
            }
            return result;
        }

        void ChangesTheVisibilityOfItems(ModelDoc2 swModel)
        {
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplaySketches)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swViewSketchRelations)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayCurves)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayCompAnnotations)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayReferencePoints2)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayOrigins)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayPlanes)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayReferencePoints)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayLiveSections)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayLights)), false);
            swModel.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swDisplayCenterOfMassSymbol)), false);
        }


    }
}