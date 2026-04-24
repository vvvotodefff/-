using System;
using System.IO.Compression;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using DrawingConfigInterface;
using static FileHelper;
using static System.Net.WebRequestMethods;
using System.Diagnostics;
using OfficeOpenXml;
using System.ComponentModel;
using System.Runtime.InteropServices.ComTypes;
using System.Collections;
using System.Xml.Linq;
using SaveAsDWG;

namespace ЗавестиАнтарус
{

    public class AutoScaleDrawingDimension
    {

        #region Глобальные переменные для класса 
        private static ISldWorks swApp;
        private static DrawingDoc swDrawing;
        private static ModelDoc2 swModel;
        private static string ModelPath;
        private static string DrawingPath;
        #endregion

        /// <summary>
        /// Завершает все процессы из указанного списка.
        /// </summary>
        public static void KillSelectedProcesses(HashSet<string> solidNames)
        {
            solidNames = new HashSet<string>(solidNames, StringComparer.OrdinalIgnoreCase);
            var toKill = Process.GetProcesses()
                            .Where(p => {
                                try { return solidNames.Contains(p.ProcessName); }
                                catch { return false; }                         
                            })
                            .ToList();

            foreach (var proc in toKill)
            {
                string safename;

                try { safename = proc.ProcessName; }
                catch { safename = $"PID_{proc.Id}"; }

                try
                {
                    if (proc.Id == Process.GetCurrentProcess().Id)
                    {
                        proc.Dispose();
                        continue;
                    }

                    proc.Kill();
                    // ждём до 5 секунд, пока процесс завершится
                    if (!proc.WaitForExit(5000))
                    {
                        Log($"Процесс {safename} (ID={proc.Id}) не завершился за 5 секунд.");
                    }
                    else
                    {
                        Log($"Процесс {safename} (ID={proc.Id}) успешно завершён.");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Ошибка при завершении процесса {safename} (ID={proc.Id}): {ex.Message}");
                }
                finally
                {
                    proc.Dispose();
                }
            }
        }




        /// <summary>
        /// Попытка выполнить операцию с записью в лог результата
        /// </summary>
        /// <param name="operation"></param>
        /// <param name="operationName"></param>
        private static bool TryExecute(Action operation, string operationName)
        {
            try
            {
                operation();
                Log($"Попытка выполнить {operationName} успешна");
                return true;

            }
            catch(Exception ex)
            {
                Log($"Попытка  выполнить {operationName} не успешна.Ошибка:{ex.Message}");
                return false;
            }

        }

        /// <summary>
        /// Открывает Solidworks и возвращает приложение SOLIDWORKS
        /// </summary>
        public static class SolidWorksManager
        {
            private static ISldWorks SwApp;

            public static ISldWorks swApp
            {
                get
                {
                    // Если экземпляр уже существует, проверяем его на работоспособность
                    if (SwApp != null)
                    {
                        try
                        {
                            // Попытка обращения к свойству ActiveDoc
                            var doc = SwApp.ActiveDoc;
                        }
                        catch (Exception)
                        {
                            // Если объект не отвечает, сбрасываем его
                            SwApp = null;
                        }
                    }
                    // Если экземпляр отсутствует или был сброшен, создаем новый
                    if (SwApp == null)
                    {
                        SwApp = OpenSolidWorks();
                        SwApp.Visible = true;
                    }
                    return SwApp;
                }
            }

            private static ISldWorks OpenSolidWorks()
            {
                ISldWorks app = null;
                try
                {
                    // Попытка получить запущенный экземпляр SolidWorks
                    app = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (COMException)
                {
                    // Если SolidWorks не запущен, создаем новый экземпляр
                    Type swType = Type.GetTypeFromProgID("SldWorks.Application");
                    app = (ISldWorks)Activator.CreateInstance(swType);
                    app.UserControl = true;
                    // Если необходимо, можно добавить задержку для полной инициализации
                }
                return app;
            }
        }

        /// <summary>
        /// Сохранение файла в SAT формате версии 3.0 (Открывается в AUTOCAD и Revit)
        /// </summary>
        /// <param name="filePath"></param>
        public static void ExportToSat(ISldWorks swApp, string filePath, string PathToSaving)
        {
            if (swApp == null)
            {
                Console.WriteLine("Ошибка: SolidWorks не запущен!");
                return;
            }

            IModelDoc2 modelDoc = swApp.ActiveDoc;

            if (modelDoc == null)
            {
                Console.WriteLine("Ошибка: Откройте документ в SOLIDWORKS!");
                return;
            }

            // Обработка расширения файла
            string newFilePath = filePath.ToUpper().EndsWith(".SLDASM")
                ? filePath.Replace(".SLDASM", ".SAT")
                : filePath.Replace(".SLDPRT", ".SAT");
            string NewFile = $@"{PathToSaving}\3D {Path.GetFileName(PathToSaving)}.SAT";
            int errors = 0;
            int warnings = 0;
            // Сохранение в SAT
            ModelDocExtension modelDoc52 = modelDoc.Extension;
            bool boolstatus = swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swAcisOutputVersion, 4);
            boolstatus = modelDoc52.SaveAs3(
                NewFile,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null,
                0,
                ref errors,
                ref warnings
            );

            if (boolstatus)
                Console.WriteLine($"Файл успешно сохранен: {Path.GetFileName(PathToSaving)}.xlsx");
            else
                Console.WriteLine($"Ошибка сохранения: {errors}, предупреждения: {warnings}");

        }

        /// <summary>
        /// Получение пути сохранения файла
        /// </summary>
        /// <param NameOfProperties=Название свойства></param>
        /// <returns></returns>
        public static string GetCustomProperties(string NameOfProperties)
        {
            string CustomProperties = swModel.GetCustomInfoValue("Установка", NameOfProperties);
            return CustomProperties;
        }

        /// <summary>
        /// Сохранение файла в PDF и DWG в целевую папку
        /// </summary>
        /// <param name="swDrawing">текущий чертеж</param>
        public static void SaveAsPDFandDWG(string Cod1C,string rootdir)
        {
                string FilePath = $@"{rootdir}\{Cod1C}\2D {Cod1C}";
                ((ModelDoc2)swDrawing).SaveAs(FilePath + ".PDF");
                ((ModelDoc2)swDrawing).SaveAs(FilePath + ".DWG");
        }

        /// <summary>
        /// Автоматическая группировка и размещение размеров на виде
        /// </summary>
        /// <param name="swView">Текущий вид</param>
        /// <param name="swModel">активная модель</param>
        public static void AutoPositionDimension(View swView)
        {
            Object[] vDispDim = (Object[])swView.GetDisplayDimensions();
            for (int j = 0; j < vDispDim.Length; j++)
            {
                DisplayDimension swDispDim = (DisplayDimension)vDispDim[j];
                Annotation swAnn = (Annotation)swDispDim.GetAnnotation();
                if ((!swAnn.IsDangling()) & (swAnn.Visible == (int)swAnnotationVisibilityState_e.swAnnotationVisible))
                {
                    swAnn.Select3(true, null);
                }
            }
            ModelDocExtension swModelDocExt = ((ModelDoc2)swDrawing).Extension;
            swModelDocExt.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001);
            ((ModelDoc2)swDrawing).ClearSelection2(true);
        }

        /// <summary>
        /// Выполнение обработки размеров по видам
        /// </summary>
        /// <param name="swDrawing">Чертеж</param>
        /// <param name="NameOfView">Наименования видов</param>
        /// <param name="NotHideDimension">Размеры исключения</param>
        public static void ActionWithDimension(string[] NameOfView, string[] NotHideDimension)
        {
            int i = 0;
            for (i = 0; i < NameOfView.Length; i++)
            {
                ((ModelDoc2)swDrawing).Extension.SelectByID2(NameOfView[i], "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                View swView = ((ModelDoc2)swDrawing).SelectionManager.GetSelectedObject6(1, -1);
                if (i != NameOfView.Length - 1)
                {
                    HideDimensionView(swView, NotHideDimension);
                }
                else 
                { 
                    HideDimensionView(swView, NotHideDimension, specialView: true);
                }

                AutoPositionDimension(swView);
            }
        }

        /// <summary>
        /// Скрыть лишние габаритные размеры
        /// </summary>
        /// <param name="swView">Наименование вида</param>
        /// <param name="NotHideDimension">Размеры исключения</param>
        public static void HideDimensionView(View swView, string[] NotHideDimension, bool specialView = false)
        {
            string MaxValueName = "";
            double MaxValue = 0;
            Object[] vDispDim = (Object[])swView.GetDisplayDimensions();

            for (int j = 0; j < vDispDim.Length; j++)
            {
                DisplayDimension swDispDim = (DisplayDimension)vDispDim[j];
                Annotation swAnn = (Annotation)swDispDim.GetAnnotation();
                Dimension swDim = swDispDim.GetDimension2(0);

                if (!swAnn.IsDangling() && Array.IndexOf(NotHideDimension, swDim.FullName) == -1)
                {
                    swAnn.Visible = (int)swAnnotationVisibilityState_e.swAnnotationVisible;
                    double[] CurrentValue = swDim.GetValue3((int)swInConfigurationOpts_e.swThisConfiguration, null);
                    if (CurrentValue[0] > MaxValue)
                    {
                        MaxValue = CurrentValue[0];
                        MaxValueName = swDim.FullName;
                    }
                }
            }

            bool hideMaxDimension = false;

            if (specialView && NotHideDimension.Length > 0)
            {
                for (int j = 0; j < vDispDim.Length; j++)
                {
                    DisplayDimension swDispDim = (DisplayDimension)vDispDim[j];
                    Dimension swDim = swDispDim.GetDimension2(0);
                    if (swDim.FullName == NotHideDimension[0])
                    {
                        double[] value = swDim.GetValue3((int)swInConfigurationOpts_e.swThisConfiguration, null);
                        
                        if (value[0] >= MaxValue)
                        {
                            hideMaxDimension = true;
                        }

                        break;
                    }
                }
            }

            for (int k = 0; k < vDispDim.Length; k++)
            {
                DisplayDimension swDispDim = (DisplayDimension)vDispDim[k];
                Annotation swAnn = (Annotation)swDispDim.GetAnnotation();
                Dimension swDim = swDispDim.GetDimension2(0);
                if ((swDim.FullName != MaxValueName || hideMaxDimension) && Array.IndexOf(NotHideDimension, swDim.FullName) == -1)
                {
                    swAnn.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
                }
            }
        }

        /// <summary>
        /// Автоматический масштаб чертежа в зависимости от габаритов
        /// </summary>
        /// <returns>Возвращает чертеж</returns>
        public static void AutoScaleDrawing()
        {
            double ScaleValue = double.Parse(GetCustomProperties("Габарит"));
            Sheet swSheet = swDrawing.GetCurrentSheet();
            try
            {
                swSheet.SetScale(1, 5 * Math.Round(ScaleValue / 400), false, false);
            }
            catch
            {
                Console.WriteLine("Не удалось применить масштаб");
            }
        }

        public static void AutoScaleDrawingA4(double ScaleValue)
        {
            Sheet swSheet = swDrawing.GetCurrentSheet();

            try
            {
                swSheet.SetScale(1, ScaleValue, false, false);
            }
            catch
            {
                Console.WriteLine("Не удалось применить масштаб");
            }
        }

        /// <summary>
        /// Возвращает величину масштаба Х к которой нужно привести чертеж формата 1:Х
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public static double GetScaleValue(ModelDoc2 model) 
        {

            AssemblyDoc swAssy = (AssemblyDoc)model;
            object box = swAssy.GetBox((int)swBoundingBoxOptions_e.swBoundingBoxIncludeRefPlanes);
            double[] corners = (double[])box;
            double dx = Math.Abs(corners[3] - corners[0]) / 120.0;
            double dy = Math.Abs(corners[4] - corners[1]) / 70.0;
            double dz = Math.Abs(corners[5] - corners[2]) / 80.0;
            double maxDim = Math.Max(dx, Math.Max(dy, dz)) * 1000.0;
            maxDim = maxDim = Math.Round(maxDim / 5.0) * 5.0;
            return maxDim;
        }

        /// <summary>
        /// Определение типа установки и запуск ActionWithDimension (обработку размеров на чертеже)
        /// </summary>
        /// <param name="swDrawing">Чертеж</param>
        public static void HideExcessDimension()
        {
            string fileName = Path.GetFileName(DrawingPath);
            if (DrawingConfigProvider.Configurations.TryGetValue(fileName, out DrawingConfig config))
            {
                string[] NameOfView = config.NameOfView;
                string[] NotHideDimension = config.NotHideDimension;
                ActionWithDimension(NameOfView, NotHideDimension);
            }
            else
            {
                Console.WriteLine("Выбранный файл не поддерживается");
            }
        }

        /// <summary>
        /// Создание 3D модели в SAT формате
        /// </summary>
        /// <param name="OpenPath"></param>
        /// <param name="OpenPathSAT"></param>
        /// <param name="FinalFolder"></param>
        public static void Create3Dmodel(string OpenPath,string OpenPathSAT,string FinalFolder) 
        {
            swApp = SolidWorksManager.swApp;
            swModel = swApp.OpenDoc6(OpenPath, 2, 0, "", 0, 0);
            ModelPath = Path.Combine(Path.GetDirectoryName(swModel.GetPathName()), Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".SLDASM");
            DrawingPath = Path.Combine(Path.GetDirectoryName(ModelPath), Path.GetFileNameWithoutExtension(ModelPath) + ".SLDDRW");
            swModel.ShowConfiguration("Установка");
            swModel.ForceRebuild3(false);
            string NewFile = $@"{FinalFolder}\3D {Path.GetFileName(FinalFolder)}.IGES";
            SaveAsDwgRunner.Run(swApp, new[] { NewFile });
            //ModelDoc2 swModelSAT = swApp.OpenDoc6(OpenPathSAT, (int)swDocumentTypes_e.swDocASSEMBLY, 1, "", 0, 0);
            //ExportToSat(swApp, OpenPathSAT, FinalFolder);
            //swApp.CloseDoc(OpenPathSAT);
        }

        /// <summary>
        /// Создание 3D и 2D
        /// </summary>
        /// <param name="e"></param>
        public static void CreateDrawingAndModel(string e) 
        {
            

            Log($"Начинается обработка файла: {e}");

            #region Определение переменных
            string rd;
            string filePath = Path.Combine(Costants.RootDirectorySW, Path.GetFileName(e));
            string newFileName = RenameFile(Path.Combine(Costants.RootDirectorySW, Path.GetFileName(e)));
            string Cod1C = GetCodeFromFileName(Path.GetFileNameWithoutExtension(e));
            string OpenPath = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}.SLDASM";
            string DrawOpenPath = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}.SLDDRW";
            string OpenPathSAT = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}(SAT).SLDASM";

            if (e.Contains("НаСерч"))
            {
                rd = Costants.RootDirectorySEARCH;
            }
            else 
            {
                rd = Costants.RootDirectory;
            }

            string FinalFolder = $@"{rd}\{Cod1C}";
            #endregion
            
            CopyFileWithOverwrite(e, newFileName);

            if (!Directory.Exists(FinalFolder))
            {
                Directory.CreateDirectory(FinalFolder);
            }

            if (!TryExecute(() => Create3Dmodel(OpenPath, OpenPathSAT, FinalFolder), "Создать 3D модель"))
            {
                throw new Exception("Ошибка при создании 3D модели");
            }

            swDrawing = (DrawingDoc)swApp.OpenDoc6(DrawingPath, 3, 0, "", 0, 0);

            if (!TryExecute(() => AutoScaleDrawing(), "Автоматическая установка масштаба"))
            {
                throw new Exception("Ошибка при установке масштаба");
            }

            if(!TryExecute(() => HideExcessDimension(),"Скрыть лишние размеры"))
            {
                throw new Exception("Ошибка при скрытии лишних размеров");
            }

            if (!TryExecute(() => SaveAsPDFandDWG(Cod1C,rd), "Сохранить в DWG и PDF"))
            {
                throw new Exception("Ошибка при сохранении в DWG и PDF");
            }
        }

        public static void CreateDrawingBMI(string e)
        {
            Log($"Начинается обработка файла: {e}");

            #region Определение переменных
            string filePath = Path.Combine(Costants.RootDirectorySW, Path.GetFileName(e));
            string newFileName = RenameFile(Path.Combine(Costants.RootDirectorySW, Path.GetFileName(e)));
            string Cod1C = GetCodeFromFileName(Path.GetFileNameWithoutExtension(e));
            string OpenPath = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}.SLDASM";
            string DrawOpenPath = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}.SLDDRW";
            string OpenPathSAT = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}(SAT).SLDASM";
            string FinalFolder = $@"{Costants.RootDirectory}\{Cod1C}";
            string filepathBMI = $@"\\sol.elita\Spec\.CADAutomation\SW\БМИ\БМИ_{Path.GetFileNameWithoutExtension(newFileName)}";
            object DNvsas;
            object DNnapor;
            object JOCKEY;
            object box = null;
            double[] corners = new double[6];
            #endregion

            CopyFileWithOverwrite(e, newFileName);

            if (!Directory.Exists(FinalFolder))
            {
                Directory.CreateDirectory(FinalFolder);
            }

            swApp = SolidWorksManager.swApp;
            swModel = swApp.OpenDoc6(OpenPath, 2, 0, "", 0, 0);
            ModelPath = Path.Combine(Path.GetDirectoryName(swModel.GetPathName()), Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".SLDASM");
            DrawingPath = Path.Combine(Path.GetDirectoryName(ModelPath), Path.GetFileNameWithoutExtension(ModelPath) + ".SLDDRW");
            swModel.ShowConfiguration("Установка");
            swModel.ForceRebuild3(false);
            DrawingConfig config;
            if (!DrawingConfigBMI.Config.TryGetValue(Path.GetFileName(newFileName), out config))
            {
                Console.WriteLine("Нет конфигурации для файла " + newFileName);
                return;
            }

            EnsureTempUsable();
            ExcelPackage.License.SetNonCommercialOrganization("SANEK");
            using (ExcelPackage package = new ExcelPackage(new FileInfo(newFileName)))
            {
                var sheet = package.Workbook.Worksheets[0]; // первый лист

                // Всасывающий
                DNvsas = FindValueUnderMarker(sheet, config.DN_Collector[0]);

                if ((string)DNvsas == "DN32")
                {
                    DNvsas = "DN40";
                }
                // Напорный
                DNnapor = FindValueUnderMarker(sheet, config.DN_Collector[1]);
                
                if ((string)DNnapor == "DN32")
                {
                    DNnapor = "DN40";
                }

                JOCKEY = FindValueUnderMarker(sheet, config.DN_Collector[2]);

            }
            string fileBMI = GetFileBMIParam(@"\\sol.elita\Spec\.CADAutomation\SW\Блок-боксы\");

            if (fileBMI == null)
            {
                return;
            }

            string CodeBMI = GetCodeFromFileName(fileBMI);

            СopyAndDeleteSourceFile(@"\\sol.elita\Spec\.CADAutomation\SW\Блок-боксы\"+fileBMI, Costants.fileConfigBBFronts);

            if ((string)JOCKEY == "Реш")
            {
                JOCKEY = "Непог";
            }
            else if ((string)JOCKEY == null)
            {
                JOCKEY = "Пог";
            }

            Dictionary<string, string> keyBMI = GetKeyValue(Costants.fileConfigBBFronts);

            using (var package = new ExcelPackage(new FileInfo($@"{filepathBMI}.xlsx")))
            {
                var sheet = package.Workbook.Worksheets[0]; // первый лист
                sheet.Cells[3, 2].Value = keyBMI["ВыводПатрубка"];
                sheet.Cells[3, 3].Value = keyBMI["Наименование"];
                sheet.Cells[3, 4].Value = DNnapor;
                sheet.Cells[3, 5].Value = DNvsas;
                sheet.Cells[3, 6].Value = JOCKEY;
                package.Save();
            }
            
            //Запуск программы по применений размеров ББ
            swApp.RunMacro(Costants.fileMacroBBFronts, "БМИ", "CreateBMI");
            swModel = swApp.OpenDoc6(filepathBMI + ".SLDASM", 2, 0, "", 0, 0);
            swModel.ForceRebuild3(false);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            box=swAssy.GetBox((int)swBoundingBoxOptions_e.swBoundingBoxIncludeRefPlanes);
            corners = (double[])box;
            double dx = Math.Abs(corners[3] - corners[0])/120.0;
            double dy = Math.Abs(corners[4] - corners[1])/70.0;
            double dz = Math.Abs(corners[5] - corners[2])/80.0;
            double maxDim = Math.Max(dx, Math.Max(dy, dz)) * 1000.0;
            maxDim= maxDim = Math.Round(maxDim / 5.0) * 5.0;
            swDrawing = (DrawingDoc)swApp.OpenDoc6(filepathBMI + ".SLDDRW", 3, 0, "", 0, 0);
            AutoScaleDrawingA4(maxDim);

            ((ModelDoc2)swDrawing).SaveAs($@"\\sol.elita\Spec\.CADAutomation\Блок-Боксы\БМИ\{CodeBMI}" + ".PDF");
            ((ModelDoc2)swDrawing).SaveAs($@"\\sol.elita\Spec\.CADAutomation\Блок-Боксы\БМИ\{CodeBMI}" + ".DWG");
        }

        /// <summary>
        /// Проверяет доступность TEMP/TMP и при необходимости выставляет рабочую папку на процесс.
        /// </summary>
        static void EnsureTempUsable()
        {
            string temp = Path.GetTempPath();
            try
            {
                string probe = Path.Combine(temp, "epplus_probe.tmp");
                System.IO.File.WriteAllText(probe, "ok");
                System.IO.File.Delete(probe);
                
            }
            catch
            {
                // Если текущий TEMP недоступен — создадим свой и пропишем на процесс
                string fallback = @"C:\Temp";
                Directory.CreateDirectory(fallback);
                System.Environment.SetEnvironmentVariable("TEMP", fallback, EnvironmentVariableTarget.Process);
                System.Environment.SetEnvironmentVariable("TMP", fallback, EnvironmentVariableTarget.Process);
            }
        }

    }
}
