using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static ЗавестиАнтарус.AutoScaleDrawingDimension;
using static FileHelper;
using System.Runtime.InteropServices;
using System.Threading;
using static VZU.CreateDrawingVZU;
using System.Collections.Generic;
using Microsoft.Win32;

namespace ВыполнитьЗадачиSolidWorks
{
        /// <summary>
        /// Обнаружение новых файлов
        /// </summary>
    public class AntarusFileWatcher
    {
       
        private FileSystemWatcher watcher;
        private readonly BlockingCollection<string> jobQueue = new BlockingCollection<string>();
        private readonly HashSet<string> queuedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object queuedFilesLock = new object();
        private Thread processingThread;

        private const string WorkerMode = "--worker";
        private const string SelfTestMode = "--self-test";
        public HashSet<string> officeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "winword",
                "wordconv",
                "excel"
            };

        public HashSet<string> solidNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "sldworks",
                "sldworks_fs",
                "sldprocmon",
                "swcefsubproc",
                "sldexitapp"
            };

        public HashSet<string> acadNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "acad",
                "accoreconsole"
            };
        /// <summary>
        /// Действие программы при запуске
        /// </summary>
        public void Start()
        {
            processingThread = new Thread(ProcessQueue)
            {
                IsBackground = true,
                Name = "CAD automation queue"
            };
            processingThread.Start();

            watcher = new FileSystemWatcher
            {
                Path = Costants.OnPerfFilePath,
                //Filter = "*.",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };

            watcher.Created += OnFileCreated;
            watcher.EnableRaisingEvents = true;
            ProcessExistingFiles();
            Log("Консольное приложение запущено, мониторинг папки: " + Costants.OnPerfFilePath);
            Console.WriteLine("Нажмите Enter для завершения работы.");
            Console.ReadLine();
            watcher.EnableRaisingEvents = false;
            jobQueue.CompleteAdding();
        }

        /// <summary>
        /// Действия программы при обнаружении файла
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void OnFileCreated(object source, FileSystemEventArgs e)
        {
            if (e.FullPath.Contains("~$") || e.FullPath.Contains(".tmp")) return;
            EnqueueFile(e.FullPath);
        }

        private void EnqueueFile(string fullPath)
        {
            lock (queuedFilesLock)
            {
                if (!queuedFiles.Add(fullPath))
                {
                    Log($"Файл уже стоит в очереди: {fullPath}");
                    return;
                }
            }

            jobQueue.Add(fullPath);
            Log($"Файл добавлен в очередь: {fullPath}");
        }

        private void ProcessQueue()
        {
            foreach (string fullPath in jobQueue.GetConsumingEnumerable())
            {
                try
                {
                    ProcessFile(fullPath);
                }
                catch (Exception ex)
                {
                    Log($"Критическая ошибка обработки файла {fullPath}: {ex}");
                }
                finally
                {
                    lock (queuedFilesLock)
                    {
                        queuedFiles.Remove(fullPath);
                    }
                }
            }
        }

        private void ProcessFile(string fullPath)
        {

            //#region На время пока выполняем на серч
            //bool ЗавестиАнтарус = false;
            //string dir = Path.GetDirectoryName(fullPath);
            //string[] files = Directory.GetFiles(dir);

            //foreach (string file in files)
            //{
            //    if (file.Contains("ЗавестиАнтарус.xlsx") && !file.Contains("~$"))
            //    {
            //        ЗавестиАнтарус = true;
            //        RunWithTimeout(
            //        () => SolidWorksManager.swApp,  // Инициализация SolidWorks
            //        swApp => CreateDrawingAndModel(file), // Вызов метода с параметрами
            //        "Не удалось создать 3D модель и(или) чертеж на: " + file,
            //        80,// Таймаут в секундах
            //        file
            //        );
            //    }
            //}
            //#endregion

            Log($"Начинается обработка файла из очереди: {fullPath}");

            string jobType = CadJobClassifier.GetJobType(fullPath);
            if (jobType == null)
            {
                Console.WriteLine("Неизвестный тип файла: " + fullPath);
                Log("Неизвестный тип файла: " + fullPath);
                return;
            }

            if (!WaitForFileReady(fullPath, TimeSpan.FromMinutes(3)))
            {
                Log("Файл не стал доступен для обработки: " + fullPath);
                MoveProcessedFile(fullPath, false);
                return;
            }

            //Запуск программы создания 3D и 2D моделей Антарус
            if (jobType == CadJobType.Antarus) // Если заводим на серч то добавляем "&& !ЗавестиАнтарус"
            {
                CloseOfficeProcesses();
                RunWorkerWithTimeout(
                CadJobType.Antarus,
                "Не удалось создать 3D модель и(или) чертеж на: " + fullPath,
                80,// Таймаут в секундах
                fullPath,
                killSolidWorks: true,
                killOffice: true
                );
            }

            else if (jobType == CadJobType.Bmi)
            {
                RunWorkerWithTimeout(
                CadJobType.Bmi,
                "Не удалось создать 3D модель и(или) чертеж на: " + fullPath,
                800,// Таймаут в секундах
                fullPath,
                killSolidWorks: true,
                killOffice: true
                );
            }

            //Запуск программы расчета нагрузок ББ
            else if (jobType == CadJobType.StaticLoad)
            {
               KillSelectedProcesses(solidNames);
               CloseOfficeProcesses();
                RunWorkerWithTimeout(
               CadJobType.StaticLoad,
               "Не удалось создать 3D модель и(или) чертеж на: " + fullPath,
               300,// Таймаут в секундах
                fullPath,
                killSolidWorks: true,
                killOffice: true
               );
            }

            //Запуск программы чертежей фасадов
            else if (jobType == CadJobType.Fronts)
            {
                KillSelectedProcesses(solidNames);
                RunWorkerWithTimeout(
                CadJobType.Fronts,
                "Не удалось создать 3D модель и(или) чертеж на: " + fullPath,
                120,// Таймаут в секундах
                fullPath,
                killSolidWorks: true,
                killOffice: false
                );
            }

            else if (jobType == CadJobType.Vzu)
            {
                RunWorkerWithTimeout(
                CadJobType.Vzu,
                "Не удалось создать чертеж на: " + fullPath,
                200,// Таймаут в секундах
                fullPath,
                killSolidWorks: false,
                killOffice: true,
                killAutoCad: true
                );
            }

            else if (jobType == CadJobType.Search)
            {
                //KillSelectedProcesses(solidNames);
                RunWorkerWithTimeout(
                CadJobType.Search,
                "Не удалось создать 3D модель и(или) чертеж на: " + fullPath,
                80,// Таймаут в секундах
                fullPath,
                killSolidWorks: true,
                killOffice: true
                );
            }
        }

        /// <summary>
        /// Поиск имеющихся файлов в папке
        /// </summary>
        private void ProcessExistingFiles()
        {
            
            try
            {
                string[] files = Directory.GetFiles(Costants.OnPerfFilePath);

                foreach (var filePath in files)
                {
                    
                    Log($"Обнаружен существующий файл: {Path.GetFileName(filePath)}");
                    EnqueueFile(filePath);
                    
                }
            }
            catch (Exception ex)
            {
                Log("Ошибка при обработке существующих файлов: " + ex.Message);
            }
        }

        /// <summary>
        /// Выполнение программы получение чертежей фасадов на ББ
        /// </summary>
        /// <param name="filePath"></param>
        private static void ProcessFileDrawingFronts(string filePath,ISldWorks swApp)
        {

            //Копирование файла с параметрами в целевую папку
            CopyFileWithOverwrite(filePath, Costants.fileConfigBBFronts);
            //Запуск программы создания чертежей на фасады
            swApp.RunMacro(Costants.fileMacroBBFronts, "Фасады", "CreateDrawing");
        }
       
        /// <summary>
        /// Выполнение программы расчета запаса прочности на ББ
        /// </summary>
        /// <param name="filePath"></param>
        private static void ProcessBBStaticLoad(string filePath, ISldWorks swApp)
        {
            //Копироние файла с параметрами в папку в целевую папку
            CopyFileWithOverwrite(filePath, Costants.fileConfigBBStaticLoad);
            //Запуск расчета прочности
            swApp.RunMacro(Costants.fileMacroBBStaticLoad, "НагрузкиББ", "SetSizes");
        }

        private bool WaitForFileReady(string path, TimeSpan timeout)
        {
            return FileReadiness.WaitForReady(path, timeout, TimeSpan.FromSeconds(1), 2, message => Log(message));
        }

        private void RunWorkerWithTimeout(
            string jobType,
            string logMessage,
            int timeoutSeconds,
            string filePath,
            bool killSolidWorks,
            bool killOffice,
            bool killAutoCad = false)
        {
            bool success = false;

            try
            {
                for (int attempt = 1; attempt <= 2 && !success; attempt++)
                {
                    if (killOffice) PrepareOfficeForAutomation();

                    Log($"Попытка {attempt} выполнить задание {jobType}: {filePath}");

                    using (Process worker = StartWorker(jobType, filePath))
                    {
                        if (!worker.WaitForExit(timeoutSeconds * 1000))
                        {
                            Log($"Попытка {attempt} неуспешна по таймауту: {logMessage}");
                            TryKillProcess(worker);
                            CleanupAutomationProcesses(killSolidWorks, killOffice, killAutoCad);
                            continue;
                        }

                        if (worker.ExitCode == 0)
                        {
                            Log($"Попытка {attempt} успешна!");
                            success = true;
                        }
                        else
                        {
                            Log($"Попытка {attempt} завершилась ошибкой. ExitCode={worker.ExitCode}. {logMessage}");
                            CleanupAutomationProcesses(killSolidWorks, killOffice, killAutoCad);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Основное действие завершилось с ошибкой: {ex}");
                CleanupAutomationProcesses(killSolidWorks, killOffice, killAutoCad);
            }
            finally
            {
                CleanupAutomationProcesses(killSolidWorks, killOffice, killAutoCad);
                MoveProcessedFile(filePath, success);
            }
        }

        private Process StartWorker(string jobType, string filePath)
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"{WorkerMode} {QuoteArgument(jobType)} {QuoteArgument(filePath)}",
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            return Process.Start(startInfo);
        }

        private static string QuoteArgument(string value)
        {
            if (value == null) return "\"\"";
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private void TryKillProcess(Process process)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill();
                    process.WaitForExit(5000);
                }
            }
            catch (Exception ex)
            {
                Log($"Не удалось завершить worker-процесс: {ex.Message}");
            }
        }

        private void KillCadProcesses(bool killSolidWorks, bool killOffice, bool killAutoCad)
        {
            if (killSolidWorks) KillSelectedProcesses(solidNames);
            if (killOffice) KillSelectedProcesses(officeNames);
            if (killAutoCad) KillSelectedProcesses(acadNames);
        }

        private void CleanupAutomationProcesses(bool killSolidWorks, bool killOffice, bool killAutoCad)
        {
            if (killOffice) CloseOfficeProcesses();
            KillCadProcesses(killSolidWorks, false, killAutoCad);
        }

        private void PrepareOfficeForAutomation()
        {
            CloseOfficeProcesses();
            ClearWordRecoveryState();
        }

        private void CloseOfficeProcesses()
        {
            TryQuitOfficeApplication("Word.Application");
            TryQuitOfficeApplication("Excel.Application");
            KillSelectedProcesses(officeNames);
            ClearWordRecoveryState();
        }

        private void ClearWordRecoveryState()
        {
            DeleteWordRecoveryRegistryKeys();

            string appData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
            string localAppData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
            string temp = Path.GetTempPath();

            DeleteFilesByPatterns(Path.Combine(appData, @"Microsoft\Word"), "*.asd", "*.tmp", "~*.tmp");
            DeleteFilesByPatterns(Path.Combine(localAppData, @"Microsoft\Office\UnsavedFiles"), "*.*");
            DeleteFilesByPatterns(temp, "~$*.doc*", "~WR*.tmp", "*.asd");
        }

        private void DeleteWordRecoveryRegistryKeys()
        {
            try
            {
                using (RegistryKey officeKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office"))
                {
                    if (officeKey == null) return;

                    foreach (string version in officeKey.GetSubKeyNames())
                    {
                        try
                        {
                            using (RegistryKey resiliencyKey = officeKey.OpenSubKey(version + @"\Word\Resiliency", true))
                            {
                                if (resiliencyKey == null) continue;
                                resiliencyKey.DeleteSubKeyTree("DocumentRecovery", false);
                                Log($"Очищено состояние восстановления Word для Office {version}.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"Не удалось очистить DocumentRecovery Word для Office {version}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Не удалось проверить состояние восстановления Word в реестре: {ex.Message}");
            }
        }

        private void DeleteFilesByPatterns(string directory, params string[] patterns)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) return;

                foreach (string pattern in patterns)
                {
                    foreach (string file in Directory.GetFiles(directory, pattern, SearchOption.TopDirectoryOnly))
                    {
                        try
                        {
                            File.SetAttributes(file, FileAttributes.Normal);
                            File.Delete(file);
                            Log($"Удален файл восстановления Office: {file}");
                        }
                        catch (Exception ex)
                        {
                            Log($"Не удалось удалить файл восстановления Office {file}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Не удалось очистить папку восстановления Office {directory}: {ex.Message}");
            }
        }

        private void TryQuitOfficeApplication(string progId)
        {
            object app = null;

            try
            {
                app = Marshal.GetActiveObject(progId);

                try
                {
                    app.GetType().InvokeMember(
                        "DisplayAlerts",
                        System.Reflection.BindingFlags.SetProperty,
                        null,
                        app,
                        new object[] { 0 });
                }
                catch { }

                app.GetType().InvokeMember(
                    "Quit",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    app,
                    null);

                Log($"Office-приложение закрыто через COM: {progId}");
            }
            catch (COMException)
            {
            }
            catch (Exception ex)
            {
                Log($"Не удалось закрыть Office через COM ({progId}): {ex.Message}");
            }
            finally
            {
                if (app != null)
                {
                    try { Marshal.FinalReleaseComObject(app); } catch { }
                }
            }
        }

        private void MoveProcessedFile(string filePath, bool success)
        {
            try
            {
                string finalPath = ProcessedTaskFile.Move(filePath, success, CreateProcessedTaskPaths());
                Log($"Обработка файла завершена: {filePath}. Итоговый путь: {finalPath}");
            }
            catch (Exception ex)
            {
                Log($"Не удалось переместить файл после обработки {filePath}: {ex}");
            }
        }

        private ProcessedTaskPaths CreateProcessedTaskPaths()
        {
            return new ProcessedTaskPaths
            {
                FinalDirectory = Costants.filePathFinal,
                FinalBadDirectory = Costants.filePathFinalBad,
                SearchFinalDirectory = Costants.filePathFinalSEARCH,
                SearchFinalBadDirectory = Costants.filePathFinalBadSEARCH
            };
        }

        private static int ExecuteWorker(string[] args)
        {
            if (args.Length < 3)
            {
                Log("Worker запущен без обязательных аргументов.");
                return 2;
            }

            string jobType = args[1];
            string filePath = args[2];
            ISldWorks swApp = null;

            try
            {
                switch (jobType)
                {
                    case CadJobType.Antarus:
                    case CadJobType.Search:
                        swApp = SolidWorksManager.swApp;
                        CreateDrawingAndModel(filePath);
                        break;

                    case CadJobType.Bmi:
                        swApp = SolidWorksManager.swApp;
                        CreateDrawingBMI(filePath);
                        break;

                    case CadJobType.StaticLoad:
                        swApp = SolidWorksManager.swApp;
                        ProcessBBStaticLoad(filePath, swApp);
                        break;

                    case CadJobType.Fronts:
                        swApp = SolidWorksManager.swApp;
                        ProcessFileDrawingFronts(filePath, swApp);
                        break;

                    case CadJobType.Vzu:
                        VZUCreateDWG(filePath);
                        break;

                    default:
                        Log("Неизвестный тип worker-задания: " + jobType);
                        return 2;
                }

                return 0;
            }
            catch (Exception ex)
            {
                Log($"Worker завершился ошибкой для {filePath}: {ex}");
                return 1;
            }
            finally
            {
                if (swApp != null)
                {
                    try { swApp.CloseAllDocuments(true); } catch { }
                }
            }
        }

        private static int ExecuteSelfTest()
        {
            string antarus = CadJobClassifier.GetJobType("995299.MLVФ ЗавестиАнтарус.xlsx");
            string staticLoad = CadJobClassifier.GetJobType("4100х2400х2900 РасчитатьНагрузки_ББ (ТЗ-25500).txt");
            string fronts = CadJobClassifier.GetJobType("4100х2400х2900 ЧертежиФасадов_ББ (ТЗ-25500).txt");
            string vzu = CadJobClassifier.GetJobType("750725. ЧертежВЗУ.txt");

            if (antarus != CadJobType.Antarus) return 1;
            if (staticLoad != CadJobType.StaticLoad) return 1;
            if (fronts != CadJobType.Fronts) return 1;
            if (vzu != CadJobType.Vzu) return 1;

            Console.WriteLine("Self-test OK");
            return 0;
        }

        [STAThread]
        public static int Main(string[] args)
        {
            EmbeddedAssemblyResolver.Initialize();

            if (args != null && args.Length > 0 && args[0] == SelfTestMode)
            {
                return ExecuteSelfTest();
            }

            if (args != null && args.Length > 0 && args[0] == WorkerMode)
            {
                return ExecuteWorker(args);
            }

            var app = new AntarusFileWatcher();
            app.Start();
            return 0;
        }
    }
}
