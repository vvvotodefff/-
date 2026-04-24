using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static ЗавестиАнтарус.AutoScaleDrawingDimension;
using static FileHelper;
using System.Runtime.InteropServices;
using System.Threading;
using static VZU.CreateDrawingVZU;
using System.Collections.Generic;

namespace ВыполнитьЗадачиSolidWorks
{
        /// <summary>
        /// Обнаружение новых файлов
        /// </summary>
    public class AntarusFileWatcher
    {
       
        private FileSystemWatcher watcher;

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
        /// <summary>
        /// Действие программы при запуске
        /// </summary>
        public void Start()
        {
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
        }

        /// <summary>
        /// Действия программы при обнаружении файла
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void OnFileCreated(object source, FileSystemEventArgs e)
        {
            if (e.FullPath.Contains("~$") || e.FullPath.Contains(".tmp")) return;

            //#region На время пока выполняем на серч
            //bool ЗавестиАнтарус = false;
            //string dir = Path.GetDirectoryName(e.FullPath);
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

            Log($"Обнаружен новый файл: {e.FullPath}");
            Thread.Sleep(1000);
            //Запуск программы создания 3D и 2D моделей Антарус
            if (e.FullPath.Contains("ЗавестиАнтарус") && e.FullPath.Contains(".xlsx")) // Если заводим на серч то добавляем "&& !ЗавестиАнтарус"
            {
                KillSelectedProcesses(officeNames);
                RunWithTimeout(
                () => SolidWorksManager.swApp,  // Инициализация SolidWorks
                swApp => CreateDrawingAndModel(e.FullPath), // Вызов метода с параметрами
                "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
                80,// Таймаут в секундах
                e.FullPath
                );
            }

            else if (e.FullPath.Contains("ЧертежБМИ") && e.FullPath.Contains(".xlsx"))
            {
                RunWithTimeout(
                () => SolidWorksManager.swApp,  // Инициализация SolidWorks
                swApp => CreateDrawingBMI(e.FullPath), // Вызов метода с параметрами
                "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
                800,// Таймаут в секундах
                e.FullPath
                );
            }

            //Запуск программы расчета нагрузок ББ
            else if (e.FullPath.Contains("РасчитатьНагрузки_ББ") && e.FullPath.Contains(".txt"))
            {
               KillSelectedProcesses(solidNames);
               KillSelectedProcesses(officeNames);
                RunWithTimeout(
               () => SolidWorksManager.swApp,  // Инициализация SolidWorks
               swApp => ProcessBBStaticLoad(e.FullPath, swApp), // Вызов метода с параметрами
               "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
               300,// Таймаут в секундах
                e.FullPath
               );
            }

            //Запуск программы чертежей фасадов
            else if (e.FullPath.Contains("ЧертежиФасадов_ББ") && e.FullPath.Contains(".txt"))
            {
                KillSelectedProcesses(solidNames);
                RunWithTimeout(
                () => SolidWorksManager.swApp,  // Инициализация SolidWorks
                swApp => ProcessFileDrawingFronts(e.FullPath, swApp), // Вызов метода с параметрами
                "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
                120,// Таймаут в секундах
                e.FullPath
                );
            }

            else if (e.FullPath.ToLower().Contains("взу") && e.FullPath.Contains(".txt"))
            {
                RunWithTimeout(
                "Не удалось создать чертеж на: " + e.FullPath,
                200,// Таймаут в секундах
                e.FullPath
                );
            }

            else if (e.FullPath.Contains("НаСерч") && e.FullPath.Contains(".xlsx"))
            {
                //KillSelectedProcesses(solidNames);
                RunWithTimeout(
                () => SolidWorksManager.swApp,  // Инициализация SolidWorks
                swApp => CreateDrawingAndModel(e.FullPath), // Вызов метода с параметрами
                "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
                80,// Таймаут в секундах
                e.FullPath
                );
            }

            else
            {
                //if (ЗавестиАнтарус) return;   // Если заводим на серч то добавляем
                Console.WriteLine("Неизвестный тип файла: " + e.FullPath);
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
                    OnFileCreated(this, new FileSystemEventArgs(WatcherChangeTypes.Created, Costants.OnPerfFilePath, Path.GetFileName(filePath)));
                    
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
        private void ProcessFileDrawingFronts(string filePath,ISldWorks swApp) 
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
        private void ProcessBBStaticLoad(string filePath, ISldWorks swApp)
        {
            //Копироние файла с параметрами в папку в целевую папку
            CopyFileWithOverwrite(filePath, Costants.fileConfigBBStaticLoad);
            //Запуск расчета прочности
            swApp.RunMacro(Costants.fileMacroBBStaticLoad, "НагрузкиББ", "SetSizes");
        }

        /// <summary>
        /// Запустить метод с таймаутом выполнения
        /// </summary>
        /// <param name="initSwApp">Экземпляр SolidWorks</param>
        /// <param name="action">Выполняемая задача</param>
        /// <param name="logMessage">Сообщение в случае неудачи</param>
        /// <param name="timeoutSeconds">Время на выполнение</param>
        /// <param name="e">Путь до найденного файла</param>
        public void RunWithTimeout(Func<ISldWorks> initSwApp, Action<ISldWorks> action, string logMessage, int timeoutSeconds, string e)
        {
            bool success = true;
            bool completed = false;

            int i = 1;
            Exception threadException = null;

            try
            {
                while (completed == false && (i <= 2))
                {

                    Log($"Попытка {i} создать чертеж и модель");

                    ISldWorks swApp = initSwApp();

                    if (swApp == null)
                    {
                        Log("Ошибка инициализации SolidWorks.");
                        success = false;
                        break;
                    }
                    threadException = null;
                    Thread thread = new Thread(() =>
                    {
                        try
                        {
                            action(swApp);
                        }
                        catch (Exception ex)
                        {
                            threadException = ex;
                        }
                    });

                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();

                    if (!thread.Join(TimeSpan.FromSeconds(timeoutSeconds)))
                    {
                        Log($"Попытка {i} неуспешна—{logMessage}");
                        success = false;
                        KillSelectedProcesses(solidNames);
                        KillSelectedProcesses(officeNames);
                    }
                    else
                    {
                        if (threadException != null)
                        {
                            Log($"Попытка {i} завершилась исключением: {threadException.Message}");
                            success = false;
                            KillSelectedProcesses(solidNames);
                            KillSelectedProcesses(officeNames);
                        }
                        else
                        {
                            Log($"Попытка {i} успешна!");
                            swApp.CloseAllDocuments(true);
                            completed = true;
                        }
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                Log($"Основное действие завершилось с ошибкой: {ex.Message}");
                success = false;
                KillSelectedProcesses(solidNames);
                KillSelectedProcesses(officeNames);
            }
            finally
            {
                //Формирование конечного пути файла с параметрами
                string fpfinal;
                string fpfinalbad;

                if (e.Contains("НаСерч"))
                {
                    fpfinal = Costants.filePathFinalSEARCH;
                    fpfinalbad = Costants.filePathFinalBadSEARCH;
                }
                else
                {
                    fpfinal = Costants.filePathFinal;
                    fpfinalbad = Costants.filePathFinalBad;
                }

                string finalFilePath = Path.Combine(success ? fpfinal : fpfinalbad, Path.GetFileName(e));

                if (finalFilePath.Contains("ЧертежиФасадов_ББ") || finalFilePath.Contains("РасчитатьНагрузки_ББ"))
                {
                    finalFilePath = RemoveFileWithAddPostFix(finalFilePath);
                }

                //Удаление файла если он существует
                if (File.Exists(finalFilePath))
                {
                    File.Delete(finalFilePath);
                }

                //Перемещение файла
                File.Move(e, finalFilePath);

                //Записть в лог
                Log($"Обработка файла завершена: {e}");
            }
        }


            /// <summary>
            /// Запустить метод с таймаутом выполнения
            /// </summary>
            /// <param name="logMessage">Сообщение в случае неудачи</param>
            /// <param name="timeoutSeconds">Время на выполнение</param>
            /// <param name="e">Путь до найденного файла</param>
        public void RunWithTimeout(string logMessage, int timeoutSeconds, string e)
        {
            bool success = true;
            bool completed = false;
            int i = 1;
            Exception threadException = null;

            try
            {
                while (completed == false && (i <= 2))
                {
                    Log($"Попытка {i} создать чертеж и модель");

                    threadException = null;
                    Thread thread = new Thread(() =>
                    {
                        try
                        {
                            VZUCreateDWG(e);
                        }
                        catch (Exception ex)
                        {
                            threadException = ex;
                        }
                    });

                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();

                    if (!thread.Join(TimeSpan.FromSeconds(timeoutSeconds)))
                    {
                        Log($"Попытка {i} неуспешна—{logMessage}");
                        success = false;
                    }
                    else
                    {
                        if (threadException != null)
                        {
                            Log($"Попытка {i} завершилась исключением: {threadException.Message}");
                            success = false;
                        }
                        else
                        {
                            Log($"Попытка {i} успешна!");
                            completed = true;
                        }
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                Log($"Основное действие завершилось с ошибкой: {ex.Message}");
                success = false;
            }

            finally
            {
                //Формирование конечного пути файла с параметрами
                string finalFilePath = Path.Combine(success ? Costants.filePathFinal : Costants.filePathFinalBad, Path.GetFileName(e));

                if (finalFilePath.Contains("ЧертежиФасадов_ББ") || finalFilePath.Contains("РасчитатьНагрузки_ББ"))
                {
                    finalFilePath = RemoveFileWithAddPostFix(finalFilePath);
                }

                //Удаление файла если он существует
                if (File.Exists(finalFilePath))
                {
                    File.Delete(finalFilePath);
                }

                //Перемещение файла
                File.Move(e, finalFilePath);

                //Записть в лог
                Log($"Обработка файла завершена: {e}");
            }

        }

        public static void Main()
        {
            var app = new AntarusFileWatcher();
            app.Start();
        }
    }
}
