using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using OfficeOpenXml;
using System.Threading;
using System.Reflection;
using System.ComponentModel;
using System.Globalization;
using System.Collections;
using System.Runtime.InteropServices;
using static FileHelper;

namespace VZU
{
    /// <summary>
    /// Класс создания всей документации на ВЗУ
    /// </summary>
    public class CreateDrawingVZU
    {
        public static AcadApplication acad;

        /// <summary>
        /// Определяет тип полученнго документа и запускает методы соответствующие данному классу
        /// </summary>
        /// <param name="filepath">Путь до файла с параметра</param>
        public static void VZUCreateDWG(string filepath)
        {
            try
            {
                // Подготовка путей и кода
                var code = CreateFinalPath(filepath);
                var name = Path.GetFileName(filepath)?.ToLower() ?? "";
                var baseDir = $@"\\sol.elita\Spec\.CADAutomation\Модели\ВЗУ\{code}";
                string finalPath;

                // Запускаем AutoCAD
                acad = new AcadApplication();
                Thread.Sleep(1000);
                acad.Visible = true;

                // Если это чертеж — обрабатываем множество файлов
                if (name.Contains("чертежвзу"))
                {
                    var targetFiles = ReadWriteParametersVZUDrawing(filepath);
                    for (int i = 0; i < targetFiles.Length; i++)
                    {
                        var suffix = targetFiles.Length == 1
                            ? $"{code}. ЧертежВЗУ.pdf"
                            : $"{code} ({i + 1}). ЧертежВЗУ.pdf";

                        finalPath = Path.Combine(baseDir, suffix);
                        CreateVZU(targetFiles[i], finalPath);
                    }
                    return;
                }


                var map = new Dictionary<string, (string pdfName, string dwgPath)>(StringComparer.OrdinalIgnoreCase)
                {
                    ["фундамент"] = ($"{code}. ЗаданиеНаФундаментВЗУ.pdf", @"\\sol.elita\Spec\.CADAutomation\ВЗУ\ЗаданиеНаФундамент.DWG"),
                    ["спецификация"] = ($"{code}. СпецификацияВЗУ.pdf", @"\\sol.elita\Spec\.CADAutomation\ВЗУ\Спецификация.DWG"),
                    ["принципиалка взу 1,2 кат, вентиляция"] = ($"{code}. Принципиалка взу 1,2 кат, вентиляция.pdf", @"\\sol.elita\Spec\.CADAutomation\ВЗУ\принципиалка взу 1,2 кат, вентиляция.DWG"),
                    ["принципиалка взу 1,2 кат"] = ($"{code}. Принципиалка взу 1,2 кат.pdf", @"\\sol.elita\Spec\.CADAutomation\ВЗУ\принципиалка взу 1,2 кат.DWG"),
                    ["принципиалка взу 3 кат, вентиляция"] = ($"{code}. Принципиалка взу 3 кат, вентиляция.pdf", @"\\sol.elita\Spec\.CADAutomation\ВЗУ\принципиалка взу 3 кат, вентиляция.DWG"),
                    ["принципиалка взу 3 кат"] = ($"{code}. принципиалка взу 3 кат.pdf", @"\\sol.elita\Spec\.CADAutomation\ВЗУ\принципиалка взу 3 кат.DWG")
                };

                // Находим первое совпадение по ключу
                var entry = map.FirstOrDefault(kvp => name.Contains(kvp.Key));
                if (entry.Key != null)
                {
                    finalPath = Path.Combine(baseDir, entry.Value.pdfName);
                    if (File.Exists(finalPath))
                        File.Delete(finalPath);

                    CreatePDFFileVZU(entry.Value.dwgPath, filepath, finalPath);
                    return;
                }

                throw new InvalidOperationException("Неизвестный тип задания ВЗУ: " + filepath);
            }
            finally
            {
                if (acad != null)
                {
                    try { acad.Quit(); }
                    catch (System.Exception ex) { Log("Не удалось закрыть AutoCAD: " + ex.Message); }
                    finally
                    {
                        try { Marshal.FinalReleaseComObject(acad); } catch { }
                        acad = null;
                    }
                }
            }
        }

        /// <summary>
        /// Открывает документ, обновляет данные, отправляет на печать
        /// </summary>
        /// <param name="pathfile">Путь до файла который нужно открыть</param>
        /// <param name="finalpath">Путь до финального файла</param>
        static void CreateVZU(string pathfile, string finalpath)
        {
            AcadDocument doc = null;
            bool closed = false;
            try
            {
                doc = acad.Documents.Open(pathfile);

                doc.SendCommand("_DATALINKUPDATE\n_u\n_k\n");
                Thread.Sleep(2000);
                PlotWait.PlotToPdfAndWait(acad, doc, finalpath);
                PdfCloser.ClosePdfWindowByTitle(finalpath, TimeSpan.FromSeconds(10));
                doc.Close(true);
                closed = true;
            }
            catch (System.Exception ex)
            {
                Log($"Ошибка создания ВЗУ из {pathfile}: {ex}");
                throw;
            }
            finally
            {
                if (doc != null)
                {
                    if (!closed)
                    {
                        try { doc.Close(false); } catch { }
                    }

                    try { Marshal.FinalReleaseComObject(doc); } catch { }
                }
            }

        }

        /// <summary>
        /// Функция читает файл задание, записывает в файл-параметры, возвращает путь до нужного файла
        /// </summary>
        /// <param name="filePath">Путь до файла с параметрами</param>
        /// <returns></returns>
        static string[] ReadWriteParametersVZUDrawing(string filePath)
        {
            string rootpath = @"\\sol.elita\Spec\.CADAutomation\ВЗУ";
            string patternpath = $@"{rootpath}\Шаблон.xlsx";
            string[] targetpath;
            Dictionary<string, string> parameters = FileHelper.ReadParameters(filePath,'=');
            parameters.TryGetValue("НаименованиеПозицииВЗУ", out string nameValue);
            parameters.TryGetValue("Специалист", out string creatorValue);
            parameters.TryGetValue("ТЗ", out string tzValue);
            FileInfo fi = new FileInfo(patternpath);
            ExcelPackage.License.SetNonCommercialOrganization("SANEK");
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.Cells[6, 3].Value = GetSurname(creatorValue);
                worksheet.Cells[1, 7].Value = "ТЗ-№" + tzValue;
                worksheet.Cells[3, 7].Value = nameValue;
                worksheet.Cells[6, 6].Value = DateTime.Now.ToString("dd.MM.yy");
                worksheet.Cells[7, 6].Value = DateTime.Now.ToString("dd.MM.yy");
                FileInfo finalfolder = new FileInfo(@"S:\.CADAutomation\ВЗУ\ТаблицаПараметровВЗУ.xlsx");
                package.SaveAs(finalfolder);
            }
            parameters.TryGetValue("ДиаметрПодключения", out string DN);
            if (DN.Contains("250"))
            {
                targetpath = new string[] {
                                            $@"{rootpath}\{DN}.DWG"
                                          };
            }
            else
            {
                targetpath = new string[] {
                                            $@"{rootpath}\{DN}.DWG"
                                          };
            }

            return targetpath;
        }

        

        /// <summary>
        /// Функция формирующая путь до файла, возвращает код товара
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        static string CreateFinalPath(string filepath)
        {
            string filename = Path.GetFileName(filepath);
            string nameWithoutExt = Path.GetFileNameWithoutExtension(filename);
            int dotIndex = nameWithoutExt.IndexOf('.');
            string code = nameWithoutExt.Substring(0, dotIndex).Trim();
            string finalfolder = $@"\\sol.elita\Spec\.CADAutomation\Модели\ВЗУ\{code}";

            if (!Directory.Exists(finalfolder))
            {
                Directory.CreateDirectory(finalfolder);
            }
            return code;
        }

        /// <summary>
        /// Функция по "нормализации" формата вывода фамилии
        /// </summary>
        /// <param name="fullName">ФИО</param>
        /// <returns>Нормализованное, строковое значение фамилии</returns>
        static string GetSurname(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
                return "";

            //Разбиваем по пробелу и берём первую часть
            var parts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string rawSurname = parts[0]; //Получаем фамилию

            //Переводим в нижний регистр, а затем в первую букву в заглавную
            TextInfo ti = new CultureInfo("ru-RU", false).TextInfo;
            string properSurname = ti.ToTitleCase(rawSurname.ToLower());
            return properSurname; // вернёт фамилию
        }

        /// <summary>
        /// Изменяет базовое значение на чертеже AUTOCAD
        /// </summary>
        /// <param name="baseValue">Базовое значение</param>
        /// <param name="keyValue">Нужное значение</param>
        /// <param name="acaddoc">Документ</param>
        /// <returns>Результат формата булево</returns>
        public static bool ChangeBaseValue(string baseValue, string keyValue, AcadDocument acaddoc)
        {
            bool result = false;

            // Выбираем пространство: ModelSpace или PaperSpace
            IEnumerable entities = acaddoc.ModelSpace.Count > 0
                ? (IEnumerable)acaddoc.ModelSpace
                : (IEnumerable)acaddoc.PaperSpace;

            foreach (object obj in entities)
            {
                if (obj is AcadEntity entity)
                {
                    result |= ProcessEntity(entity, baseValue, keyValue);
                }
            }

            return result;
        }

        /// <summary>
        /// Запускает замену "Ключа" на необходимое значение, в зависимости от типа AcadEntity
        /// </summary>
        /// <param name="entity">AcadEntity</param>
        /// <param name="baseValue">Ключ</param>
        /// <param name="keyValue">Верное значение</param>
        /// <returns></returns>
        static bool ProcessEntity(AcadEntity entity, string baseValue, string keyValue)
        {
            switch (entity)
            {
                case AcadDimension dim:
                    return ProcessDimension(dim, baseValue, keyValue);
                case AcadTable table:
                    return ProcessTable(table, baseValue, keyValue);
                case AcadText text:
                    return ProcessText(text, baseValue, keyValue);
                case AcadMText mtext:
                    return ProcessMText(mtext, baseValue, keyValue);
                case AcadMLeader mleader:
                    return ProcessMLeader(mleader, baseValue, keyValue);
                default:
                    return false;
            }
        }

        /// <summary>
        /// Замена атрибута у экземпляра класса AcadDimension
        /// </summary>
        /// <param name="dim"></param>
        /// <param name="baseValue"></param>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        static bool ProcessDimension(AcadDimension dim, string baseValue, string keyValue)
        {
            string[] prefixes = { "", "D1=", "D2=" };
            foreach (var prefix in prefixes)
            {
                if (prefix + dim.TextOverride == baseValue)
                {
                    dim.TextOverride = prefix + keyValue;
                    return true;
                }
            }

            string text = dim.TextOverride;
            if (!string.IsNullOrEmpty(text) && text.Contains(baseValue))
            {
                dim.TextOverride = text.Replace(baseValue, keyValue);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Замена атрибута у экземпляра класса AcadTable
        /// </summary>
        /// <param name="table"></param>
        /// <param name="baseValue"></param>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        static bool ProcessTable(AcadTable table, string baseValue, string keyValue)
        {
            bool updated = false;
            for (int r = 0; r < table.Rows; r++)
            {
                for (int c = 0; c < table.Columns; c++)
                {
                    string cellText = table.GetText(r, c);
                    if (!string.IsNullOrEmpty(cellText) && cellText.Contains(baseValue))
                    {
                        string celltext = cellText.Replace(baseValue, keyValue);
                        table.SetText(r, c, celltext);
                        updated = true;
                    }
                }
            }
            return updated;
        }

        /// <summary>
        /// Замена атрибута у экземпляра класса AcadText
        /// </summary>
        /// <param name="text"></param>
        /// <param name="baseValue"></param>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        static bool ProcessText(AcadText text, string baseValue, string keyValue)
        {
            string content = text.TextString;
            if (!string.IsNullOrEmpty(content) && content.Contains(baseValue))
            {
                text.TextString = content.Replace(baseValue, keyValue);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Замена атрибута у экземпляра класса AcadMText
        /// </summary>
        /// <param name="mtext"></param>
        /// <param name="baseValue"></param>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        static bool ProcessMText(AcadMText mtext, string baseValue, string keyValue)
        {
            string content = mtext.TextString;
            if (!string.IsNullOrEmpty(content) && content.Contains(baseValue))
            {
                mtext.TextString = content.Replace(baseValue, keyValue);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Замена атрибута у экземпляра класса AcadMLeader
        /// </summary>
        /// <param name="mleader"></param>
        /// <param name="baseValue"></param>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        static bool ProcessMLeader(AcadMLeader mleader, string baseValue, string keyValue)
        {
            if (mleader == null) return false;
            Console.WriteLine($"{mleader.TextString}-{(int)mleader.ContentType}");
            try
            {
                // acMTextContent = 1
                if ((int)mleader.ContentType == 2)
                {
                    string content = mleader.TextString;
                    if (!string.IsNullOrEmpty(content) && content.Contains(baseValue))
                    {
                        mleader.TextString = content.Replace(baseValue, "DN" + keyValue);
                        return true;
                    }
                }
            }
            catch { /* на случай разных версий COM */ }

            return false;
        }

        /// <summary>
        /// Метод читающий файл параметры, изменяющий значения на чертеже и отправляющий на печать
        /// </summary>
        /// <param name="file"></param>
        /// <param name="filepath"></param>
        /// <param name="finalpath"></param>
        static void CreatePDFFileVZU(string file, string filepath, string finalpath)
        {
            var result = ReadParameters(filepath, '=');
            AcadDocument doc = null;
            bool closed = false;

            try
            {
                doc = acad.Documents.Open(file);
                doc.SendCommand("_DATALINKUPDATE\n_u\n_k\n");
                Thread.Sleep(2000);

                foreach (var kvp in result)
                {
                    string baseValue = kvp.Key;
                    string keyValue = kvp.Value;
                    ChangeBaseValue(baseValue, keyValue, doc);
                }

                PlotWait.PlotToPdfAndWait(acad, doc, finalpath);
                PdfCloser.ClosePdfWindowByTitle(finalpath, TimeSpan.FromSeconds(10));

                doc.Close(false);
                closed = true;
            }
            catch (System.Exception ex)
            {
                Log($"Ошибка создания PDF ВЗУ из {file}: {ex}");
                throw;
            }
            finally
            {
                if (doc != null)
                {
                    if (!closed)
                    {
                        try { doc.Close(false); } catch { }
                    }

                    try { Marshal.FinalReleaseComObject(doc); } catch { }
                }
            }
        }
    }

    static class PlotWait
    {
        static ManualResetEventSlim done = new ManualResetEventSlim(false);
        static _DAcadApplicationEvents_BeginPlotEventHandler onBegin;
        static _DAcadApplicationEvents_EndPlotEventHandler onEnd;

        /// <summary>
        /// Метод отправляющий файл на печать в PDF и ожидание когда он создастся
        /// </summary>
        /// <param name="acad"></param>
        /// <param name="doc"></param>
        /// <param name="pdfPath"></param>
        /// <exception cref="TimeoutException"></exception>
        public static void PlotToPdfAndWait(AcadApplication acad, AcadDocument doc, string pdfPath)
        {
            try { doc.SetVariable("BACKGROUNDPLOT", 0); } catch { }

            onBegin = new _DAcadApplicationEvents_BeginPlotEventHandler(BeginPlot);
            onEnd = new _DAcadApplicationEvents_EndPlotEventHandler(EndPlot);
            acad.BeginPlot += onBegin;
            acad.EndPlot += onEnd;
            done.Reset();

            try
            {
                var plot = doc.Plot;
                plot.PlotToFile(pdfPath, "DWG To PDF.pc3");

                if (!done.Wait(TimeSpan.FromMinutes(5)))
                    throw new TimeoutException("EndPlot не пришёл — драйвер завис?");

                WaitFileStable(pdfPath, TimeSpan.FromMinutes(1));
            }
            finally
            {
                try { acad.BeginPlot -= onBegin; } catch { }
                try { acad.EndPlot -= onEnd; } catch { }
            }
        }

        static void BeginPlot(string drawingName) => done.Reset();
        static void EndPlot(string drawingName) => done.Set();

        /// <summary>
        /// Ожидание стабилизации сохраняемого файла
        /// </summary>
        /// <param name="path"></param>
        /// <param name="timeout"></param>
        static void WaitFileStable(string path, TimeSpan timeout)
        {
            var until = DateTime.UtcNow + timeout;
            long last = -1; int stable = 0;

            while (DateTime.UtcNow < until)
            {
                if (File.Exists(path))
                {
                    long sz = 0;
                    try { sz = new FileInfo(path).Length; } catch { }
                    if (sz > 0 && sz == last && CanOpen(path))
                    {
                        if (++stable >= 2) return; 
                    }
                    else stable = 0;
                    last = sz;
                }
                Thread.Sleep(1000);
            }
        }

        static bool CanOpen(string path)
        {
            try { using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) return fs.Length > 0; }
            catch { return false; }
        }
    }


        static class PdfCloser
    {
        // WinAPI
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;

        /// <summary>
        /// Находит окно, в заголовке которого встречается имя PDF, и шлёт WM_CLOSE.
        /// Ждёт исчезновения окна до timeout.
        /// </summary>
        public static bool ClosePdfWindowByTitle(string pdfPath, TimeSpan timeout)
        {
            if (string.IsNullOrEmpty(pdfPath)) return false;
            string needle = System.IO.Path.GetFileName(pdfPath);
            if (string.IsNullOrEmpty(needle)) return false;
            string needleLower = needle.ToLowerInvariant();

            DateTime until = DateTime.UtcNow + timeout;
            IntPtr target = IntPtr.Zero;

            // Ищем окно несколько раз
            while (DateTime.UtcNow < until)
            {
                target = FindWindowByTitleContains(needleLower);
                if (target != IntPtr.Zero)
                {
                    // Закрываем
                    PostMessage(target, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);

                    // ждём, что окно исчезнет
                    if (WaitWindowGone(target, TimeSpan.FromSeconds(5)))
                        return true;
                    // если не исчезло — попробуем найти заново (если хэндл сменился)
                }
                Thread.Sleep(200);
            }
            return false;
        }

        private static IntPtr FindWindowByTitleContains(string needleLower)
        {
            IntPtr found = IntPtr.Zero;

            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;

                int len = GetWindowTextLength(hWnd);
                if (len <= 0) return true;

                var sb = new StringBuilder(len + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                string title = sb.ToString();

                if (!string.IsNullOrEmpty(title) && title.ToLowerInvariant().Contains(needleLower))
                {
                    found = hWnd;
                    return false; // стоп
                }
                return true; // продолжить
            }, IntPtr.Zero);

            return found;
        }

        private static bool WaitWindowGone(IntPtr hWnd, TimeSpan timeout)
        {
            DateTime until = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < until)
            {
                int len = GetWindowTextLength(hWnd);
                if (len <= 0) return true; // окно пропало/неактивно
                Thread.Sleep(200);
            }
            return false;
        }
    }

}
