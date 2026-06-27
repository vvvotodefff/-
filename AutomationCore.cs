using System;
using System.IO;
using System.Threading;

namespace ВыполнитьЗадачиSolidWorks
{
    public static class CadJobType
    {
        public const string Antarus = "antarus";
        public const string Bmi = "bmi";
        public const string StaticLoad = "static-load";
        public const string Fronts = "fronts";
        public const string Vzu = "vzu";
        public const string Search = "search";
    }

    public static class CadJobClassifier
    {
        public static string GetJobType(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return null;

            if (fullPath.Contains("ЗавестиАнтарус") && fullPath.Contains(".xlsx")) return CadJobType.Antarus;
            if (fullPath.Contains("ЧертежБМИ") && fullPath.Contains(".xlsx")) return CadJobType.Bmi;
            if (fullPath.Contains("РасчитатьНагрузки_ББ") && fullPath.Contains(".txt")) return CadJobType.StaticLoad;
            if (fullPath.Contains("ЧертежиФасадов_ББ") && fullPath.Contains(".txt")) return CadJobType.Fronts;
            if (fullPath.ToLower().Contains("взу") && fullPath.Contains(".txt")) return CadJobType.Vzu;
            if (fullPath.Contains("НаСерч") && fullPath.Contains(".xlsx")) return CadJobType.Search;

            return null;
        }
    }

    public static class FileReadiness
    {
        public static bool WaitForReady(
            string path,
            TimeSpan timeout,
            TimeSpan pollInterval,
            int requiredStableReads,
            Action<string> log = null)
        {
            DateTime until = DateTime.UtcNow + timeout;
            long lastLength = -1;
            DateTime lastWrite = DateTime.MinValue;
            int stableCount = 0;

            while (DateTime.UtcNow < until)
            {
                try
                {
                    var info = new FileInfo(path);
                    if (!info.Exists)
                    {
                        stableCount = 0;
                    }
                    else
                    {
                        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            bool stable = info.Length > 0
                                && info.Length == lastLength
                                && info.LastWriteTimeUtc == lastWrite
                                && stream.Length == info.Length;

                            if (stable)
                            {
                                stableCount++;
                                if (stableCount >= requiredStableReads)
                                {
                                    log?.Invoke("Файл готов к обработке: " + path);
                                    return true;
                                }
                            }
                            else
                            {
                                stableCount = 0;
                            }

                            lastLength = info.Length;
                            lastWrite = info.LastWriteTimeUtc;
                        }
                    }
                }
                catch (IOException)
                {
                    stableCount = 0;
                }
                catch (UnauthorizedAccessException)
                {
                    stableCount = 0;
                }

                Thread.Sleep(pollInterval);
            }

            return false;
        }
    }

    public sealed class ProcessedTaskPaths
    {
        public string FinalDirectory { get; set; }
        public string FinalBadDirectory { get; set; }
        public string SearchFinalDirectory { get; set; }
        public string SearchFinalBadDirectory { get; set; }
    }

    public static class ProcessedTaskFile
    {
        public static string GetFinalPath(string filePath, bool success, ProcessedTaskPaths paths)
        {
            string fpfinal;
            string fpfinalbad;

            if (filePath.Contains("НаСерч"))
            {
                fpfinal = paths.SearchFinalDirectory;
                fpfinalbad = paths.SearchFinalBadDirectory;
            }
            else
            {
                fpfinal = paths.FinalDirectory;
                fpfinalbad = paths.FinalBadDirectory;
            }

            string finalFilePath = Path.Combine(success ? fpfinal : fpfinalbad, Path.GetFileName(filePath));

            if (finalFilePath.Contains("ЧертежиФасадов_ББ") || finalFilePath.Contains("РасчитатьНагрузки_ББ"))
            {
                finalFilePath = AddTimestampPostfixIfExists(finalFilePath);
            }

            return finalFilePath;
        }

        public static string Move(string filePath, bool success, ProcessedTaskPaths paths)
        {
            string finalPath = GetFinalPath(filePath, success, paths);

            if (File.Exists(finalPath))
            {
                File.Delete(finalPath);
            }

            File.Move(filePath, finalPath);
            return finalPath;
        }

        private static string AddTimestampPostfixIfExists(string filePath)
        {
            if (!File.Exists(filePath)) return filePath;

            string directory = Path.GetDirectoryName(filePath) ?? string.Empty;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            string safeTimestamp = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");

            return Path.Combine(directory, $"{fileNameWithoutExtension}+{safeTimestamp}{extension}");
        }
    }
}
