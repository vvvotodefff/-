using System;
using System.IO;
using ВыполнитьЗадачиSolidWorks;

namespace AutomationCore.Tests
{
    internal static class Program
    {
        private static int failures;

        private static int Main()
        {
            try
            {
                string repoRoot = FindRepoRoot();
                string autotests = Path.Combine(repoRoot, "Автотесты");

                ClassifiesAutotestFiles(autotests);
                WaitsForReadyFile();
                MovesProcessedFiles();

                if (failures == 0)
                {
                    Console.WriteLine("Все автотесты пройдены.");
                    return 0;
                }

                Console.WriteLine("Автотесты завершились с ошибками: " + failures);
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Критическая ошибка автотестов:");
                Console.WriteLine(ex);
                return 2;
            }
        }

        private static void ClassifiesAutotestFiles(string autotests)
        {
            AssertEqual(CadJobType.Antarus,
                CadJobClassifier.GetJobType(Path.Combine(autotests, "995299.MLVФ ЗавестиАнтарус.xlsx")),
                "ЗавестиАнтарус классифицируется как antarus");

            AssertEqual(CadJobType.StaticLoad,
                CadJobClassifier.GetJobType(Path.Combine(autotests, "4100х2400х2900 РасчитатьНагрузки_ББ (ТЗ-25500).txt")),
                "РасчитатьНагрузки_ББ классифицируется как static-load");

            AssertEqual(CadJobType.Fronts,
                CadJobClassifier.GetJobType(Path.Combine(autotests, "4100х2400х2900 ЧертежиФасадов_ББ (ТЗ-25500).txt")),
                "ЧертежиФасадов_ББ классифицируется как fronts");

            AssertEqual(CadJobType.Vzu,
                CadJobClassifier.GetJobType(Path.Combine(autotests, "750725. ЧертежВЗУ.txt")),
                "ЧертежВЗУ классифицируется как vzu");

            AssertEqual(CadJobType.Vzu,
                CadJobClassifier.GetJobType(Path.Combine(autotests, "993759. принципиалка взу 3 кат.txt")),
                "принципиалка ВЗУ классифицируется как vzu");

            AssertEqual(CadJobType.Bmi,
                CadJobClassifier.GetJobType("123456. ЧертежБМИ.xlsx"),
                "ЧертежБМИ классифицируется как bmi");

            AssertEqual(CadJobType.Search,
                CadJobClassifier.GetJobType("123456. MLV НаСерч.xlsx"),
                "НаСерч классифицируется как search");

            AssertEqual(null,
                CadJobClassifier.GetJobType("readme.txt"),
                "Неизвестный файл не классифицируется");
        }

        private static void WaitsForReadyFile()
        {
            string dir = CreateTempDirectory();
            string path = Path.Combine(dir, "ready.txt");
            File.WriteAllText(path, "ready");

            bool ready = FileReadiness.WaitForReady(
                path,
                TimeSpan.FromSeconds(2),
                TimeSpan.FromMilliseconds(10),
                requiredStableReads: 1);

            AssertTrue(ready, "Готовый файл определяется как доступный");
        }

        private static void MovesProcessedFiles()
        {
            string dir = CreateTempDirectory();
            string finalDir = Directory.CreateDirectory(Path.Combine(dir, "done")).FullName;
            string badDir = Directory.CreateDirectory(Path.Combine(dir, "bad")).FullName;
            string searchDone = Directory.CreateDirectory(Path.Combine(dir, "search-done")).FullName;
            string searchBad = Directory.CreateDirectory(Path.Combine(dir, "search-bad")).FullName;

            var paths = new ProcessedTaskPaths
            {
                FinalDirectory = finalDir,
                FinalBadDirectory = badDir,
                SearchFinalDirectory = searchDone,
                SearchFinalBadDirectory = searchBad
            };

            string source = Path.Combine(dir, "995299.MLVФ ЗавестиАнтарус.xlsx");
            File.WriteAllText(source, "task");

            string moved = ProcessedTaskFile.Move(source, success: true, paths);
            AssertTrue(File.Exists(moved), "Успешное задание перенесено");
            AssertEqual(finalDir, Path.GetDirectoryName(moved), "Успешное задание попало в done");

            string searchSource = Path.Combine(dir, "123456. MLV НаСерч.xlsx");
            File.WriteAllText(searchSource, "task");

            string searchMoved = ProcessedTaskFile.Move(searchSource, success: false, paths);
            AssertTrue(File.Exists(searchMoved), "Неуспешное search-задание перенесено");
            AssertEqual(searchBad, Path.GetDirectoryName(searchMoved), "Неуспешное search-задание попало в search-bad");

            string frontsSource = Path.Combine(dir, "4100х2400х2900 ЧертежиФасадов_ББ (ТЗ-25500).txt");
            File.WriteAllText(frontsSource, "task");
            string existingFinal = Path.Combine(finalDir, Path.GetFileName(frontsSource));
            File.WriteAllText(existingFinal, "existing");

            string frontsMoved = ProcessedTaskFile.Move(frontsSource, success: true, paths);
            AssertTrue(File.Exists(existingFinal), "Существующий файл фасадов не перезаписан");
            AssertTrue(File.Exists(frontsMoved), "Новый файл фасадов перенесен с постфиксом");
            AssertTrue(Path.GetFileNameWithoutExtension(frontsMoved).Contains("+"), "При конфликте фасадов добавлен постфикс");
        }

        private static string FindRepoRoot()
        {
            string dir = AppContext.BaseDirectory;

            while (!string.IsNullOrEmpty(dir))
            {
                if (Directory.Exists(Path.Combine(dir, "Автотесты")))
                {
                    return dir;
                }

                dir = Directory.GetParent(dir)?.FullName;
            }

            throw new DirectoryNotFoundException("Не найдена папка Автотесты.");
        }

        private static string CreateTempDirectory()
        {
            string path = Path.Combine(Path.GetTempPath(), "ZavestiAntarusTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private static void AssertTrue(bool condition, string name)
        {
            if (condition)
            {
                Console.WriteLine("[OK] " + name);
                return;
            }

            failures++;
            Console.WriteLine("[FAIL] " + name);
        }

        private static void AssertEqual(string expected, string actual, string name)
        {
            if (string.Equals(expected, actual, StringComparison.Ordinal))
            {
                Console.WriteLine("[OK] " + name);
                return;
            }

            failures++;
            Console.WriteLine("[FAIL] " + name);
            Console.WriteLine("  expected: " + (expected ?? "<null>"));
            Console.WriteLine("  actual:   " + (actual ?? "<null>"));
        }
    }
}
