using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


    public static class Costants
    {

        /// <summary>
        /// Папка где находятся модели SolidWorks установок 
        /// </summary>
        public const string RootDirectorySW = @"\\sol.elita\Spec\.CADAutomation\SW\Модели";


        /// <summary>
        /// Папка где находятся чертежи и 3D модели установок 
        /// </summary>
        public const string RootDirectory = @"\\sol.elita\Spec\.CADAutomation\Модели";


        /// <summary>
        /// Папка где хранятся модели серч
        /// </summary>
        public const string RootDirectorySEARCH = @"\\sol.elita\Spec\.CADAutomation\МоделиСерч";

        /// <summary>
        /// Файл ведения Лога
        /// </summary>
        public const string logFile = @"\\sol.elita\Spec\.CADAutomation\service_log.txt";


        /// <summary>
        /// Папка для отслеживания появления новых файлов   
        /// </summary>
        public const string OnPerfFilePath = @"\\sol.elita\Spec\.CADAutomation\ЗадачиНаВыполнение";

        /// <summary>
        /// Папка для успешно выполненных заданий
        /// </summary>
        public const string filePathFinal = @"\\sol.elita\Spec\.CADAutomation\ЗадачиВыполненые";

        public const string filePathFinalSEARCH = @"\\sol.elita\Spec\.CADAutomation\ЗадачиВыполненые";

        /// <summary>
        /// Папка для не выполненных заданий
        /// </summary>
        public const string filePathFinalBad = @"\\sol.elita\Spec\.CADAutomation\ЗадачиНеВыполненые";

        public const string filePathFinalBadSEARCH = @"\\sol.elita\Spec\.CADAutomation\СерчНеВыполнено";


    /// <summary>
    /// Файл конфигурации ЧОВ/Фасады Блок-бокса
    /// </summary>
    public const string fileConfigBBFronts= @"\\sol.elita\Spec\.CADAutomation\SW\Блок-боксы\ЧертежиФасадов_ББ.txt";

        /// <summary>
        /// Файл конфигурации расчета прочности Блок-бокса
        /// </summary>
        public const string fileConfigBBStaticLoad = @"\\sol.elita\Spec\.CADAutomation\SW\Блок-боксы\РасчитатьНагрузки_ББ.txt";

        /// <summary>
        /// Файл макроса получения ЧОВ/Фасадов Блок-Бокса
        /// </summary>
        public const string fileMacroBBFronts = @"\\sol.elita\Spec\.CADAutomation\SW\Блок-боксы\Макросы для 1С\Фасады для Блок-Бокса.swp";

        /// <summary>
        /// Файл макроса получения отчета расчета прочности
        /// </summary>
        public const string fileMacroBBStaticLoad = @"\\sol.elita\Spec\.CADAutomation\SW\Блок-боксы\Макросы для 1С\Расчет прочности Блок-Бокса.swp";

        

    }
