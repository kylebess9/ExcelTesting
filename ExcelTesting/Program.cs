namespace ExcelTesting
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        /// 
        
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ExcelFileHandler excelHolder = new ExcelFileHandler("C:\\Users\\p1773957\\source\\repos\\ExcelTesting\\ExcelTesting\\data\\Working - 17503615 Annual Planning.xlsm");
            excelHolder.processSpreadSheet(excelHolder.grabHeaders());
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}