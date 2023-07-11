using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTesting
{
    public class ExcelFileHandler
    {
        internal class projectObject
        {
            public string portfolio, costCenterNum, Group, Clarity, projectStart, projectFinish, ITDM, projectName, projectDescription;

            public projectObject(string portfolio, string costCenterNum, string Group, string Clarity, string projectStart,  string projectFinish, string ITDM, string projectName, string projectDescription)
            {
                this.portfolio = portfolio;
                this.costCenterNum = costCenterNum;
                this.Group = Group;
                this.Clarity = Clarity;
                this.projectStart = projectStart;
                this.projectFinish = projectFinish;
                this.ITDM = ITDM;
                this.projectName = projectName;
                this.projectDescription = projectDescription;
            }
        }


        List<projectObject> projectObjects;
        Excel.Application xlApp;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range range;

        public ExcelFileHandler()
        {
            xlApp = new Excel.Application();
            
        }

        public void LoadFile()
        {
            wb = xlApp.Workbooks.Open("C:\\Users\\p1773957\\source\\repos\\ExcelTesting\\ExcelTesting\\data\\Working - 17503615 Annual Planning.xlsm");
            ws = wb.Worksheets[1];
            range = ws.UsedRange;
            projectObjects = new List<projectObject>();
        }

        public void printWorkSheet()
        {
            String test = range.Cells[2,1].Value2;
            Console.WriteLine(range.ToString());
        }
        public void pullData()
        {
            int i = 2;
            while (range.Cells[i, 1].Value2 != string.Empty)
            {
                string port = range.Cells[i, 1].Value2;
                string cost = range.Cells[i, 2].Value2.ToString();
                string group = range.Cells[i, 3].Value2;
                string clarity = range.Cells[i, 4].Value2;
                string start = range.Cells[i, 5].Value.ToString();
                string end = range.Cells[i, 6].Value2.ToString();
                string itdm = range.Cells[i, 7].Value2;
                string name = range.Cells[i, 8].Value2;
                string desc = range.Cells[i, 9].Value2;

                projectObjects.Add(new projectObject(port, cost, group, clarity, start, end, itdm, name, desc));
                i++;
            }
            Console.WriteLine(projectObjects.Count);
        }

    }
}
