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

        }


        List<projectObject> projectObjects;
        Excel.Application xlApp;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range range;
        Dictionary<string, int> headerMap;
        public ExcelFileHandler(string path)
        {
            xlApp = new Excel.Application();
            wb = xlApp.Workbooks.Open(path);
            ws = wb.Worksheets[1];
            range = ws.UsedRange;
            projectObjects = new List<projectObject>();
            headerMap = grabHeaders();

        }

        public void closeFile()
        {
            xlApp.Workbooks.Close();
        }

        public void printWorkSheet()
        {
            String test = range.Cells[2,1].Value2;
            Console.WriteLine(range.ToString());
        }

        public Dictionary<string, int> grabHeaders()
        {
            Dictionary<string, int> returnApplicableHeaders = new Dictionary<string, int>();
            string header = "";
            int iterator = 1;
            do
            {
                if (range.Cells[1, iterator].Value != null)
                {
                    header = range.Cells[1, iterator].Value.ToString();
                    header = header.ToLower();
                }
                else
                    break;
                

                if(header.Contains("start") || header.Contains("begin"))
                {
                    returnApplicableHeaders.Add("start", iterator);
                }
                else if(header.Contains("end") || header.Contains("finish"))
                {
                    returnApplicableHeaders.Add("end", iterator);
                }

                iterator++;
            }
            while (header != string.Empty);
            return returnApplicableHeaders;
        }

        public void processSpreadSheet(Dictionary<string, int> headers)
        {
            bool flag = true;
            int iterator = 2;
            while(flag)
            {
                projectObject addObj = new projectObject();
                foreach (string key in headers.Keys)
                {
                    
                    if (range.Cells[iterator, headers[key]].Value == null)
                    {
                        flag = false;
                        break;
                    }
                    string value = range.Cells[iterator, headers[key]].Value.ToString();
                    
                    if(key == "start")
                    {
                        addObj.projectStart = value;
                    }
                    else if(key == "end")
                    {
                        addObj.projectFinish = value;
                    }
                    if (flag)
                    {
                        projectObjects.Add(addObj);
                    }
                }
                iterator++;
            }
        }

    }
}
