using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Filter_xls
{
    class Read_From_Excel
    {
        public static List<Student> getStudentList(String path) 
        {
            List<Student> students = new List<Student>();
             
            Excel.Application xlApp = new Excel.Application();            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;            int colCount = xlRange.Columns.Count;
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= rowCount; i++)
            {                string name=null;                string surename = null;                int birthday = -1, knownlang = -1;
                string phonenumber =null;                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    name =xlRange.Cells[i, 1].Value2.ToString();                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                    surename = xlRange.Cells[i, 2].Value2.ToString();
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
                    Int32.TryParse(xlRange.Cells[i, 3].Value2.ToString(),out birthday);
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
                    Int32.TryParse(xlRange.Cells[i, 4].Value2.ToString(), out knownlang);
                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                    phonenumber=xlRange.Cells[i, 5].Value2.ToString();

                Student student = new Student(name, surename, birthday, knownlang, phonenumber);
                students.Add(student);
          }

            //cleanup            GC.Collect();            GC.WaitForPendingFinalizers();            
            //rule of thumb for releasing com objects:            //  never use two dots, all COM objects must be referenced and released individually            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);            return students;
        }

        public static FilterValues getFilterList(string path)
        {
            FilterValues filters = new FilterValues();

            Excel.Application xlApp = new Excel.Application();            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            if (xlRange.Cells[1, 1] != null && xlRange.Cells[1, 1].Value2 != null)
                filters.NameCompare = xlRange.Cells[1, 1].Value2.ToString();
            if (xlRange.Cells[2, 1] != null && xlRange.Cells[2, 1].Value2 != null)
                filters.SurnameCompare = xlRange.Cells[2, 1].Value2.ToString();
            if (xlRange.Cells[3, 1] != null && xlRange.Cells[3, 1].Value2 != null)
                filters.BirhtDataFrom = Int32.Parse(xlRange.Cells[3, 1].Value2.ToString());
            if (xlRange.Cells[3, 2] != null && xlRange.Cells[3, 2].Value2 != null)
                filters.BirthDataTo = Int32.Parse(xlRange.Cells[3, 2].Value2.ToString());
            if (xlRange.Cells[4, 1] != null && xlRange.Cells[4, 1].Value2 != null)
                filters.KnownLangFrom = Int32.Parse(xlRange.Cells[4, 1].Value2.ToString());
            if (xlRange.Cells[4, 2] != null && xlRange.Cells[4, 2].Value2 != null)
                filters.KnownLangTo = Int32.Parse(xlRange.Cells[4, 2].Value2.ToString());
            if (xlRange.Cells[5, 1] != null && xlRange.Cells[5, 1].Value2 != null)
                filters.PhoneNumberCompare = xlRange.Cells[5, 1].Value2.ToString();

            //cleanup
            GC.Collect();            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad



            //release com objects to fully kill excel process from running in the background

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return filters;

        }
    }
}
