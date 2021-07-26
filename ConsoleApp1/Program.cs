using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            excel();
            Console.ReadKey();
        }
        void InitializeOledbConnection(string filename, string extrn)
        {
            string connString = "";

            if (extrn == ".xls")
                //Connectionstring for excel v8.0    

                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
            else
                //Connectionstring fo excel v12.0    
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";

            //OledbConn = new OLEDBConnection(connString);
        }

        static void excel()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\isc\Downloads\Telegram Desktop\letters_.xlsx", 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            string parent_id ="";
            bool is_parent_id = false;
            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                //for (cCnt = 1; cCnt <= cl; cCnt++)
                //{
                //    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                //    Console.WriteLine(str);
                //    //MessageBox.Show(str);
                //}
                var x = (string)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                is_parent_id = (bool)((range.Cells[rCnt, 4] as Excel.Range).Value2 ?? false);
                if (is_parent_id)
                {
                    parent_id = Guid.NewGuid().ToString();
                    (range.Cells[rCnt, 1] as Excel.Range).Value2 = parent_id;
                    (range.Cells[rCnt, 2] as Excel.Range).Value2 = "";
                }
                else
                {
                    (range.Cells[rCnt,1] as Excel.Range).Value2 = Guid.NewGuid().ToString();
                    (range.Cells[rCnt, 2] as Excel.Range).Value2 = parent_id;

                }

                //Console.WriteLine(rCnt);
                Console.WriteLine(rCnt+" \t=> \t"+parent_id);
            }
            xlWorkBook.SaveAs();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

        }

        static void addToDatabase()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\isc\Downloads\Telegram Desktop\letters_.xlsx", 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            List<letterModel> letters = new List<letterModel>();
            string _id = "";
            string _parent_id = "";
            string _letter = "";
            bool is_parent_id = false;
            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                //for (cCnt = 1; cCnt <= cl; cCnt++)
                //{
                //    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                //    Console.WriteLine(str);
                //    //MessageBox.Show(str);
                //}
                var x = (string)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                is_parent_id = (bool)((range.Cells[rCnt, 4] as Excel.Range).Value2 ?? false);
                if (is_parent_id)
                {
                    _parent_id = Guid.NewGuid().ToString();
                    (range.Cells[rCnt, 1] as Excel.Range).Value2 = _parent_id;
                    (range.Cells[rCnt, 2] as Excel.Range).Value2 = "";
                }
                else
                {
                    (range.Cells[rCnt, 1] as Excel.Range).Value2 = Guid.NewGuid().ToString();
                    (range.Cells[rCnt, 2] as Excel.Range).Value2 = _parent_id;

                }
                letters.Add(new letterModel()
                {
                    id = _id,
                    parent_id = _parent_id,
                    letter = _letter
                }); 

                Console.WriteLine(_parent_id);
            }
            xlWorkBook.SaveAs();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

        }
        static void excel2()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\isc\Downloads\Telegram Desktop\test.xlsx", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, true, 0, false, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;


            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                //for (cCnt = 1; cCnt <= cl; cCnt++)
                //{
                //    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                //    Console.WriteLine(str);
                //    //MessageBox.Show(str);
                //}
                (range.Cells[rCnt, 1] as Excel.Range).Value2 = Guid.NewGuid().ToString("N");

                str = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                str = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                Console.WriteLine(str);
            }

            xlWorkBook.SaveAs(); 
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);

        }
    }
}
