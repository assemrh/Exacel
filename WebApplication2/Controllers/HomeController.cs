using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Web.Mvc;
using WebApplication2.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
 

            return Json("ok", JsonRequestBehavior.AllowGet);
        }

        public ActionResult Index1()
        {
            DropAllTables();
            int count = 0;
            var path = @"C:\websites\letters.json";
            string json = System.IO.File.ReadAllText(path);
            dynamic array = Newtonsoft.Json.JsonConvert.DeserializeObject<List<leterModel>>(json);
            List<string> vs1 = new List<string>();
            foreach(leterModel x in array)
            {
                List<string> cols = new List<string>() {  "parent_id", "letter", "created_at" };
                List<object> vals = new List<object>() {  x.parent_id, x.letter, DateTime.Now };
                Guid guid = new Guid();
                Guid guid1 = new Guid();
                if (Guid.TryParse(x.id, out guid))
                {
                    if(!string.IsNullOrWhiteSpace(x.id) && !string.IsNullOrWhiteSpace(x.letter) &&!string.IsNullOrWhiteSpace(x.parent_id))
                    {
                        string str = $"INSERT INTO letters (id, parent_id, letter, created_at)  VALUES ";
                        str += $" ('{guid}', '{x.parent_id}', N'{ x.letter}', '{DateTime.Now}');\r\n";
                        vs1.Add(str);
                        ExecQuery(str);
                    }

                    //if ( await Database.InsertRow("letters", guid, cols, vals))
                    if(Guid.TryParse(x.parent_id, out guid1))
                        count++;
                    //if (count == 500)
                       // goto Burya;

                }
                

            }
            Burya:
            //str = str.Remove(str.Length - 2, 1);
            //str += " go ";
            string[] vs = vs1.ToArray();
            try
            {
                System.IO.File.WriteAllLines(@"C:\websites\lines.txt", vs);
                Console.WriteLine("Lines written to file successfully.");
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
            //ExecQuery(str);
            return Json(" ( " + count + " ) Row added successfully", JsonRequestBehavior.AllowGet);
        }
         private static void DropAllTables()
        {
            ExecQuery(@"DELETE FROM [dbo].[letters] WHERE 1=1");
        }
        public  static void ExecQuery(String str)
        {
            SqlConnection cn = new SqlConnection(Database.ConnectionString);
            SqlCommand cmd = new SqlCommand(str, cn);
            cn.Open();
            try
            {
               cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                string errMessage = ex.Message;
            }
            cn.Close();
        }

        public class leterModel
        {

            public string id { get; set; }
            public string parent_id { get; set; }
            public string letter { get; set; }

        }


        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult leters()
        {
            ViewBag.Message = "Your contact page.";
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
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\isc\Downloads\Telegram Desktop\letters_.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            List<leterModel> letters = new List<leterModel>();
            string _id = "";
            string _parent_id = "";
            string _letter = "";
            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                //for (cCnt = 1; cCnt <= cl; cCnt++)
                //{
                //    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                //    Console.WriteLine(str);
                //    //MessageBox.Show(str);
                //}
                
                    _id = (string) (range.Cells[rCnt, 1] as Excel.Range).Value2;
                    _parent_id= (string)(range.Cells[rCnt, 2] as Excel.Range).Value2 ;
                    _letter = (string)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                if(!string.IsNullOrWhiteSpace(_letter))
                letters.Add(new leterModel()
                {
                    id = _id,
                    parent_id = _parent_id,
                    letter = _letter
                });

                //Console.WriteLine(_parent_id);
            }
            //xlWorkBook.SaveAs();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            return Json(letters, JsonRequestBehavior.AllowGet);
        }
    }
}