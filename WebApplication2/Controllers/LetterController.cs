using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication2.Controllers
{
    public class LetterController : Controller
    {
        //  : Add Letters to database 
        [HttpPost]
        //[Route("AddLetters")]
        public ActionResult Letter()
        {
            DropAllTables();
            
            int count = 0;
            var path = @"C:\websites\letters.json";
            string json = System.IO.File.ReadAllText(path);
            dynamic array = Newtonsoft.Json.JsonConvert.DeserializeObject<List<leterModel>>(json);
            List<string> vs1 = new List<string>();
            foreach (leterModel x in array)
            {
                if (!string.IsNullOrWhiteSpace(x.id) && !string.IsNullOrWhiteSpace(x.letter) && !string.IsNullOrWhiteSpace(x.parent_id))
                {
                    string str = $"INSERT INTO letters (id, parent_id, letter, created_at)  VALUES ";
                    str += $" ('{x.id}', '{x.parent_id}', N'{ x.letter}', '{DateTime.Now}');\r\n";
                    vs1.Add(str);
                    ExecQuery(str);
                    count++;
                }

            }

            return Json(" ( " + count + " ) Row added successfully", JsonRequestBehavior.AllowGet);
        }


        private static void DropAllTables()
        {
            ExecQuery(@"DELETE FROM [dbo].[letters] WHERE 1=1");
        }
        public static void ExecQuery(String str)
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
    }
}