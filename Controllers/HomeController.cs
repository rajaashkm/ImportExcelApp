using ImportExcelApp.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Net.Mail;
using System.Reflection;
using System.Security.Cryptography;
using static System.Formats.Asn1.AsnWriter;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportExcelApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<List<Questions>> Import(IFormFile file)
        {

            var list = new List<Questions>();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowcount = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rowcount; row++)
                    {
                        list.Add(new Questions{
                            ID = Convert.ToInt32(worksheet.Cells[row,1].Value.ToString()),
                            Question = worksheet.Cells[row, 2].Value.ToString(),
                            RequirementID = worksheet.Cells[row, 3].Value.ToString(),
                            QuestionType = Convert.ToBoolean(worksheet.Cells[row, 4].Value.ToString()),
                            Score = Convert.ToDouble(worksheet.Cells[row, 5].Value.ToString()),
                            Required = Convert.ToBoolean(worksheet.Cells[row, 6].Value.ToString()),
                            Explanation = Convert.ToBoolean(worksheet.Cells[row, 7].Value.ToString()),
                            Attachment = Convert.ToBoolean(worksheet.Cells[row, 8].Value.ToString()),
                            RefNumber = worksheet.Cells[row, 9].Value.ToString()
                        });
                    }
                }
            }

                //return list;
            DataTable dt = ToDataTable(list);
            string conString = @"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=SampleDB;Integrated Security=true";
            
            using (SqlConnection con = new SqlConnection(conString))
            {
                con.Open();
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {

                    sqlBulkCopy.DestinationTableName ="Question_tbl";
                    sqlBulkCopy.ColumnMappings.Add("ID", "ID");
                    sqlBulkCopy.ColumnMappings.Add("Question", "Question");
                    sqlBulkCopy.ColumnMappings.Add("RequirementID", "RequirementID");
                    sqlBulkCopy.ColumnMappings.Add("QuestionType", "QuestionType");
                    sqlBulkCopy.ColumnMappings.Add("Score", "Score");
                    sqlBulkCopy.ColumnMappings.Add("Required", "Required");
                    sqlBulkCopy.ColumnMappings.Add("Explanation", "Explanation");
                    sqlBulkCopy.ColumnMappings.Add("Attachment", "Attachment");
                    sqlBulkCopy.ColumnMappings.Add("RefNumber", "RefNumber");
                    //con.Open();
                    sqlBulkCopy.WriteToServer(dt);
                    con.Close();
                }
            }

            return list;
        }

        private DataTable ToDataTable<Questions>(List<Questions> list)
        {
            DataTable dataTable = new DataTable(typeof(Questions).Name);
            PropertyInfo[] Props = typeof(Questions).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                dataTable.Columns.Add(prop.Name);
            }
            foreach (Questions item in list)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}