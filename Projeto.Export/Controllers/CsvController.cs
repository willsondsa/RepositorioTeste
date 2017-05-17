using Excel;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Projeto.Export.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Projeto.Export.Controllers
{
    public class CsvController : Controller
    {
        // GET: Csv
        public FileResult Convert(HttpPostedFileBase file)
        {
            string path = "";
            string _path = "";
            path = Server.MapPath("~/UploadedFiles/");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);

            }
            if (file.ContentLength > 0)
            {
                //IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(file.InputStream);
                string _FileName = Path.GetFileName(file.FileName);
                 _path = Path.Combine(Server.MapPath("~/UploadedFiles"), _FileName);
                path= Path.Combine(Server.MapPath("~/UploadedFiles"),"teste.csv");
                readXSL(file.InputStream,path);
  
            }
            return File(path, "text/csv", "filename.csv"); ;

            // return File(path, "text/csv", "filename.csv"); ;
        }
        public ActionResult Index()
        {

            return View();
        }
       

        private static string ReadFileHtmlToString(Stream file)
        {
            string html = String.Empty;
            StreamReader rd = new StreamReader(file, true);
            while (!rd.EndOfStream)
            {
                string linha = rd.ReadLine();
                html = String.Concat(html, linha);
            }
            rd.Close();
            return html;
        }

        private void readXSL(Stream st,string path)
        {

            HSSFWorkbook workbook = new HSSFWorkbook(st);
            // recupera a Sheet de nome Plan1
            ISheet sheet = workbook.GetSheetAt(0);

            // recupera as linhas da Sheet
            IEnumerator rows = sheet.GetRowEnumerator();
            using (StreamWriter str = new StreamWriter(path))
            { 
                while (rows.MoveNext())
                {
                    IRow row = (HSSFRow)rows.Current;

                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        ICell cell = row.GetCell(i);
                        string ji = cell.ToString();
                        str.Write(ji+";");
                        
                    }

                }
            }

        }
        private void ConvertToXlsIntoCsv(string filePath)
        {
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ filePath + ";Extended Properties='Excel 8.0;HDR=Yes;'";
             using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string jh=dr.GetValue(0).ToString();
                        var row1Col0 = dr[0];
                       // Console.WriteLine(row1Col0);
                    }
                }
            }
        }




        public FileResult Down(string name)
        {
            return File(name, "text/csv", "filename.csv");
        }
    }

}
