using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXMLDemo;

namespace WebTest.Controllers
{
    public class FileController : Controller
    {
        //
        // GET: /File/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Download()
        {
            //this.Response.OutputStream
            //FileContentResult f=new FileContentResult(
            List<TestClass> list = new List<TestClass>();
            list.Add(new TestClass { Prop1 = "p1", Prop2 = "p2", Field1 = "f1", Field2 = "f2" });
            DataTable dt = ExcelUtility.ConvertObjectsToDataTable(list);
            System.IO.MemoryStream stream = ExcelUtility.ExportExcelStreamFromDataTable(dt);
            FileContentResult fResult = new FileContentResult(stream.ToArray(), "application/x-xlsx");
            fResult.FileDownloadName = "test.xlsx";
            return fResult;
        }
    }

    public class TestClass
    {
        public string Prop1 { get; set; }
        public string Prop2 { get; set; }

        public string Field1;
        public string Field2;
    }
}
