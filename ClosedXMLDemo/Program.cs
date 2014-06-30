using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXMLDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            List<TestClass> list = new List<TestClass>();
            list.Add(new TestClass { Prop1 = "p1", Prop2 = "p2", Field1 = "f1", Field2 = "f2" });
            DataTable dt = ExcelUtility.ConvertObjectsToDataTable(list);
            ExcelUtility.ExportExcelFromDataTable(dt, "test.xlsx");
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
