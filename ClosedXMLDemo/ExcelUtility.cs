using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ClosedXMLDemo
{
    public class ExcelUtility
    {
        public static DataTable ConvertObjectsToDataTable(IEnumerable<object> objects)
        {
            DataTable dt = null;

            if (objects != null && objects.Count() > 0)
            {
                Type type = objects.First().GetType();
                dt = new DataTable(type.Name);

                foreach (PropertyInfo property in type.GetProperties())
                {
                    dt.Columns.Add(new DataColumn(property.Name));
                }

                foreach (FieldInfo field in type.GetFields())
                {
                    dt.Columns.Add(new DataColumn(field.Name));
                }

                foreach (object obj in objects)
                {
                    DataRow dr = dt.NewRow();
                    foreach (DataColumn column in dt.Columns)
                    {
                        PropertyInfo propertyInfo = type.GetProperty(column.ColumnName);
                        if (propertyInfo != null)
                        {
                            dr[column.ColumnName] = propertyInfo.GetValue(obj, null);
                        }

                        FieldInfo fieldInfo = type.GetField(column.ColumnName);
                        if (fieldInfo != null)
                        {
                            dr[column.ColumnName] = fieldInfo.GetValue(obj);
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        public static void ExportExcelFromDataTable(DataTable dt, string fileName)
        {
            XLWorkbook workbook = new XLWorkbook();
            workbook.AddWorksheet(dt, "Sheet1");
            workbook.SaveAs(fileName);
        }
    }
}
