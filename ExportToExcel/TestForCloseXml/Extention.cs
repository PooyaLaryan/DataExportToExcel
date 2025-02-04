using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Reflection;
using static TestForCloseXml.MyAttribute;

namespace TestForCloseXml
{
    public static class Extention 
    {
        public static DataTable GetDataTable<T>(this IEnumerable<T> data) where T : class
        {
            if (data == null)
            {
                throw new ArgumentNullException("list");
            }

            IEnumerable<PropertyInfo> propInfo = typeof(T).GetProperties().Where(prop => Attribute.IsDefined(prop, typeof(ExcelColumnAttribute)));

            List<string> columns = new List<string>();
            DataTable dataTable = new DataTable();
            foreach (PropertyInfo propertyInfo in propInfo)
            {
                var customAttr = propertyInfo.GetCustomAttribute(typeof(ExcelColumnAttribute), true);
                if (customAttr != null)
                {
                    dataTable.Columns.Add(propertyInfo.Name, propertyInfo.PropertyType);
                }
            }

            foreach (T item in data)
            {
                DataRow row = dataTable.NewRow();
                foreach (PropertyInfo propertyInfo in propInfo)
                {
                    row[propertyInfo.Name] = propertyInfo.GetValue(item);
                }
                dataTable.Rows.Add(row);
            }

            return dataTable;
        }
    }
}
