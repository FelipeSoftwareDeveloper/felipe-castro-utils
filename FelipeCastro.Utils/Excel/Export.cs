using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FelipeCastro.Utils.Excel
{
    public static class Export
    {
        public static MemoryStream ToExcel<T>(this IEnumerable<T> data)
        {
            var type = typeof(T);
            var name = type.Name;
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(name);

            var properties = type.GetProperties()
            .Where(p => p.IsDefined(typeof(CollumnAttribute), false))
            .Select(p => new
            {
                Property = p,
                Attribute = p.GetCustomAttribute<CollumnAttribute>()
            })
            .ToList();

            worksheet.GenareteHeader(properties.Select(x => x.Property));
            worksheet.GenareteData(data, properties.Select(x => x.Property));

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            return stream;
        }

        public static IXLWorksheet GenareteHeader(this IXLWorksheet worksheet, IEnumerable<PropertyInfo> properties)
        {
            var columnIndex = 1;
            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<CollumnAttribute>();
                if (attribute != null)
                {
                    worksheet.Cell(1, columnIndex).Value = attribute.Name;
                    columnIndex++;
                }
            }

            return worksheet;
        }

        public static IXLWorksheet GenareteData<T>(this IXLWorksheet worksheet, IEnumerable<T> data, IEnumerable<PropertyInfo> properties)
        {
            var rowIndex = 2;
            foreach (var item in data)
            {
                var columnIndex = 1;
                foreach (var property in properties)
                {
                    var attribute = property.GetCustomAttribute<CollumnAttribute>();
                    if (attribute != null)
                    {
                        var value = property.GetValue(item);
                        worksheet.Cell(rowIndex, columnIndex).SetValue(value.ToString());
                        columnIndex++;
                    }
                }
                rowIndex++;
            }
            return worksheet;
        }
    }
}