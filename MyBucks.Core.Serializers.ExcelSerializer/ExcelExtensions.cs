using System.ComponentModel;
using System.Linq;
using System.Reflection;
using MyBucks.Core.Serializers.ExcelSerializer.Attributes;

namespace MyBucks.Core.Serializers.ExcelSerializer
{
    public static class ExcelExtensions
    {
        public static bool HeaderOnExcel(this PropertyInfo prop)
        {
            var headerAttrib = (HeaderOnExcelAttribute[])
                prop.GetCustomAttributes(typeof(HeaderOnExcelAttribute), false);

            return headerAttrib.Length > 0 && headerAttrib.First().Show;
        }
        
        public static string GetSpreadSheetHeaderPosition(this PropertyInfo prop)
        {
            var positions = (HeaderOnExcelAttribute[])
                prop.GetCustomAttributes(typeof(HeaderOnExcelAttribute), false);

            if (positions.Length == 0)
            {
                return "A1";
            }
            return positions[0].Position;
        }

        public static bool DataOnExcel(this PropertyInfo prop)
        {
            var dataAttrib = (DataOnExcelAttribute[])
                prop.GetCustomAttributes(typeof(DataOnExcelAttribute), false);

            return dataAttrib.Length > 0 && dataAttrib.First().Show;
        }

        public static string GetSpreadSheetDataStartPosition(this PropertyInfo prop)
        {
            var positions = (DataOnExcelAttribute[])
                prop.GetCustomAttributes(typeof(DataOnExcelAttribute), false);

            if (positions.Length == 0)
            {
                return "B1";
            }
            return positions[0].Position;
        }
        
        public static string GetDescription(this PropertyInfo prop)
        {
            var descriptions = (DescriptionAttribute[])
                prop.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (descriptions.Length == 0)
            {
                return prop.Name;
            }
            return descriptions[0].Description;
        }
    }
}