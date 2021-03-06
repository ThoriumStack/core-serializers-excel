﻿using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Thorium.Core.DataIntegration.Attributes;

namespace Thorium.Core.Serializers.ExcelSerializer
{
    public static class ExcelExtensions
    {
        public static bool HeaderOnExcel(this PropertyInfo prop)
        {
            var headerAttrib = (ExcelColumnAttribute[])
                prop.GetCustomAttributes(typeof(ExcelColumnAttribute), false);

            return headerAttrib.Length > 0 && headerAttrib.First().ShowHeader;
        }

        public static string GetSpreadSheetHeaderPosition(this PropertyInfo prop)
        {
            var positions = (ExcelColumnAttribute[])
                prop.GetCustomAttributes(typeof(ExcelColumnAttribute), false);

            if (positions.Length == 0)
            {
                return "A1";
            }
            return positions[0].Column + positions[0].Row.ToString();
        }

        public static bool DataOnExcel(this PropertyInfo prop)
        {
            var dataAttrib = (ExcelColumnAttribute[])
                prop.GetCustomAttributes(typeof(ExcelColumnAttribute), false);

            return dataAttrib.Length > 0;
        }

        public static String GetSpreadSheetDataStartPosition(this PropertyInfo prop)
        {
            var positions = (ExcelColumnAttribute[])
                prop.GetCustomAttributes(typeof(ExcelColumnAttribute), false);

            if (positions.Length == 0)
            {
                return "B1";
            }
            if (positions[0].ShowHeader)
                return positions[0].Column + (positions[0].Row + 1).ToString();
            else
                return positions[0].Column + (positions[0].Row).ToString();
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