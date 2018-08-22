using System;
using System.Collections.Generic;
using System.Text;

namespace MyBucks.Core.Serializers.ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class)]
    public class ExcelColumnAttribute: System.Attribute
    {
        public int ColumnIndex { get; set; }

        public ExcelColumnAttribute(int column)
        {
            ColumnIndex = column;
        }
    }
}
