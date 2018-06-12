using System;

namespace MyBucks.Core.Serializers.ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class)]
    public class DataOnExcelAttribute : System.Attribute
    {
        public bool Show { get; set; } = true;
        public string Position { get; set; } = "B1";

        public DataOnExcelAttribute(string position)
        {
            Position = position;
        }
    }
}