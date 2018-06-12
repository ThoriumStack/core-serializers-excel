using System;

namespace MyBucks.Core.Serializers.ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class)]
    public class HeaderOnExcelAttribute : System.Attribute
    {
        public bool Show { get; set; } = true;
        public string Position { get; set; } = "A1";

        public HeaderOnExcelAttribute(string position)
        {
            Position = position;
        }

    }
}