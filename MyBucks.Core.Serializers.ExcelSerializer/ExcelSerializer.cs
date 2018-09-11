using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using MyBucks.Core.DataIntegration.Interfaces;
using OfficeOpenXml;
using MyBucks.Core.DataIntegration.Attributes;

namespace MyBucks.Core.Serializers.ExcelSerializer
{
    public class ExcelSerializer : IIntegrationDataSerializer
    {
        public bool HasHeaderRecord { get; set; } = false;

        public bool AddDate { get; set; } = false;

        public string WorkSheetName { get; set; } = "Sheet1";


        public MemoryStream GenerateRawData<TData>(IEnumerable<TData> data)
        {
            var outputMemoryStream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var dataWorksheet = package.Workbook.Worksheets.Add(WorkSheetName);

                // Add a date to the worksheet (optional)
                if (AddDate)
                {
                    dataWorksheet.Cells["A1"].Value = "Date";
                    dataWorksheet.Cells["B1"].Value = DateTime.Now.ToString("dd.MM.yyyy");
                    dataWorksheet.Cells["A1:B1"].Style.Font.Bold = true;
                }

                //Add a header record to worksheet (optional)
                if (HasHeaderRecord)
                {
                    var headers = GetExcelHeaders(typeof(TData));

                    if (headers != null && headers.Count > 0)
                    {
                        foreach (var header in headers)
                        {
                            dataWorksheet.Cells[header.Item1].Value = header.Item2;
                            dataWorksheet.Cells[header.Item1].Style.Font.Bold = true;
                        }
                    }
                }

                //Add the data to the worksheet (compulsory)
                var rows = GetExcelData<TData>(data);
                
                foreach (var row in rows)
                {
                    dataWorksheet.Cells[row.Item1].Value = row.Item2;
                }

                package.SaveAs(outputMemoryStream);
            }
            return outputMemoryStream;
        }

        private List<Tuple<string,string>> GetExcelData<TData>(IEnumerable<TData> rows)
        {
            var result = new List<Tuple<string, string>>();

            var increment = 0;

            foreach (var row in rows)
            {
                var properties = row.GetType().GetProperties().Where(c => c.DataOnExcel()).ToList();

                foreach (var property in properties)
                {
                    var value = property.GetValue(row, null).ToString();
                    result.Add(new Tuple<string, string>(RowIncrement(property.GetSpreadSheetDataStartPosition(), increment), value));
                }

                increment += 1;
            }

            return result;
        }

        private List<Tuple<string,string>> GetExcelHeaders(Type model)
        {
            var result = new List<Tuple<string, string>>();
            var properties = model.GetProperties().Where(c => c.HeaderOnExcel()).ToList();

            foreach (var property in properties)
            {
                result.Add(new Tuple<string,string>(property.GetSpreadSheetHeaderPosition(),property.GetDescription()));
            }

            return result;
        }

        private string RowIncrement(string input,int increment)
        {
            var result = input;

            if (input != null && input.Length == 2)
            {
                var Letter = (input.ToCharArray())[0];
                var numericStr = (input.ToCharArray())[1].ToString();

                var numeric = 0;
                if (int.TryParse(numericStr,out numeric))
                {
                    result = $"{Letter}{numeric + increment}";
                }
            }
            return result;
        }

        
        public IEnumerable<TData> GetData<TData>(MemoryStream rawData) where TData : new()
        {
            ExcelPackage excel = new ExcelPackage(rawData);
            var workSheet = excel.Workbook.Worksheets.First();

            var collection = ConvertSheetToObjects<TData>(workSheet);

            return collection;
        }

        public void ReadMany<TData, TDiscriminator>(IList<TData> destination, Func<TDiscriminator, bool> discriminator, MemoryStream stream)
            where TData : new()
            where TDiscriminator : new()
        {
            throw new NotImplementedException("Reading has not been implemented for excel files.");
        }

        public void ReadSingle<TData, TDiscriminator>(Action<TData> assignAction, Func<TDiscriminator, bool> discriminator, MemoryStream rawData)
            where TData : new()
            where TDiscriminator : new()
        {
            throw new NotImplementedException("Reading has not been implemented for excel files.");
        }

        private static IEnumerable<T> ConvertSheetToObjects<T>(ExcelWorksheet worksheet) where T : new()
        {

            Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(ExcelColumnAttribute);

            var columns = typeof(T)
                    .GetProperties()
                    .Where(x => x.CustomAttributes.Any(columnOnly))
            .Select(p => new
            {
                Property = p,
                Column = p.GetCustomAttributes<ExcelColumnAttribute>().First().ColumnNumber //safe because if where above
            }).ToList();


            var rows = worksheet.Cells
                .Select(cell => cell.Start.Row)
                .Distinct()
                .OrderBy(x => x);


            //Create the collection container
            var collection = rows.Skip(1)
                .Select(row =>
                {
                    var tnew = new T();
                    columns.ForEach(col =>
                    {
                        //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                        var val = worksheet.Cells[row, col.Column];
                        //If it is numeric it is a double since that is how excel stores all numbers
                        if (val.Value == null)
                        {
                            col.Property.SetValue(tnew, null);
                            return;
                        }
                        if (col.Property.PropertyType == typeof(Int32))
                        {
                            col.Property.SetValue(tnew, val.GetValue<int>());
                            return;
                        }
                        if (col.Property.PropertyType == typeof(double))
                        {
                            col.Property.SetValue(tnew, val.GetValue<double>());
                            return;
                        }
                        if (col.Property.PropertyType == typeof(DateTime))
                        {
                            col.Property.SetValue(tnew, val.GetValue<DateTime>());
                            return;
                        }
                        if (col.Property.PropertyType == typeof(decimal))
                        {
                            col.Property.SetValue(tnew, val.GetValue<decimal>());
                            return;
                        }
                        //Its a string
                        col.Property.SetValue(tnew, val.GetValue<string>());
                    });

                    return tnew;
                });


            //Send it back
            return collection;
        }
    }
}
