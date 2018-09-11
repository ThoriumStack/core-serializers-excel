using System;
using System.Collections.Generic;
using System.ComponentModel;
using Xunit;
using MyBucks.Core.Serializers.ExcelSerializer;
using MyBucks.Core.DataIntegration.Attributes;
using MyBucks.Core.DataIntegration.Transports;
using System.IO;
using System.Linq;
using MyBucks.Core.DataIntegration;

namespace Tests
{
    public class TestExcel
    {
        [Fact]
        public void TestExcelRead()
        {
            List<ExcelTestClass> data = new List<ExcelTestClass>();

            var serializer = new ExcelSerializer();
            serializer.HasHeaderRecord = true;
            StreamTransport transport = new StreamTransport();
            transport.InputStream = GenerateRawData();

            var dataIntegrator = new Integrator();

            var build = new InputBuilder()
                .SetSerializer(serializer)
                .ReadAll(data);

            dataIntegrator.ReceiveData(build, transport);

            Assert.NotNull(data);
            Assert.Equal(2, data.Count);
            Assert.Equal("Thomas Jefferson", data[1].Name);
            Assert.Equal(34.8M, data[0].NetWorth);
        }

        [Fact]
        public void TestExcelWrite()
        {
            var serializer = new ExcelSerializer();
            serializer.HasHeaderRecord = true;
            StreamTransport transport = new StreamTransport();

            var build = new OutputBuilder()
                        .SetSerializer(serializer)
                        .AddListData(GetTestData());

            var dataIntegrator = new Integrator();

            var result = dataIntegrator.SendData(build, transport);

            Stream ResultStream = transport.GetLastRawData();

            Assert.True(ResultStream.Length > 0, "Stream cannot be empty");
        }

        private MemoryStream GenerateRawData()
        {
            var s = new ExcelSerializer();
            s.HasHeaderRecord = true;
            var testData = GetTestData();

            var stream = s.GenerateRawData(testData);
            stream.Position = 0;
            return stream;
        }

        private List<ExcelTestClass> GetTestData()
        {
            return new List<ExcelTestClass> {
                new ExcelTestClass { BirthDate = DateTime.MinValue, Name = "George Washington", NetWorth = 34.8M},
                new ExcelTestClass { BirthDate = DateTime.MinValue.AddYears(80), Name = "Thomas Jefferson", NetWorth =7451254.54M},
            };
        }

    }

    public class ExcelTestClass
    {
        [Description("First Name")]
        [ExcelColumn("A", 1)]
        public string Name { get; set; }

        [Description("Birth Date")]
        [ExcelColumn("B", 1)]
        public DateTime BirthDate { get; set; }

        [Description("Net Worth")]
        [ExcelColumn("C", 1)]
        public Decimal NetWorth { get; set; }
    }
}