using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Shouldly;
using SimpleExcelMapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace TestSimpleExcelMapper
{
    public class TestSimpleExcelMapper
    {
        [Fact]
        public void Test_Generic_Mapping()
        {
            //arrange 

            var headerCol = new InputPoco()
            {
                Title = "Title",
                ReportDate = "Date",
                ContractID = "ContractId",
                Territory = "Territory",
                Fee = "Fee"
            };

            var input = new List<InputPoco>{ new InputPoco
                {
                    Title = "Movie I",
                    ReportDate = DateTime.Now.Date.ToString(),
                    ContractID = "xox",
                    Territory = "DE",
                    Fee = (2.99m).ToString()
                }
            };

            var workbook = CreateTestWorkbook(headerCol, input );
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.GetRow(0);

            SimpleExcelMapper<LabeledPoco> simpleMapper = new SimpleExcelMapper<LabeledPoco>();
            List<LabeledPoco> result = simpleMapper.Map(sheet).ToList();

            var expectedResult = new LabeledPoco
                {
                    Title = "Movie I",
                    ReportDate = DateTime.Now.Date,
                    ContractID = "xox",
                    Territory = "DE",
                    Fee = 2.99m
                };

            result.Count().ShouldBe(1);
            
            result.First().Title.ShouldBe(expectedResult.Title);
            result.First().ReportDate.ShouldBe(expectedResult.ReportDate);
            result.First().ContractID.ShouldBe(expectedResult.ContractID);
            result.First().Territory.ShouldBe(expectedResult.Territory);
            result.First().Fee.ShouldBe(expectedResult.Fee);

        }

        [Fact]
        public void Test_Mapping_With_Correct_Arbitrary_Labels()
        {
            //arrange 

            var headerCol = new InputPoco()
            {
                Title = "Content",
                ReportDate = "Date",
                ContractID = "Contract",
                Territory = "Marketplace",
                Fee = "Total"
            };

            var input = new List<InputPoco>{ new InputPoco
                {
                    Title = "Movie I",
                    ReportDate = DateTime.Now.Date.ToString(),
                    ContractID = "xox",
                    Territory = "DE",
                    Fee = (2.99m).ToString()
                }
            };

            var workbook = CreateTestWorkbook(headerCol, input);
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.GetRow(0);

            SimpleExcelMapper<LabeledPoco> simpleMapper = new SimpleExcelMapper<LabeledPoco>();
            List<LabeledPoco> result = simpleMapper.Map(sheet).ToList();

            var expectedResult = new LabeledPoco
            {
                Title = "Movie I",
                ReportDate = DateTime.Now.Date,
                ContractID = "xox",
                Territory = "DE",
                Fee = 2.99m
            };

            result.First().Title.ShouldBe(expectedResult.Title);
            result.First().ReportDate.ShouldBe(expectedResult.ReportDate);
            result.First().ContractID.ShouldBe(expectedResult.ContractID);
            result.First().Territory.ShouldBe(expectedResult.Territory);
            result.First().Fee.ShouldBe(expectedResult.Fee);
        }

        [Fact]
        public void Test_Mapping_With_Some_Incorrect_Arbitrary_Labels()
        {
            //arrange 

            var headerCol = new InputPoco()
            {
                Title = "Content",
                ReportDate = "Date",
                ContractID = "report_type",
                Territory = "Territorities",
                Fee = "Total"
            };

            var input = new List<InputPoco>{ new InputPoco
                {
                    Title = "Movie I",
                    ReportDate = DateTime.Now.Date.ToString(),
                    ContractID = "xox",
                    Territory = "DE",
                    Fee = (2.99m).ToString()
                }
            };

            var workbook = CreateTestWorkbook(headerCol, input);
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.GetRow(0);

            SimpleExcelMapper<LabeledPoco> simpleMapper = new SimpleExcelMapper<LabeledPoco>();
            List<LabeledPoco> result = simpleMapper.Map(sheet).ToList();

            var expectedResult = new LabeledPoco
            {
                Title = "Movie I",
                ReportDate = DateTime.Now.Date,
                ContractID = "xox",
                Territory = "DE",
                Fee = 2.99m
            };


            result.First().Title.ShouldBe(expectedResult.Title);
            result.First().ReportDate.ShouldBe(expectedResult.ReportDate);

            result.First().ContractID.ShouldNotBe(expectedResult.ContractID);
            result.First().Territory.ShouldNotBe(expectedResult.Territory);

            result.First().Fee.ShouldBe(expectedResult.Fee);
        }



        [Fact]
        public void Test_String_Mapping()
        {
            var headerCol = new InputPoco()
            {
                Title = "Title",
                ReportDate = "Date",
                ContractID = "ContractId",
                Territory = "Territory",
                Fee = "Fee"
            };

            var input = new List<InputPoco>{ new InputPoco
                {
                    Title = "Movie I",
                }
            };

            var workbook = CreateTestWorkbook(headerCol, input);
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.GetRow(0);

            SimpleExcelMapper<LabeledPoco> simpleMapper = new SimpleExcelMapper<LabeledPoco>();
            List<LabeledPoco> result = simpleMapper.Map(sheet).ToList();

            var expectedResult = new LabeledPoco
            {
                Title = "Movie I",
            };
            
            result.First().Title.ShouldBe(expectedResult.Title);
            result.First().ReportDate.ShouldBe(expectedResult.ReportDate);
            result.First().ContractID.ShouldBe(expectedResult.ContractID);
            result.First().Territory.ShouldBe(expectedResult.Territory);
            result.First().Fee.ShouldBe(expectedResult.Fee);


        }



        [Fact]
        public void Test_DateTime_Mapping()
        {
            var headerCol = new InputPoco()
            {
                Title = "Title",
                ReportDate = "Date",
                ContractID = "ContractId",
                Territory = "Territory",
                Fee = "Fee"
            };

            var input = new List<InputPoco>{ new InputPoco
                {
                    ReportDate = DateTime.Now.ToShortDateString()
                }
            };

            var workbook = CreateTestWorkbook(headerCol, input);
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.GetRow(0);

            SimpleExcelMapper<LabeledPoco> simpleMapper = new SimpleExcelMapper<LabeledPoco>();
            List<LabeledPoco> result = simpleMapper.Map(sheet).ToList();

            var expectedResult = new LabeledPoco
            {
                ReportDate = DateTime.Now.Date
            };
            
            result.First().Title.ShouldBe(expectedResult.Title);
            result.First().ReportDate.ShouldBe(expectedResult.ReportDate);
            result.First().ContractID.ShouldBe(expectedResult.ContractID);
            result.First().Territory.ShouldBe(expectedResult.Territory);
            result.First().Fee.ShouldBe(expectedResult.Fee);
        }

        private static XSSFWorkbook CreateTestWorkbook( InputPoco header, List<InputPoco> inputList)
        {
            //create Test Worksheet
            var workbook = CreateBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var headerRow = sheet.CreateRow(0);
            
            //header starts on zero
            ProcessRow(sheet.CreateRow(0), header);
            //then the rest starts
            int i = 1;
            foreach (var item in inputList)
            {
                ProcessRow(sheet.CreateRow(i), item);
                i++;
            }
            return workbook;
        }

        private static void ProcessRow(IRow row, InputPoco input)
        {
            var properties = input.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);

            int j = 0;
            foreach (var prop in properties)
            {
                var value =  prop.GetValue(input, null);
                row.CreateCell(j).SetCellValue( (value != null)?  (string) value : null);

                j++;
                
            }

        }


        private static XSSFWorkbook CreateBlankWorkbook()
        {
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("sheet1");
            return workbook;
        }

        private static void SaveTestWorkbook(XSSFWorkbook workbook, string filename)
        {
            // save workbook locally
            FileStream excelfile = new FileStream(filename, FileMode.Create, System.IO.FileAccess.Write);
            workbook.Write(excelfile);
            excelfile.Close();
        }
    }

    public class InputPoco
    {
        public string Title {get; set;}
        public string ReportDate { get; set; }
        public string ContractID { get; set; }
        public string Territory { get; set; }
        public string Fee { get; set; }
    }
}
