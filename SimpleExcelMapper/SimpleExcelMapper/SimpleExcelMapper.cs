using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace SimpleExcelMapper
{


    public class SimpleExcelMapper<Poco> where Poco : class, new ()  
    {
        public Poco Map(ISheet sheet, int rowid)
        {
            var row = sheet.GetRow(rowid);

            //get list of column header 
            var headerArray = sheet.GetRow(1).Cells.Cast<ICell>().Select(x => x.StringCellValue).ToArray();

            //get properties stored in Poco class
            var properties = typeof(Poco).GetProperties().Where(x => x.GetCustomAttributes(typeof(LabelAttribute), true).Any());

            Poco poco = new Poco();
            foreach (var prop in properties)
            {
                PropertyMapHelper.Map(prop, row, headerArray, poco, typeof(Poco));

            }
            return poco;
        }

        public IEnumerable<Poco> Map(ISheet sheet)
        {

            //get list of column header 
            var headerArray = sheet.GetRow(0).Cells.Cast<ICell>().Select(x => x.StringCellValue).ToArray();




            //get properties stored in Poco class
            var properties = typeof(Poco).GetProperties().Where(x => x.GetCustomAttributes(typeof(LabelAttribute), true).Any());

            List<Poco> pocos = new List<Poco>();

            IEnumerable<IRow> rows = Enumerable.Range(1, sheet.LastRowNum).Select(sheet.GetRow);

            foreach (IRow row in rows)
            {

                Poco poco = new Poco();
                foreach (var prop in properties)
                {
                    PropertyMapHelper.Map(prop, row, headerArray, poco, typeof(Poco));

                }
                pocos.Add(poco);
            }
            return pocos;
        }

        //private static string TestForStringValue(ICell cell)
        //{
        //    try
        //    {
        //        return cell.StringCellValue;
        //    }
        //    catch (Exception ex)
        //    {
        //        return "";
        //    }

        //}

    }
}
