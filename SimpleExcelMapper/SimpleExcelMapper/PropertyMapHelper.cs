using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;


//based on https://github.com/exceptionnotfound/DataNamesMappingDemo and modified for 
// mapping of excel to poco via NPOI

namespace SimpleExcelMapper
{
    public static class PropertyMapHelper
    {
        public static void Map(PropertyInfo property, IRow row, Array headerArray, object poco, Type pocoType)
        {
            List<string> columnNamesFromPocoAttributes = AttributeHelper.GetLabels(pocoType, property.Name);

            foreach (var columnName in columnNamesFromPocoAttributes)
            {
                if (!String.IsNullOrWhiteSpace(columnName))
                {
                    var columnIndex = Array.IndexOf(headerArray, columnName);
                    if (columnIndex > -1)
                    {
                        var propertyValue = row.Cells[columnIndex];
                        if (propertyValue != null)
                        {
                            ParsePrimitive(property, poco, propertyValue);
                            break;
                        }
                        
                    }
                

                }
            }

        }

        private static void ParsePrimitive(PropertyInfo prop, object poco, object value)
        {
            if (prop.PropertyType == typeof(string))
            {
                prop.SetValue(poco, value.ToString().Trim(), null);
            }
            else if (prop.PropertyType == typeof(int))
            {
                prop.SetValue(poco, int.Parse(value.ToString()), null);
            }
            else if (prop.PropertyType == typeof(decimal))
            {
                Decimal result;
                bool isValidDecimal = Decimal.TryParse(value.ToString(),NumberStyles.Currency, new CultureInfo("de-DE"), out result);
                if (isValidDecimal)
                {
                    prop.SetValue(poco, result, null);
                }
            }
            else if (prop.PropertyType == typeof(DateTime))
            {
                DateTime date;
                bool isValidDate = DateTime.TryParse(value.ToString(), out date);
                if (isValidDate)
                {
                    prop.SetValue(poco, date, null);
                }
                else
                {
                    isValidDate = DateTime.TryParseExact(value.ToString(), "MMddyyyy", new CultureInfo("en-US"), DateTimeStyles.AssumeLocal, out date);
                    if (isValidDate)
                    {
                        prop.SetValue(poco, date, null);
                    }
                }
            }
        }
    }
}
