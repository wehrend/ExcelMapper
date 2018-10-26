using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleExcelMapper
{
    public static class AttributeHelper
    {
        public static List<string> GetLabels(Type type, string propertyName)
        {
            var property = type.GetProperty(propertyName).GetCustomAttributes(false).Where(x => x.GetType().Name == "LabelAttribute").FirstOrDefault();
            if (property != null)
            {
                return ((LabelAttribute)property).ValueNames;
            }
            return new List<string>();
        }
    }
}
