using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleExcelMapper
{
    [AttributeUsage(AttributeTargets.Property)]
    public class LabelAttribute : Attribute
    {
        public List<string> _valueNames { get; set; }

        public List<string> ValueNames
        {
            get
            {
                return _valueNames;
            }
            set
            {
                _valueNames = value;
            }
        }

        public LabelAttribute()
        {
            _valueNames = new List<string>();
        }

        public LabelAttribute(params string[] valueNames)
        {
            _valueNames = valueNames.ToList();
        }
    }
}
