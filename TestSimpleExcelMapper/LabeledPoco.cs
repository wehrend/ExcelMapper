using SimpleExcelMapper;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestSimpleExcelMapper
{
    public class LabeledPoco
    {
        [LabelAttribute("Content", "Movie", "Title")]
        public string Title { get; set; }

        [LabelAttribute("ReportDate", "Date","date")]
        public DateTime ReportDate { get; set; }

        [LabelAttribute("ContractId", "Contract")]
        public string ContractID { get; set; }

        [LabelAttribute("Marketplace", "Country", "marketplace")]
        public string Territory { get; set; }

        [LabelAttribute("Fee", "Total")]
        public decimal Fee { get; set; }
    }
}
