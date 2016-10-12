using ExcelExportHelper;
using System;

namespace ExcelExportDemo
{
    public class UserManagerTest
    {
        [ExcelInfo("名称")]
        public string Name { get; set; }
        [ExcelInfo("年龄", ExcelStyle = ExcelStyle.left)]
        public int Old { get; set; }
        [ExcelInfo("金额", ExcelStyle = ExcelStyle.money)]
        public double Money { get; set; }
        [ExcelInfo("时间", ExcelStyle = ExcelStyle.date | ExcelStyle.right)]
        public DateTime CreateDate { get; set; }
    }
}