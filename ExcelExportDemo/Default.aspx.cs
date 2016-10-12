using ExcelExportHelper;
using System;
using System.Collections.Generic;

namespace ExcelExportDemo
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            List<UserManagerTest> testList = new List<UserManagerTest>
            {
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "王二狗", Old = 20, Money=3.76
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                }
            };
            ExcelDownload downLoad = new ExcelDownload("员工信息", "年度员工汇总");
            downLoad.ExportExcel(testList);
        }
    }
}