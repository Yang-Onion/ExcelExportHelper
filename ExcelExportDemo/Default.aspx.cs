using ExcelExportHelper;
using System;
using System.Collections.Generic;

namespace ExcelExportDemo
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var downLoad = new ExcelDownload("文件一", "页签一");
            var testList = new List<UserManagerTest>
            {
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "王二狗", Old = 20, Money=3.76
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                },
                new UserManagerTest
                {
                    CreateDate = DateTime.Now, Name = "李铁梅", Old = 30,Money=9.78
                }
            };
            downLoad.ExportExcel(testList);
        }
    }
}