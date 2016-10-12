using System;

namespace ExcelExportHelper
{
    /// <summary>
    /// 实体类生成特性
    /// 可以控制显示中文名
    /// </summary>
    public class ExcelInfoAttribute : Attribute
    {
        /// <summary>
        /// 显示中文名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 列宽
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 列样式
        /// </summary>
        public ExcelStyle ExcelStyle { get; set; }

        /// <summary>
        /// 默认左对齐，宽度2800
        /// </summary>
        public ExcelInfoAttribute(string name)
        {
            Name = name;
            Width = 2800;
            ExcelStyle = ExcelStyle.left;
        }
    }
}