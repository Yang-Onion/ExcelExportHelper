using System;

namespace ExcelExportHelper
{
    /// <summary>
    /// Excel生成样式
    /// </summary>
    [Flags]
    public enum ExcelStyle
    {
        /// <summary>
        /// 标题灰色背景
        /// </summary>
        title = 0x001,

        /// <summary>
        /// 左对齐
        /// </summary>
        left = 0x002,

        /// <summary>
        /// 右对齐
        /// </summary>
        right = 0x004,

        /// <summary>
        /// 时间格式，右对齐
        /// </summary>
        date = 0x008,

        /// <summary>
        /// 金钱格式，右对齐
        /// </summary>
        money = 0x016
    }
}