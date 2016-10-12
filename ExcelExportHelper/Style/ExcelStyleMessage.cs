using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelExportHelper
{
    /// <summary>
    /// Excel样式管理
    /// </summary>
    internal static class ExcelStyleMessage
    {
        /// <summary>
        /// 样式集合
        /// </summary>
        private static Dictionary<string, ICellStyle> styleList { get; set; }

        static ExcelStyleMessage()
        {
            styleList = new Dictionary<string, ICellStyle>();
        }

        /// <summary>
        /// 获取枚举对应CellStyle
        /// </summary>
        /// <param name="excelStyle">Excel样式枚举</param>
        /// <returns></returns>
        internal static ICellStyle GetCellStyle<T>(T workbook, ExcelStyle excelStyle) where T : IWorkbook
        {
            if (styleList.ContainsKey(excelStyle.ToString()))
            {
                return styleList[excelStyle.ToString()];
            }
            ICellStyle _cellStyle = workbook.CreateCellStyle();
            _cellStyle.BorderTop = BorderStyle.Thin;
            _cellStyle.BorderRight = BorderStyle.Thin;
            _cellStyle.BorderLeft = BorderStyle.Thin;
            _cellStyle.BorderBottom = BorderStyle.Thin;
            CellStyleMethod styleMethod;
            if (excelStyle.ToString().IndexOf(',') > -1)
            {
                foreach (var styleItem in excelStyle.ToString().Replace(" ", "").Split(','))
                {
                    if (Enum.IsDefined(typeof(ExcelStyle), styleItem))
                    {
                        ExcelStyle styleModel = (ExcelStyle)Enum.Parse(typeof(ExcelStyle), styleItem, true);
                        styleMethod = GetStyleMethod(styleModel);
                        styleMethod.SetCell(_cellStyle);
                    }
                }
                return _cellStyle;
            }
            styleMethod = GetStyleMethod(excelStyle);
            styleMethod.SetCell(_cellStyle);
            styleList.Add(excelStyle.ToString(), _cellStyle);
            return _cellStyle;
        }

        /// <summary>
        /// 根据枚举加载对应操作类
        /// </summary>
        /// <param name="excelStyle">样式枚举</param>
        /// <returns>操作类</returns>
        private static CellStyleMethod GetStyleMethod(ExcelStyle excelStyle)
        {
            switch (excelStyle)
            {
                case ExcelStyle.title:
                    return new TitleBackgroundMethod();

                case ExcelStyle.left:
                    return new LeftAligmentMethod();

                case ExcelStyle.right:
                    return new RightAligmentMethod();

                case ExcelStyle.date:
                    return new DateFormatMethod();

                case ExcelStyle.money:
                    return new MoneyFormatMethod();

                default:
                    throw new ArgumentException("参数无效");
            }
        }
    }
}