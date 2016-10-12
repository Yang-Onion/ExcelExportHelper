using NPOI.SS.UserModel;

namespace ExcelExportHelper
{
    /// <summary>
    /// 样式操作抽象类，使用多态执行多种操作
    /// </summary>
    internal abstract class CellStyleMethod
    {
        internal static IWorkbook workbook { get; set; }
        internal CellType CellType { get; set; }
        internal abstract ICellStyle SetCell(ICellStyle cellStyle);
    }
}