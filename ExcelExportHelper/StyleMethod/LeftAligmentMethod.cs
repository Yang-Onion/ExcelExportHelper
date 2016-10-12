using NPOI.SS.UserModel;

namespace ExcelExportHelper
{
    /// <summary>
    /// 左对齐
    /// </summary>
    internal class LeftAligmentMethod : CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            cellStyle.Alignment = HorizontalAlignment.Left;
            return cellStyle;
        }
    }
}