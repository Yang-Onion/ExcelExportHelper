using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelExportHelper
{
    /// <summary>
    /// 金额格式方法
    /// </summary>
    internal class MoneyFormatMethod : CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            IDataFormat format = workbook.CreateDataFormat();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("￥#,##0");
            return cellStyle;
        }
    }
}