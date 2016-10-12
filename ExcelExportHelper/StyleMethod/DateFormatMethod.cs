using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelExportHelper
{
    /// <summary>
    /// 日期格式
    /// </summary>
    internal class DateFormatMethod : CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            IDataFormat format = workbook.CreateDataFormat();
            cellStyle.DataFormat = format.GetFormat("yyyy/mm/dd");
            return cellStyle;
        }
    }
}