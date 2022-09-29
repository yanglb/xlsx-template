using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateGO.Utils
{
    public static class CellUtils
    {
        /// <summary>
        /// 获取单元格字符内容
        /// </summary>
        public static string GetCellString(Cell cell, SharedStringTable? sharedStringTable)
        {
            var value = cell.InnerText;
            if (string.IsNullOrEmpty(value)) return "";

            // 从 SharedStringTable 中获取
            if (cell.DataType != null && cell.DataType == CellValues.SharedString && sharedStringTable != null)
            {
                value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
