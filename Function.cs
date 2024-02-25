using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Analysis.Base
{
    public class Functions
    {
        static public int TryGetNumberFromCell(ICell cell)
        {
            try
            {
                int value = (int)cell.NumericCellValue;
                return value;
            }
            catch
            {
                int value = ParseInt(cell.StringCellValue);
                return value;
            }
        }

        static public string TryGetStringFromCell(ICell cell)
        {
            if (cell == null)
            {
                return "";
            }
            try
            {
                string value = cell.StringCellValue;
                return value;
            }
            catch
            {
                double value = cell.NumericCellValue;
                return value.ToString();
            }
        }

        public static bool IsMergedCell(ISheet sheet, int row, int column)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                if (row >= range.FirstRow && row <= range.LastRow && column >= range.FirstColumn && column <= range.LastColumn)
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsMergedWithPreviousRow(ISheet sheet, int row, int column)
        {
            if (row == 0) // 如果是第一行，不存在上一行，直接返回false
            {
                return false;
            }

            IRow currentRow = sheet.GetRow(row);
            IRow previousRow = sheet.GetRow(row - 1);

            if (currentRow == null || previousRow == null) // 如果当前行或上一行为空，返回false
            {
                return false;
            }

            bool isMergedCell = IsMergedCell(sheet, row, column);
            bool isMergedCellInPreviousRow = IsMergedCell(sheet, row - 1, column);

            if (isMergedCell != isMergedCellInPreviousRow) // 如果当前行和上一行的合并状态不一致，返回false
            {
                return false;
            }

            if (!isMergedCell || !isMergedCellInPreviousRow) // 如果当前行和上一行的单元格都没有合并，返回false
            {
                return false;
            }

            // 分别对两单元格取值，上下两个值有以下几种情况：
            // 1. 都为空，那么肯定是合并的
            // 2. 都不为空，那么肯定是不合并的
            // 3. 上面的不是空下面的时空，合并的
            // 4. 上面的是空下面的不是空，不合并的
            string currentCellValue = TryGetStringFromCell(currentRow.GetCell(column));
            string prevCellValue = TryGetStringFromCell(previousRow.GetCell(column));

            if (string.IsNullOrEmpty(currentCellValue) && string.IsNullOrEmpty(prevCellValue))
            {
                return true;
            } else if (!string.IsNullOrEmpty(currentCellValue) && !string.IsNullOrEmpty(prevCellValue))
            {
                return false;
            } else if (string.IsNullOrEmpty(currentCellValue) && !string.IsNullOrEmpty(prevCellValue))
            {
                return true;
            } else if (!string.IsNullOrEmpty(currentCellValue) && string.IsNullOrEmpty(prevCellValue))
            {
                return false;
            }



            return currentCellValue == prevCellValue;
        }

        //public static bool IsMergedWithPreviousRow(ISheet sheet, int row, int column)
        //{
        //    if (row == 0) // 如果是第一行，不存在上一行，直接返回false
        //    {
        //        return false;
        //    }

        //    IRow currentRow = sheet.GetRow(row);
        //    IRow previousRow = sheet.GetRow(row - 1);

        //    if (currentRow == null || previousRow == null) // 如果当前行或上一行为空，返回false
        //    {
        //        return false;
        //    }

        //    bool mergedWithPreviousRow = false;

        //    // 遍历当前行和上一行的所有单元格
        //    for (int i = 0; i <= column; i++)
        //    {
        //        ICell currentCell = currentRow.GetCell(i);
        //        ICell previousCell = previousRow.GetCell(i);

        //        // 检查当前单元格和上一行对应单元格是否都为空
        //        bool currentCellEmpty = (currentCell == null || currentCell.CellType == CellType.Blank);
        //        bool previousCellEmpty = (previousCell == null || previousCell.CellType == CellType.Blank);

        //        // 如果当前单元格和上一行对应单元格都为空，则继续检查下一个单元格
        //        if (currentCellEmpty && previousCellEmpty)
        //        {
        //            continue;
        //        }

        //        // 如果当前单元格和上一行对应单元格的合并状态相同，则说明当前单元格与上一行单元格是合并的
        //        bool currentCellMerged = IsMergedCell(sheet, row, i);
        //        bool previousCellMerged = IsMergedCell(sheet, row - 1, i);

        //        if (currentCellMerged == previousCellMerged)
        //        {
        //            mergedWithPreviousRow = true;
        //            break;
        //        }
        //    }

        //    return mergedWithPreviousRow;
        //}



        static public int ParseInt(string str)
        {
            int result;
            if (string.IsNullOrEmpty(str))
            {
                return 0;
            }
            if (int.TryParse(str, out result))
            {
                return result;
            }
            return 0;
        }
    }
}
