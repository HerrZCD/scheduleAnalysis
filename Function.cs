using NPOI.SS.UserModel;

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
