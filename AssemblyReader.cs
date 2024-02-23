using iSchedule.Base;
using NPOI.SS.UserModel;

namespace Analysis
{
    internal class AssemblyReader
    {

        internal class Result
        {
            public string graphName;
            public int graphNumber;
        }
        const int kEmptySignalColumn = 5;
        public static List<Result> ReadAssembly(string file_path, string sheetName, int startingLine, int graphColumn, int numberColumn, int signalColumn)
        {
            var result = new List<Result>();
            var workBook = DeserilizeUtil.LoadWorkbook(file_path);

            ISheet sheet = workBook.GetSheet(sheetName);

            if (sheet == null)
                return result;

            for (int row = startingLine; row <= sheet.LastRowNum + 1; row++)
            {
                // Null is when the row only contains empty cells.
                if (sheet.GetRow(row) != null)
                {
                    try 
                    {
                        string signal = sheet.GetRow(row).GetCell(signalColumn).StringCellValue;
                        if (string.IsNullOrEmpty(signal))
                            break;
                    }
                    catch(System.NullReferenceException e)
                    {
                        break;
                    }

                    string gragh = sheet.GetRow(row).GetCell(graphColumn).StringCellValue;
                    int number = Base.Functions.TryGetNumberFromCell(sheet.GetRow(row).GetCell(numberColumn));
                    if (string.IsNullOrEmpty(gragh))
                        continue;

                    result.Add(new Result
                    {
                        graphName = gragh,
                        graphNumber = number
                    });
                }
            }
            return result;
        }
    }
}
