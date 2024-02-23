using iSchedule.Data;
using NPOI.SS.UserModel;

namespace Analysis.Base
{
    public class OrderDeserilizer
    {
        const string kSheetName = "发货清单";

        const int kStartRow = 2;
        const int kIndexColumn = 0;
        const int kProductTypeColumn = 1;   // 产品型号
        const int kProductScaleColumn = 2;
        const int kUnitColumn = 3;
        const int kTotalAmountColumn = 6;

        public class Result
        {
            public List<ProjectOrder>? orders { get; set; }
            public int error_code { get; set; }
        }
        public static Result DeserilizeOrders(string file_path)
        {
            // If file is null or file name does not ends with `.xlsx`, `xls`, return BadRequest().
            if (!file_path.EndsWith(".xlsx") && !file_path.EndsWith(".xls"))
                return new Result { error_code = -1, orders = null };

            var hssfwb = iSchedule.Base.DeserilizeUtil.LoadWorkbook(file_path);

            ISheet sheet = hssfwb.GetSheet(kSheetName);

            if (sheet == null)
                return new Result { error_code = -2, orders = null };

            List<ProjectOrder> orders = new List<ProjectOrder>();
            for (int row = kStartRow; row <= sheet.LastRowNum; row++)
            {
                // Null is when the row only contains empty cells.
                if (sheet.GetRow(row) != null)
                {
                    if (string.IsNullOrEmpty(sheet.GetRow(row).GetCell(kProductTypeColumn).StringCellValue))
                    {
                        continue;
                    }
                    orders.Add(new ProjectOrder
                    {
                        ProductType = sheet.GetRow(row).GetCell(kProductTypeColumn).StringCellValue,
                        TotalAmount = Base.Functions.TryGetNumberFromCell(sheet.GetRow(row).GetCell(kTotalAmountColumn)),
                    });
                }
            }
            return new Result { error_code = 0, orders = orders };
        }
    }
}
