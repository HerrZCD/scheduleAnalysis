namespace iSchedule.Base
{
    public class DeserilizeUtil
    {
        public static NPOI.XSSF.UserModel.XSSFWorkbook? LoadWorkbook(string file_path)
        {
            // If file is null or file name does not ends with `.xlsx`, `xls`, return BadRequest().
            if (!file_path.EndsWith(".xlsx") && !file_path.EndsWith(".xls"))
            {
                return null;
            }

            // Load the excel file from file_path using NPOI.
            NPOI.XSSF.UserModel.XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(file_path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new NPOI.XSSF.UserModel.XSSFWorkbook(file);
            }
            return hssfwb;
        }
    }
}
