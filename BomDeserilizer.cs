using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Collections.Generic;
using System.Text.Json;
using Analysis.Base;
namespace iSchedule.Base
{
    public class BomData
    {
        public BomData() { }
        public string BOMDataId { get; set; }
        public BOMMaterial MaterialInfo { get; set; }
        public string WorkProcedureId { get; set; }
        public string WorkProcedureDesc { get; set; }
        public string ExtraInfo { get; set; }
        public string PreCutGraph { get; set; }
        public string CutGraph { get; set; }
        public string PressGraph { get; set; }
        public string CNCGraph { get; set; }
        public string InstallGraph { get; set; }
        public string CraftGraph { get; set; }
    }

    public class BOMMaterial
    {
        public List<BOMMaterialItem> Input { get; set; }
        public List<BOMMaterialItem> Output { get; set; }
    }

    public class BOMMaterialItem
    {
        public string MaterialId { get; set; }
        public int Number { get; set; }
    }
    public class BomDeserilizer

    {
        const string kSheetName = "BOM信息";

        const int kStartRow = 11;
        const int kBOMIdColumn = 0;
        const int kProcedureIdColumn = 1;
        const int kProcedureNameColumn = 2;
        const int kExtraData = 3;
        const int kInputMaterialId = 4;
        const int kInputMaterialNum = 5;
        const int kOutputMaterialId = 6;
        const int kOutputMaterialNumTotal = 7;
        const int KPreCutGraph = 9;
        const int KCutGraph = 10;
        const int KPressGraph = 11;
        const int KCNCGraph = 12;
        const int KInstallGraph = 13;
        const int KCraftGraph = 14;

        public class Result
        {
            public List<BomData>? bom_list { get; set; }
            public int error_code { get; set; }
        }
        public static Result DeserilizeBOMList(string file_path)
        {
            // If file is null or file name does not ends with `.xlsx`, `xls`, return BadRequest().
            if (!file_path.EndsWith(".xlsx") && !file_path.EndsWith(".xls"))
                return new Result { error_code = 1 };

            var hssfwb = DeserilizeUtil.LoadWorkbook(file_path);

            ISheet sheet = hssfwb.GetSheet(kSheetName);

            if (sheet == null)
                return new Result { error_code = -1, bom_list = null };

            List<BomData> bom_list = new List<BomData>();
            string last_bom_id = "";
            string last_input_material_id = "";
            string last_output_material_id = "";
            string last_procedure_id = "";
            string last_procedure_name = "";
            string last_extra_data = "";
            List<BOMMaterialItem> input_material_list = new List<BOMMaterialItem>();
            List<BOMMaterialItem> output_material_list = new List<BOMMaterialItem>();

            string PreCutGraph = "", CutGraph = "", PressGraph = "", CNCGraph = "", InstallGraph = "", CraftGraph = "";
            for (int row = kStartRow; row <= sheet.LastRowNum + 1; row++)
            {
                // Null is when the row only contains empty cells.
                if (sheet.GetRow(row) != null)
                {
                    string bom_id = sheet.GetRow(row).GetCell(kBOMIdColumn).StringCellValue;
                    if (string.IsNullOrEmpty(bom_id))
                    {
                        // BOM contains multiple lines.
                        var input_material_id = sheet.GetRow(row).GetCell(kInputMaterialId).StringCellValue;
                        if (!string.IsNullOrEmpty(input_material_id) && input_material_id != last_input_material_id)
                        {
                            last_input_material_id = input_material_id;
                            input_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_input_material_id,
                                Number = Analysis.Base.Functions.TryGetNumberFromCell(sheet.GetRow(row).GetCell(kInputMaterialNum))
                            });
                        }
                        // Handle output material.
                        var output_material_id = sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue;
                        if (!string.IsNullOrEmpty(output_material_id) && output_material_id != last_output_material_id)
                        {
                            last_output_material_id = sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue;
                            output_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_output_material_id,
                                Number = (int)sheet.GetRow(row).GetCell(kOutputMaterialNumTotal).NumericCellValue
                            });
                        }
                        continue;
                    }

                    if (bom_id != last_bom_id)
                    {
                        // First line of new BOM.
                        if (last_bom_id != "")
                        {
                            bom_list.Add(new BomData
                            {
                                BOMDataId = last_bom_id,
                                MaterialInfo = new BOMMaterial
                                {
                                    Input = input_material_list,
                                    Output = output_material_list
                                },
                                WorkProcedureId = last_procedure_id,
                                WorkProcedureDesc = last_procedure_name,
                                ExtraInfo = last_extra_data,
                                PreCutGraph = PreCutGraph,
                                CutGraph = CutGraph,
                                PressGraph = PressGraph,
                                CNCGraph = CNCGraph,
                                InstallGraph = InstallGraph,
                                CraftGraph = CraftGraph
                            });
                            last_input_material_id = sheet.GetRow(row).GetCell(kInputMaterialId).StringCellValue;
                            last_output_material_id = sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue;

                            PreCutGraph = sheet.GetRow(row).GetCell(KPreCutGraph).StringCellValue;
                            CutGraph = sheet.GetRow(row).GetCell(KCutGraph).StringCellValue;
                            PressGraph = sheet.GetRow(row).GetCell(KPressGraph).StringCellValue;
                            CNCGraph = sheet.GetRow(row).GetCell(KCNCGraph).StringCellValue;
                            InstallGraph = sheet.GetRow(row).GetCell(KInstallGraph).StringCellValue;
                            CraftGraph = sheet.GetRow(row).GetCell(KCraftGraph).StringCellValue;

                            last_procedure_id = sheet.GetRow(row).GetCell(kProcedureIdColumn).StringCellValue;
                            last_procedure_name = sheet.GetRow(row).GetCell(kProcedureNameColumn).StringCellValue;
                            last_extra_data = sheet.GetRow(row).GetCell(kExtraData).StringCellValue;
                            input_material_list = new List<BOMMaterialItem>();
                            output_material_list = new List<BOMMaterialItem>();
                            input_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_input_material_id,
                                Number = Analysis.Base.Functions.TryGetNumberFromCell(sheet.GetRow(row).GetCell(kInputMaterialNum)),
                            }) ; 

                            output_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_output_material_id,
                                Number = (int)sheet.GetRow(row).GetCell(kOutputMaterialNumTotal).NumericCellValue
                            });
                        }
                        else
                        {
                            // First line.
                            last_bom_id = bom_id;
                            last_input_material_id = sheet.GetRow(row).GetCell(kInputMaterialId).StringCellValue;
                            last_output_material_id = sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue;

                            last_procedure_id = sheet.GetRow(row).GetCell(kProcedureIdColumn).StringCellValue;
                            last_procedure_name = sheet.GetRow(row).GetCell(kProcedureNameColumn).StringCellValue;
                            last_extra_data = sheet.GetRow(row).GetCell(kExtraData).StringCellValue;

                            PreCutGraph = sheet.GetRow(row).GetCell(KPreCutGraph).StringCellValue;
                            CutGraph = sheet.GetRow(row).GetCell(KCutGraph).StringCellValue;
                            PressGraph = sheet.GetRow(row).GetCell(KPressGraph).StringCellValue;
                            CNCGraph = sheet.GetRow(row).GetCell(KCNCGraph).StringCellValue;
                            InstallGraph = sheet.GetRow(row).GetCell(KInstallGraph).StringCellValue;
                            CraftGraph = sheet.GetRow(row).GetCell(KCraftGraph).StringCellValue;

                            input_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_input_material_id,
                                Number = (int)sheet.GetRow(row).GetCell(kInputMaterialNum).NumericCellValue
                            });

                            output_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_output_material_id,
                                Number = (int)sheet.GetRow(row).GetCell(kOutputMaterialNumTotal).NumericCellValue
                            });
                        }

                    }
                    else
                    {
                        // Handle input material.
                        if (sheet.GetRow(row).GetCell(kInputMaterialId).StringCellValue != last_input_material_id)
                        {
                            last_input_material_id = sheet.GetRow(row).GetCell(kInputMaterialId).StringCellValue;
                            input_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_input_material_id,
                                Number = (int)sheet.GetRow(row).GetCell(kInputMaterialNum).NumericCellValue
                            });
                        }
                        // Handle output material.
                        if (sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue != last_output_material_id)
                        {
                            last_output_material_id = sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue;
                            output_material_list.Add(new BOMMaterialItem
                            {
                                MaterialId = last_output_material_id,
                                Number = (int)sheet.GetRow(row).GetCell(kOutputMaterialNumTotal).NumericCellValue
                            });
                        }
                    }
                    last_bom_id = bom_id;
                }
                else if (row <= sheet.LastRowNum + 1)
                {
                    bom_list.Add(new BomData
                    {
                        BOMDataId = last_bom_id,
                        MaterialInfo = new BOMMaterial
                        {
                            Input = input_material_list,
                            Output = output_material_list
                        },
                        PreCutGraph = PreCutGraph,
                        CutGraph = CutGraph,
                        PressGraph = PressGraph,
                        CNCGraph = CNCGraph,
                        InstallGraph = InstallGraph,
                        CraftGraph = CraftGraph,
                        WorkProcedureId = last_procedure_id,
                        WorkProcedureDesc = last_procedure_name,
                        ExtraInfo = last_extra_data,
                    });
                }
            }
            return new Result { error_code = 0, bom_list = bom_list };
        }
    }
}
