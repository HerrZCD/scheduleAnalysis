using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Collections.Generic;
using System.Text.Json;
using Analysis.Base;
using NPOI.HSSF.Record;


namespace iSchedule.Base
{
    // This class used for BOMs related to cutting which not recorded in the original BOM.
    public class ExtendBomDeserilizer
    {

        const int kStartRow = 2;
        const int kBOMIdColumn = 4;
        const int kInputMaterialId = 2;
        const int kInputMaterialNum = 3;
        const int kOutputMaterialId = 6;
        const int kOutputMaterialSingle = 7;
        const int kOutputMaterialNumTotal = 8;

        public class Result
        {
            public List<BomData>? bom_list { get; set; }
            public int error_code { get; set; }
        }

        class OutputMaterial
        {
            public string MaterialId { get; set; }
            public int Number { get; set; }
        }
        public static Result DeserilizeExtendBOMList(string filePath, string sheetName)
        {
            // We assume all the BOMs releated to cutting will have more output than input, we will use the
            // output row to check if the bom read finished.

            // We assume there is a fixed rate between boms which share the same imput material, that say, if material A
            // produce 10 Bs in BOM b, so A will also procduce 10 Cs in BOM c.

            // 多个bom共享一个输入材料的时候，我们认为这个时候的语义是N中材料采用M种切法，但是一个bom对应一个输入物料产生
            // A, b. c三种物料，就简单读成一个bom。
            var hssfwb = DeserilizeUtil.LoadWorkbook(filePath);
            ISheet sheet = hssfwb.GetSheet(sheetName);
            if (sheet == null)
                return new Result { error_code = -2, bom_list = null };

            List<BomData> bom_list = new List<BomData>();
            string lastBOMId = "";
            string lastInputMaterialId = "";
            int lastInputMaterialNum = 0;

            int outputNumberOfInputMaterial = 0;
            List<BomData> tempList = new List<BomData>();
            List<BOMMaterialItem> output_material_list = new List<BOMMaterialItem>();
            for (int row = kStartRow; row <= sheet.LastRowNum; row++)
            {
                if (row != kStartRow)
                {
                    if (!Analysis.Base.Functions.IsMergedWithPreviousRow(sheet, row, kInputMaterialId))
                    {
                        Console.WriteLine("new line: "+ row);
                    }
                }
                // Null is when the row only contains empty cells.
                if (sheet.GetRow(row) != null)
                {
                    if (string.IsNullOrEmpty(sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue))
                    {
                        break;
                    }

                    string BOMId = sheet.GetRow(row).GetCell(kBOMIdColumn).StringCellValue;
                    string inputMaterialId = sheet.GetRow(row).GetCell(kInputMaterialId).StringCellValue;
                    string outputMaterialId = sheet.GetRow(row).GetCell(kOutputMaterialId).StringCellValue;
                    int outputMaterialNum = Analysis.Base.Functions.TryGetNumberFromCell(sheet.GetRow(row).GetCell(kOutputMaterialNumTotal));

                    if (!string.IsNullOrEmpty(BOMId) && BOMId != lastBOMId)
                    {
                        if (!string.IsNullOrEmpty(lastBOMId))
                        {
                            var bom = new BomData
                            {
                                BOMDataId = lastBOMId,
                                MaterialInfo = new BOMMaterial
                                {
                                    Input = new List<BOMMaterialItem> { new BOMMaterialItem { MaterialId = lastInputMaterialId, Number = lastInputMaterialNum } },
                                    Output = output_material_list
                                }
                            };
                            bom_list.Add(bom);
                            tempList.Add(bom);
                        }
                        // Reset the variables.
                        lastBOMId = BOMId;
                        output_material_list = new List<BOMMaterialItem> { new BOMMaterialItem { MaterialId = outputMaterialId, Number = outputMaterialNum } };
                    }

                    else if (string.IsNullOrEmpty(BOMId))
                    {
                        output_material_list.Add(new BOMMaterialItem { MaterialId = outputMaterialId, Number = outputMaterialNum });
                    }

                    if (!Analysis.Base.Functions.IsMergedWithPreviousRow(sheet, row, kInputMaterialId) && row != kStartRow)
                    {
                        // Need to update the number of bom.
                        float rate = (float)outputNumberOfInputMaterial / lastInputMaterialNum;
                        foreach (var item in tempList)
                        {
                            if (tempList.Count == 1)
                            {
                                foreach (var output in item.MaterialInfo.Output)
                                {

                                    output.Number = lastInputMaterialNum == 0 ? 0 : (int)output.Number / lastInputMaterialNum;
                                }
                            } else
                            {
                                foreach (var output in item.MaterialInfo.Output)
                                {
                                    output.Number = (int)rate;
                                }
                            }
                            

                            foreach (var input in item.MaterialInfo.Input)
                            {
                                input.Number = 1;
                            }
                        }


                        lastInputMaterialId = inputMaterialId;
                        tempList.Clear();
                        lastInputMaterialNum = Analysis.Base.Functions.TryGetNumberFromCell(sheet.GetRow(row).GetCell(kInputMaterialNum));
                        outputNumberOfInputMaterial = outputMaterialNum;
                    } 

                    else
                    {
                        outputNumberOfInputMaterial += outputMaterialNum;

                    }
                    
                }
            }

            // Add the last bom.
            var lastBom = new BomData
            {
                BOMDataId = lastBOMId,
                MaterialInfo = new BOMMaterial
                {
                    Input = new List<BOMMaterialItem> { new BOMMaterialItem { MaterialId = lastInputMaterialId, Number = lastInputMaterialNum } },
                    Output = output_material_list
                }
            };

            bom_list.Add(lastBom);
            tempList.Add(lastBom);

            // Handle the last bom number.
            float lastRate = (float)outputNumberOfInputMaterial / lastInputMaterialNum;
            foreach (var item in tempList)
            {
                foreach (var output in item.MaterialInfo.Output)
                {
                    output.Number = (int)lastRate;
                }

                foreach (var input in item.MaterialInfo.Input)
                {
                    input.Number = 1;
                }
            }
            return new Result { error_code = 0, bom_list = bom_list };
        }
    } // class ExtendBomDeserilizer
} // namespace iSchedule.Base
