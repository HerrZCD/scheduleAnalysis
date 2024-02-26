using iSchedule.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

using GraphID = System.String;
using GraphRepeatTimes = System.Int32;

namespace Analysis
{
    internal class GlobalCache
    {
        public static Dictionary<GraphID, GraphRepeatTimes> AllWorksToDo = new Dictionary<GraphID, GraphRepeatTimes>();
        public static Dictionary<GraphID, GraphRepeatTimes> WorksToDoOfSingleProduct = new Dictionary<GraphID, GraphRepeatTimes>();
        public static List<BomData> BOMList = new List<BomData>();
        public static List<AssemblyReader.Result> RequiredProducts = new List<AssemblyReader.Result>();
        public static void SetBOMList(List<BomData> list)
        {
            BOMList = list;
        }

        public static void AppendBOMData(BomData data)
        {
            var exist = BOMList.Find(bom => bom.BOMDataId == data.BOMDataId);
            if (exist != null)
            {
                BOMList.Remove(exist);
                exist.MaterialInfo = data.MaterialInfo;
                BOMList.Add(exist);
            } 
            else
            {
                BOMList.Add(data);
            }
        }

        public static void WriteAllBomsToJSON()
        {
            var json = JsonSerializer.Serialize(BOMList);
            System.IO.File.WriteAllText(@"./BOMList.json", json);
        }

        public static void LoadAllBomsFromJSON()
        {
            string json = System.IO.File.ReadAllText(@"./BOMList.json");
            BOMList = JsonSerializer.Deserialize<List<BomData>>(json);

            if (BOMList == null)
            {
                Console.WriteLine("Failed to load BOMList from file.");
            }

            Console.WriteLine($"Loaded {BOMList.Count} BOMs");
        }

        public static void AppendWork(GraphID id, GraphRepeatTimes times) { 
            if (AllWorksToDo.ContainsKey(id))
            {
                AllWorksToDo[id] += times;
            }
            else
            {
                AllWorksToDo.Add(id, times);
            }
        }
        


        public static void AppendRequiredProducts(List<AssemblyReader.Result> list)
        {
            RequiredProducts.AddRange(list);
        }
        public static void AppendProduct(string id, int number)
        {
            RequiredProducts.Add(new AssemblyReader.Result
            {
                graphName = id,
                graphNumber = number
            });
        }

        public static void InfoShowBOMList()
        {
            Console.WriteLine("===============BOMList===============");
            foreach (var bom in BOMList)
            {
                Console.WriteLine("=============== " + bom.BOMDataId + " ===============");
                foreach (var input in bom.MaterialInfo.Input)
                {
                    Console.WriteLine("input: " + input.MaterialId + "  number: " + input.Number);
                }
            }
        }

        public static void InfoShowCache()
        {
            //Console.WriteLine("===============BOMList===============");
            //foreach (var bom in BOMList)
            //{
            //    Console.WriteLine(bom.graphName + " " + bom.graphNumber);
            //}

            Console.WriteLine("===============RequiredProducts===============");
            foreach (var rp in RequiredProducts)
            {
                Console.WriteLine(rp.graphName + " " + rp.graphNumber);
            }
        }

        public static List<BomData> FindBOMsContainsGivenProduct(string id)
        {
            return BOMList.FindAll(bom => bom.MaterialInfo.Output.Any(output => output.MaterialId == id));
        }
        public static BomData FindBOMFromGivenOutput(string id)
        {
            var list = FindBOMsContainsGivenProduct(id);
            if (list == null) return null;
            foreach (var bom in list)
            {
                 // Only check no loop exist.
                if (bom.MaterialInfo.Input.Any(input => input.MaterialId == id))
                {
                    continue;
                }
                else
                {
                     return bom;
                }
            }
            return null;
        }

        // We check an id, if this id do not exist in the BOMList, we should stop.
        // We record all the works from the topest bom to the most fundamental bom.
        public static void TraverseHandleBOM(string id, int number)
        {
            // 对于要生产的某个产品，我们先看看有没有指定的BOMid和这个id匹配，如果有，我们要检查
            // 是不是LOOP的情况。

            // 对于BOMid是空和LOOP的情况，我们循环这个BOM列表找到输出产品包含该产品的BOM/套料图
            if (string.IsNullOrEmpty(id) || number == 0)
            {
                return;
            }

            var bom = BOMList.Find(bom => bom.BOMDataId == id);
            if (bom == null)
            {
                // We should stop here. This may cause by
                // 1. The id is not in the BOMList. We should fix the bom list in our system.
                // 2. This id is a product, we should not go further.
                Console.WriteLine("BOM id not in BOM list: "+ id + " number: " + number);
                bom = FindBOMFromGivenOutput(id);

                if (bom != null)
                {
                    Console.WriteLine("找到套料图: " + bom.BOMDataId);
                }
                else
                {
                    Console.WriteLine("找不到套料图: " + id);
                    return;
                }
            }    
            else
            {
                foreach (var subBom in bom.MaterialInfo.Input)
                {
                    if (id == subBom.MaterialId)
                    {
                        Console.WriteLine($"发现BOMId和输入相同的情况: {id}, 正在搜寻合适的套料图...");
                        bom = FindBOMFromGivenOutput(id);
                        if (bom != null)
                        {
                            Console.WriteLine("找到套料图: " + bom.BOMDataId);
                        }
                        else
                        {
                            Console.WriteLine("找不到套料图: " + id);
                            return;
                        }
                    }
                }
            }

            int outputNum = 0;
            foreach (var o in bom.MaterialInfo.Output)
            {
                if (o.MaterialId == id)
                {
                    outputNum = o.Number;
                    break;
                    // TODO:
                    // Record all output.
                }
            }

            if (outputNum == 0)
            {
                // 如果产出就一个的话， 那么我们认为BOMid和产出是别名关系。
                if (bom.MaterialInfo.Output.Count == 1)
                {
                    outputNum = bom.MaterialInfo.Output[0].Number;
                }
                else
                {
                    Console.WriteLine($"BOM 的输出和BOM id {id}没对上");
                    return;
                }
            }

            int needRepeatTimes = (int)Math.Ceiling((double)number / (double)outputNum);
            // Record the work.
            if (AllWorksToDo.ContainsKey(bom.BOMDataId))
            {
                AllWorksToDo[bom.BOMDataId] += number;
            }
            else
            {
                AllWorksToDo.Add(bom.BOMDataId, number);
            }

            if (WorksToDoOfSingleProduct.ContainsKey(bom.BOMDataId))
            {
                WorksToDoOfSingleProduct[bom.BOMDataId] += number;
            }
            else
            {
                WorksToDoOfSingleProduct.Add(bom.BOMDataId, number);
            }

            // Go further.
            foreach (var i in bom.MaterialInfo.Input)
            {
                TraverseHandleBOM(i.MaterialId, i.Number * needRepeatTimes);
            }
        }

        public static void WriteAllWorksToFile()
        {
            List<AssemblyReader.Result> results = new List<AssemblyReader.Result>();
            results.Add(GlobalCache.RequiredProducts[2]);
            foreach (var productInfo in GlobalCache.RequiredProducts)
            // foreach (var productInfo in results)
            {
                WorksToDoOfSingleProduct.Clear();
                Console.Clear();
                Console.WriteLine($"正在处理产品: {productInfo.graphName}, 这个订单需要{productInfo.graphNumber}个");
                int repeatTimes = productInfo.graphNumber;

                TraverseHandleBOM(productInfo.graphName, repeatTimes);

                // Traverse the AllWorksToDo and write to file.
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"./AllWorksToDo.txt"))
                {
                    foreach (var work in AllWorksToDo)
                    {
                        Console.WriteLine(work.Key + " " + work.Value);
                        file.WriteLine(work.Key + " " + work.Value);
                    }
                }

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"./Works.txt", true /*append*/))
                {
                    file.WriteLine($"====正在处理产品: {productInfo.graphName}, 这个订单需要{productInfo.graphNumber}个====");
                    foreach (var work in WorksToDoOfSingleProduct)
                    {
                        file.WriteLine(work.Key + " " + work.Value);
                    }
                }

            }
        }
    }
}
