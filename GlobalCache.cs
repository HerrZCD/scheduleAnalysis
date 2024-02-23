using iSchedule.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using GraphID = System.String;
using GraphRepeatTimes = System.Int32;

namespace Analysis
{
    internal class GlobalCache
    {
        public static Dictionary<GraphID, GraphRepeatTimes> AllWorksToDo = new Dictionary<GraphID, GraphRepeatTimes>();
        public static List<BomData> BOMList = new List<BomData>();
        public static List<AssemblyReader.Result> RequiredProducts = new List<AssemblyReader.Result>();
        public static void SetBOMList(List<BomData> list)
        {
            BOMList = list;
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

        // We check an id, if this id do not exist in the BOMList, we should stop.
        // We record all the works from the topest bom to the most fundamental bom.
        public static void TraverseHandleBOM(string id, int number)
        {
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
                return;
            } else
            {
                 // Record the work.
                if (AllWorksToDo.ContainsKey(bom.BOMDataId))
                {
                    AllWorksToDo[bom.BOMDataId] += number;
                }
                else
                {
                    AllWorksToDo.Add(bom.BOMDataId, number);
                }

                // Go further.
                foreach (var subBom in bom.MaterialInfo.Input)
                {
                    if (id == subBom.MaterialId)
                    {
                        // This is a loop, we should stop here.
                        Console.WriteLine("Loop exist in: " + id);
                        return;
                    }
                    TraverseHandleBOM(subBom.MaterialId, subBom.Number * number);
                }
            }
        }

        public static void WriteAllWorksToFile()
        {
            foreach (var productInfo in GlobalCache.RequiredProducts)
            {
                int repeatTimes = productInfo.graphNumber;

                TraverseHandleBOM(productInfo.graphName, repeatTimes);

                // Traverse the AllWorksToDo and write to file.
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"./AllWorksToDo.txt"))
                {
                    foreach (var work in AllWorksToDo)
                    {
                        file.WriteLine(work.Key + " " + work.Value);
                    }
                }

            }
        }
    }
}
