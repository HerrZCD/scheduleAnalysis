using Analysis.Base;
using iSchedule.Base;

namespace Analysis
{
    public class Program
    {
        public static void ReadOrders()
        {
            var result = OrderDeserilizer.DeserilizeOrders("../productList.xlsx");
            if (result == null || result.orders == null)
            {
                Console.WriteLine("Failed to deserilize orders.");
                return;
            }
            foreach (var order in result.orders)
            {
                GlobalCache.AppendProduct(order.ProductType, order.TotalAmount);
            }
        }

        public static void LoadBOMList()
        {
            string file_path = "../BOM1.xlsx";
            var result = BomDeserilizer.DeserilizeBOMList(file_path);

            if (result == null || result.bom_list == null)
            {
                Console.WriteLine("Failed to deserilize BOM list.");
                return;
            }
            GlobalCache.SetBOMList(result.bom_list);
            // GlobalCache.InfoShowBOMList();
        }   
        public async static Task Main(string[] args)
        {
            await Execute(args);
        }
        public static async Task Execute(string[] args)
        {
           await Task.Run(() => LoadBOMList());
           await Task.Run(() => ReadOrders());

            //Console.WriteLine("Data has finished reading, start analyzing...");

            // We want to know all the boms and the repeat times of the boms if we want to produce the products.

            GlobalCache.WriteAllWorksToFile();

            //// We want to know the time the 4 functions take.
            //var watch = System.Diagnostics.Stopwatch.StartNew();

            //var bpList =  await Task.Run(() => AssemblyReader.ReadAssembly("../productList.xlsx", "BP装配", 2, 2, 3,6));
            ////var pbList = await Task.Run(() => AssemblyReader.ReadAssembly("../productList.xlsx", "PB装配", 2, 1, 3, 5));
            ////var fpList = await Task.Run(() => AssemblyReader.ReadAssembly("../productList.xlsx", "FP装配清单", 2, 2, 3, 6));
            ////var cpTRList = await Task.Run(() => AssemblyReader.ReadAssembly("../productList.xlsx", "CP TR装配", 2, 1, 3, 6));

            //GlobalCache.AppendRequiredProducts(bpList);
            ////GlobalCache.AppendRequiredProducts(pbList);
            ////GlobalCache.AppendRequiredProducts(fpList);
            ////GlobalCache.AppendRequiredProducts(cpTRList);

            //Console.WriteLine("Data has finished reading, start analyzing...");

            //// We want to know all the boms and the repeat times of the boms if we want to produce the products.

            //GlobalCache.WriteAllWorksToFile();


            //Console.WriteLine("Writing finished!");
            //watch.Stop();
            //Console.WriteLine($"AssemblyReader.ReadAssembly took {watch.ElapsedMilliseconds} ms");

            //Console.ReadLine();
            Console.ReadLine();


        }
    }
}


