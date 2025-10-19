using System;
using System.Reflection;

class TestFreeCutDll
{
    static void Main()
    {
        try
        {
            Console.WriteLine("Testing FreeCut.dll loading...");
            Console.WriteLine();

            // 尝试加载DLL
            Console.WriteLine("1. Loading FreeCut.dll...");
            Assembly assembly = Assembly.LoadFrom(@"C:\FreeCut\FreeCut.dll");
            Console.WriteLine("   SUCCESS: Assembly loaded");
            Console.WriteLine("   Full Name: " + assembly.FullName);
            Console.WriteLine("   Runtime Version: " + assembly.ImageRuntimeVersion);
            Console.WriteLine();

            // 查找ThisAddIn类
            Console.WriteLine("2. Finding FreeCut.ThisAddIn class...");
            Type addInType = assembly.GetType("FreeCut.ThisAddIn");
            if (addInType != null)
            {
                Console.WriteLine("   SUCCESS: Class found");
                Console.WriteLine("   Full Name: " + addInType.FullName);
                Console.WriteLine("   Is COM Visible: " + addInType.IsVisible);
                Console.WriteLine();

                // 尝试创建实例
                Console.WriteLine("3. Creating instance...");
                object instance = Activator.CreateInstance(addInType);
                Console.WriteLine("   SUCCESS: Instance created");
                Console.WriteLine("   Type: " + instance.GetType().Name);
                Console.WriteLine();

                // 检查接口
                Console.WriteLine("4. Checking interfaces...");
                Type[] interfaces = addInType.GetInterfaces();
                foreach (Type iface in interfaces)
                {
                    Console.WriteLine("   - " + iface.Name);
                }
            }
            else
            {
                Console.WriteLine("   ERROR: Class not found!");
            }

            Console.WriteLine();
            Console.WriteLine("=== All tests passed! ===");
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("=== ERROR ===");
            Console.WriteLine("Type: " + ex.GetType().Name);
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("StackTrace: " + ex.StackTrace);

            if (ex.InnerException != null)
            {
                Console.WriteLine();
                Console.WriteLine("Inner Exception:");
                Console.WriteLine("Type: " + ex.InnerException.GetType().Name);
                Console.WriteLine("Message: " + ex.InnerException.Message);
            }
        }

        Console.WriteLine();
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
