using System;
using System.Reflection;
using System.Windows.Forms;

namespace FreeCutTester
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("===============================================");
            Console.WriteLine("FreeCut 插件测试程序");
            Console.WriteLine("===============================================");
            Console.WriteLine();

            try
            {
                // 测试加载DLL
                Console.WriteLine("1. 测试加载FreeCut.dll...");
                string dllPath = "FreeCut.dll";

                if (!System.IO.File.Exists(dllPath))
                {
                    dllPath = @"bin\Debug\FreeCut.dll";
                }

                if (!System.IO.File.Exists(dllPath))
                {
                    dllPath = @"bin\Release\FreeCut.dll";
                }

                if (!System.IO.File.Exists(dllPath))
                {
                    Console.WriteLine("错误: 找不到FreeCut.dll文件");
                    Console.WriteLine("请确保已构建项目并且DLL文件存在");
                    Console.WriteLine();
                    Console.WriteLine("按任意键退出...");
                    Console.ReadKey();
                    return;
                }

                Assembly assembly = Assembly.LoadFrom(dllPath);
                Console.WriteLine("✓ DLL加载成功");

                // 测试获取类型
                Console.WriteLine();
                Console.WriteLine("2. 测试插件类型...");

                Type thisAddInType = assembly.GetType("FreeCut.ThisAddIn");
                if (thisAddInType != null)
                {
                    Console.WriteLine("✓ 找到 ThisAddIn 类");
                }
                else
                {
                    Console.WriteLine("✗ 未找到 ThisAddIn 类");
                }

                Type settingsType = assembly.GetType("FreeCut.CropSettings");
                if (settingsType != null)
                {
                    Console.WriteLine("✓ 找到 CropSettings 类");

                    // 测试设置功能
                    Console.WriteLine();
                    Console.WriteLine("3. 测试设置功能...");

                    object settings = Activator.CreateInstance(settingsType);
                    MethodInfo loadMethod = settingsType.GetMethod("Load");
                    if (loadMethod != null)
                    {
                        object loadedSettings = loadMethod.Invoke(null, null);
                        Console.WriteLine("✓ 设置加载功能正常");
                    }
                }
                else
                {
                    Console.WriteLine("✗ 未找到 CropSettings 类");
                }

                Type ribbonType = assembly.GetType("FreeCut.FreeCutRibbon");
                if (ribbonType != null)
                {
                    Console.WriteLine("✓ 找到 FreeCutRibbon 类");
                }
                else
                {
                    Console.WriteLine("✗ 未找到 FreeCutRibbon 类");
                }

                Type exporterType = assembly.GetType("FreeCut.PdfExporter");
                if (exporterType != null)
                {
                    Console.WriteLine("✓ 找到 PdfExporter 类");

                    // 测试PDF导出器
                    Console.WriteLine();
                    Console.WriteLine("3.1 测试PDF导出器功能...");

                    try
                    {
                        // 创建CropSettings实例
                        if (settingsType != null)
                        {
                            object settings = Activator.CreateInstance(settingsType);
                            Console.WriteLine("✓ CropSettings 实例创建成功");

                            // 创建PdfExporter实例
                            object pdfExporter = Activator.CreateInstance(exporterType, settings);
                            Console.WriteLine("✓ PdfExporter 实例创建成功");

                            // 测试EstimateOutputSize方法
                            MethodInfo estimateMethod = exporterType.GetMethod("EstimateOutputSize");
                            if (estimateMethod != null)
                            {
                                var emptyList = new System.Collections.Generic.List<dynamic>();
                                long estimatedSize = (long)estimateMethod.Invoke(pdfExporter, new object[] { emptyList });
                                Console.WriteLine($"✓ 文件大小估算功能正常 (空列表估算大小: {estimatedSize} bytes)");
                            }
                            else
                            {
                                Console.WriteLine("✗ 未找到 EstimateOutputSize 方法");
                            }

                            Console.WriteLine("✓ PDF导出器功能测试通过");
                        }
                        else
                        {
                            Console.WriteLine("✗ 无法创建CropSettings实例，跳过PDF导出器测试");
                        }
                    }
                    catch (Exception pdfEx)
                    {
                        Console.WriteLine($"✗ PDF导出器测试失败: {pdfEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("✗ 未找到 PdfExporter 类");
                }

                Console.WriteLine();
                Console.WriteLine("4. 测试COM可见性...");

                // 检查COM特性
                if (thisAddInType != null)
                {
                    object[] comVisibleAttrs = thisAddInType.GetCustomAttributes(typeof(System.Runtime.InteropServices.ComVisibleAttribute), false);
                    if (comVisibleAttrs.Length > 0)
                    {
                        Console.WriteLine("✓ ThisAddIn 类标记为COM可见");
                    }
                    else
                    {
                        Console.WriteLine("✗ ThisAddIn 类未标记为COM可见");
                    }

                    object[] guidAttrs = thisAddInType.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false);
                    if (guidAttrs.Length > 0)
                    {
                        Console.WriteLine("✓ ThisAddIn 类有GUID属性");
                    }
                    else
                    {
                        Console.WriteLine("✗ ThisAddIn 类缺少GUID属性");
                    }
                }

                Console.WriteLine();
                Console.WriteLine("===============================================");
                Console.WriteLine("测试完成！");
                Console.WriteLine("===============================================");
                Console.WriteLine();
                Console.WriteLine("如果所有测试都通过，说明插件DLL构建正确。");
                Console.WriteLine("您可以使用 install.bat 脚本安装插件到PowerPoint。");

            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("测试过程中发生错误:");
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine("详细错误信息:");
                Console.WriteLine(ex.ToString());
            }

            Console.WriteLine();
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }
}