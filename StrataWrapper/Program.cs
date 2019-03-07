#region Using Directives
using System;
using System.IO;
using System.Reflection;
#endregion

namespace StrataWrapper
{
    public static class Program
    {
        #region Entry Point
        public static Int32 Main()
        {
            Uri assemblyUri = new Uri(Assembly.GetExecutingAssembly().CodeBase);
            String baseDirectory = Path.GetDirectoryName(assemblyUri.LocalPath) ?? String.Empty;

            using (Manager manager = new Manager(baseDirectory))
            {
                Console.WriteLine("# ENVIRONMENT SETUP");
                Console.WriteLine();

                Console.Write("Performing cleanup... ");

                if (manager.PerformCleanup())
                    Console.WriteLine("done.");
                else
                {
                    Console.WriteLine("error.");
                    
                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x01;
                }

                Console.Write("Creating directories... ");

                if (manager.CreateDirectories())
                    Console.WriteLine("done.");
                else
                {
                    Console.WriteLine("error.");
                    
                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x02;
                }

                Console.WriteLine();
                Console.WriteLine("# STRATA LIBRARIES");
                Console.WriteLine();

                Version strataVersion = manager.GetVersionStrata();

                if (strataVersion == null)
                {
                    Console.WriteLine("An error occurred while retrieving the latest available Strata version.");
                    
                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");
                    
                    return 0x10;
                }

                if (strataVersion == Manager.DummyVersion)
                {
                    Console.WriteLine("No available and valid Strata releases have been found on the repository.");
                    
                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");
                    
                    return 0x11;
                }

                Console.WriteLine(FormattableString.Invariant($"The latest available Strata version is '{strataVersion}'."));
                Console.Write("Deploying Strata... ");

                if (manager.DeployStrata())
                    Console.WriteLine("done.");
                else
                {
                    Console.WriteLine("error.");

                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x12;
                }

                Console.WriteLine();
                Console.WriteLine("# IKVM BINARIES");
                Console.WriteLine();

                Version ikvmVersion = manager.GetVersionIKVM();

                if (ikvmVersion == Manager.DummyVersion)
                {
                    Console.WriteLine("No IKVM binaries have been found in the application resources.");

                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x20;
                }

                Console.WriteLine(FormattableString.Invariant($"The latest available IKVM version is '{ikvmVersion}'."));
                Console.Write("Deploying IKVM... ");

                if (manager.DeployIKVM())
                    Console.WriteLine("done.");
                else
                {
                    Console.WriteLine("error.");

                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x21;
                }

                Console.WriteLine();
                Console.WriteLine("# WRAPPERS CREATION");
                Console.WriteLine();

                Console.Write("Creating wrappers... ");

                if (manager.FinalizeWrappers())
                    Console.WriteLine("done.");
                else
                {
                    Console.WriteLine("error.");

                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x30;
                }

                Console.Write("Finalizing wrappers... ");

                if (manager.CreateWrappers())
                    Console.WriteLine("done.");
                else
                {
                    Console.WriteLine("error.");

                    Console.WriteLine();
                    Console.WriteLine("[PROCESS ABORTED]");

                    return 0x31;
                }

                Console.WriteLine();
                Console.WriteLine("[PROCESS COMPLETED]");

                return 0x00;
            }
        }
        #endregion
    }
}