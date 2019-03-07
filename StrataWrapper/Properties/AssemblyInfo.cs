#region Using Directives
using System;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
#endregion

#region Information
[assembly: AssemblyCompany("Tommaso Belluzzo")]

#if (DEBUG)
[assembly: AssemblyConfiguration("Debug Build")]
#else
[assembly: AssemblyConfiguration("Release Build")]
#endif

[assembly: AssemblyCopyright("Copyright ©2019 Tommaso Belluzzo")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyProduct("StrataXL")]
[assembly: AssemblyTitle("A tool for creating .NET wrappers of the Opengamma Strata library, part of StrataXL.")]
[assembly: AssemblyTrademark("")]
#endregion

#region Settings
[assembly: CLSCompliant(false)]
[assembly: ComVisible(false)]
[assembly: Guid("986A3E9C-F205-42B6-B9A1-A2AAA3084E41")]
[assembly: NeutralResourcesLanguage("en")]
#endregion

#region Version
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0.0.0")]
[assembly: AssemblyVersion("1.0.0.0")]
#endregion