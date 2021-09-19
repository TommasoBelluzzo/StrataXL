# StrataXL

StrataXL is an Excel integration of [OpenGamma Strata](http://strata.opengamma.io/), an open source library specialized in financial derivative instruments and market risk analysis. The technical implementation is achieved by using two nested virtual machines, through the following steps:

 1. OpenGamma Strata is plugged into [IKVM](https://www.ikvm.net/), which allows interoperability between Java and .NET.
 1. An instance of the .NET Common Language Runtime (CLR) is hosted within the Excel environment.
 1. An hybrid helper class written part in C# and part in VBA is used for handling OpenGamma Strata classes and method invocations through the late binding approach.

## Requirements

StrataXL can be used on every machine equipped with Windows 7 or greater; it is platform-agnostic, therefore both x86 and x64 environments are supported. The requirements listed below must be fulfilled:

 * The latest release of OpenGamma Strata requires the Java SE Runtime Environment 8u40 or a later release.
 * Any version of the .NET Framework running the .NET CLR v4.0.30319 is necessary; the minimum required version is 4.0.
 * Any version of Excel supporting VBA 7.0 or greater is necessary; the minimum required version is Excel 2010.
 * VBA projects must be capable of including the following references:
   * Common Language Runtime Execution Engine
   * Microsoft Excel Object Library
   * Microsoft Forms Object Library
   * Microsoft Office Object Library
   * MSCorLib
   * OLE Automation
 * The auxiliary projects have been developed under Visual Studio 2017, but they can be downgraded or upgraded without problems.

## Installation & Upgrade

 1. Download or clone this repository.
 1. Open the RuntimeLoader solution in Visual Studio and build the required projects, depending on the bitness of the local Windows and Excel versions. This process compiles the RuntimeLoader libraries and places them into the `\StrataXL\Libraries\` folder. The libraries are used by the `RuntimeHost` VBA class for creating instances of the .NET CLR.
 1. Sometimes, the native function `SetDefaultDllDirectories`, used by the `RuntimeHost` VBA class, doesn't work properly. This is likely caused by the presence of an antivirus blocking certain operating system calls or protecting the file system. This problem can be bypassed by registering the RuntimeLoader libraries with the `regsvr32` command-line utility.
 1. Open the StrataWrapper solution in Visual Studio and build the project. This will produce a console application called `StrataWrapper.exe` and place it into the `\StrataXL\` folder. The executable, when run:
    * deploys the latest release of IKVM into the `\StrataXL\Libraries\` folder;
    * downloads the latest release of OpenGamma Strata;
    * converts the JAR files of the package into .NET libraies and places the output into the `\StrataXL\Libraries\` folder.
 1. The step above can be performed as a stand-alone procedure when an upgrade of OpenGamma Strata is requested. If an upgrade of IKVM is also necessary, a new ZIP archive with a more recent release can be embedded to the Visual Studio solution.

## Usage

The spreadsheet `StrataXL-Template.xlsm` is a good starting point for creating brand new StrataXL scripts from scratch. StrataXL includes two scripted Excel spreadsheets (market data, included into the spreadsheets, and historical time series, located into the `\StrataXL\Data\` folder, are as of 15th February 2019 and must be manually updated if a later valuation date must be used):

 * `StrataXL-Curves.xlsm`, a multi-curve framework for calibrating cross-currency and single-currency interest rate curves.
 * `StrataXL-Pricing.xlsm`, a tool for pricing and analyzing the cash flows of the following built-in asset classes:
   * Bullet Payments
   * Cross-Currency Swaps
   * Forex Products (Spots, Forwards and Swaps)
   * Forward Rate Agreements
   * Interest Rate Futures
   * Interest Rate Swaps
   * Term Deposits

Debugging exceptions thrown by the underlying .NET CLR can sometimes be tricky. That's why they are logged into a file called `StrataXL.log`, which is located into the `\StrataXL\` folder. Exception messages and stack traces written into the log file can provide useful insights about the errors occurred during the execution of the VBA scripts.
