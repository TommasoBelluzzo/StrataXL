# StrataXL

StrataXL is an Excel integration of [OpenGamma Strata](http://strata.opengamma.io/), an open source library specialized in financial derivatives and market risk analysis. The technical implementation is achieved by using two nested virtual machines, through the following steps:

 * OpenGamma Strata is converted from Java to .NET with [IKVM](https://www.ikvm.net/).
 * An instance of the .NET Common Language Runtime is hosted within the Excel environment.
 * An hybrid helper class written part in C# and part in VBA is used for handling OpenGamma Strata classes and method invocations through the late binding approach.

## Requirements

 * StrataXL can be used on every machine equipped with Windows 7 or greater, provided it is capable of fulfilling the requirements listed below. It is platform-agnostic, therefore both x86 and x64 environments are supported.
 * The current release of OpenGamma Strata requires the Java SE Runtime Environment 8u40 or a later release.
 * Any version of the .NET Framework running under the .NET CLR v4.0.30319 is necessary; the minimum required version is 4.0.
 * Any version of Excel supporting VBA 7.0 or greater is necessary; the minimum required version is Excel 2010.
 * The auxiliary projects have been developed under Visual Studio 2017.

## Installation & Upgrade

 1. Open the RuntimeLoader solution in Visual Studio and build the required projects, depending on the bitness of the local Windows and Excel versions. This process compiles the RuntimeLoader libraries and places them into the `\StrataXL\Libraries\` folder. The libraries are used by the VBA class called `RuntimeHost` for creating instances of the .NET CLR.
 1. Sometimes, the native function `SetDefaultDllDirectories`, used by the `RuntimeHost` class, doesn't work properly. This is likely caused by the presence of an antivirus blocking certain operating system calls or protecting the file system. This problem can be bypassed by registering the RuntimeLoader libraries with the regsvr32 utility.
 1. Open the StrataWrapper solution in Visual Studio and build the project. This will produce a console application called StrataWrapper.exe and place it into the `\StrataXL\` folder. The executable, when run:
    * deploys the latest release of IKVM into the `\StrataXL\Libraries\` folder;
    * downloads the latest release of OpenGamma Strata;
    * converts the JAR files of the package into .NET libraies and places the output into the `\StrataXL\Libraries\` folder.
 1. The above step can be performed alone when an upgrade of OpenGamma Strata is required.

## Usage

The spreadsheet `StrataXL-Template.xlsm` can be used as a starting point for creating brand new StrataXL scripts from scratch
