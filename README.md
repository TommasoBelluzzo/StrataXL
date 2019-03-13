# StrataXL

StrataXL is an `Excel` integration of [OpenGamma Strata](http://strata.opengamma.io/), an open source library specialized in financial derivatives and market risk analysis. The technical implementation is achieved by using two nested virtual machines, through the following steps:

 * OpenGamma Strata is converted from Java to .NET with [IKVM](https://www.ikvm.net/).
 * An instance of the .NET Common Language Runtime is hosted within the Excel environment.
 * An hybrid helper class written part in C# and part in VBA is used for handling OpenGamma Strata classes and method invocations through the late binding approach.

## Requirements

 * StrataXL is platform-agnostic, hence both x86 and x64 environments are supported.
 * The current release of OpenGamma Strata requires the Java SE Runtime Environment 8u40 or a later release.
 * Any version of the .NET Framework running under the CLR v4.0.30319 is necessary; the minimum required version is 4.0.
 * Any version of Excel supporting VBA 7.0 or greater is necessary; the minimum required version is Excel 2010.
 * The auxiliary projects have been developed under Visual Studio 2017.

## Installation & Upgrade


