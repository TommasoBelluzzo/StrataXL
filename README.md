# StrataXL

StrataXL is an `Excel` integration of [OpenGamma Strata](http://strata.opengamma.io/), an open source library specialized in financial derivatives and market risk analysis. The technical implementation is achieved by using two nested virtual machines, through the following steps:

* OpenGamma Strata is converted from Java to .NET with [IKVM](https://www.ikvm.net/).
* An instance of the .NET Common Language Runtime is created within the Excel environment.
* An hybrid helper class written part in C# and part in VBA is used for handling OpenGamma Strata classes and method invocations through late binding.

## Requirements

StrataXL is platform-agnostic, both x86 and x64 environments are supported. The projects target Visual Studio 2017.

 - Java SE Runtime Environment 8u40 or later for running OpenGamma Strata.
 - .NET Framework 4.0 or greater (CLR v4.0.30319).
 - Excel 2010 or later (VBA 7).

