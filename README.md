# StrataXL

StrataXL is an `Excel` integration of [OpenGamma Strata](http://strata.opengamma.io/), an open source library specialized in financial derivatives and market risk analysis. The technical implementation is achieved by using two nested virtual machines, through the following steps:

* The library is converted from `Java` to `.NET` with [IKVM](https://www.ikvm.net/).
* An instance of the `.NET Common Language Runtime` is created within the `Excel` environment.
* An helper class written in `VBA` is used for managing `OpenGamma Strata` classes and method invocations.
