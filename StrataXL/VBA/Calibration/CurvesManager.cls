VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurvesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'====================================='
' Copyright (C) 2019 Tommaso Belluzzo '
'          Part of StrataXL           '
'====================================='

' SETTINGS

Option Base 0
Option Explicit

' MEMBERS

Dim m_CalibrationResult As Variant
Dim m_Curves() As String
Dim m_GroupRates As Variant
Dim m_Initialized As Boolean
Dim m_Host As RuntimeHost
Dim m_MarketData As Variant
Dim m_SingleCurrency As Boolean
Dim m_ValuationDate As Variant
Dim m_ValuationDatePlain As Date
Dim m_Worksheet As Worksheet

' CONSTRUCTOR

Private Sub Class_Initialize()

    Call PerformCleanup

End Sub

' DESTRUCTOR

Private Sub Class_Terminate()

    Call PerformCleanup

End Sub

' PROPERTY
' Gets the instance of the ImmutableRatesProvider representing the calibration result.

Property Get CalibrationResult() As Variant

    Set CalibrationResult = m_CalibrationResult

End Property

' PROPERTY
' Gets a boolean indicating whether the underlying framework is single-currency or cross-currency.

Property Get SingleCurrency() As Boolean

    SingleCurrency = m_SingleCurrency

End Property

' PROPERTY
' Gets the instance of the LocalData representing the reference valuation date.

Property Get ValuationDate() As Variant

    Set ValuationDate = m_ValuationDate

End Property

' PROPERTY
' Gets the valuation date.

Property Get ValuationDatePlain() As Date

    ValuationDatePlain = m_ValuationDatePlain

End Property

' METHOD
' Parses a discount curve.

Private Sub ParseCurveDiscount(ByVal curveId As Long, ByVal groupRatesBuilder As Variant, ByVal marketDataBuilder As Variant, ByVal vd As Variant, ByVal ccy As Variant, ByVal dcc As Variant, ByVal calendar As Variant, ByVal bda As Variant)

    Dim i As Long: i = (curveId * 4) - 3

    Dim curveIndex As Variant: curveIndex = m_Worksheet.Cells(3, i + 1).Value2
    
    If IsEmpty(curveIndex) Or (VarType(curveIndex) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveIndex = Replace$(UCase$(Trim$(curveIndex)), "-", "_")

    If (curveIndex = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    Set curveIndex = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.OvernightIndices", curveIndex)

    If (m_Host.InvokeMethod(curveIndex, "getCurrency") <> ccy) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " is defined on an index whose currency is different from the referece one.")
    End If

    Dim curveInterpolation As Variant: curveInterpolation = m_Worksheet.Cells(4, i + 1).Value2
    
    If IsEmpty(curveInterpolation) Or (VarType(curveInterpolation) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid interpolation in the cell " & m_Worksheet.Cells(4, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveInterpolation = Replace$(UCase$(Trim$(curveInterpolation)), " ", "_")
    
    If (curveInterpolation = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid interpolation in the cell " & m_Worksheet.Cells(4, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    Set curveInterpolation = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.curve.interpolator.CurveInterpolators", curveInterpolation)

    Dim curveExtrapolation As Variant: curveExtrapolation = m_Worksheet.Cells(5, i + 1).Value2
    
    If IsEmpty(curveExtrapolation) Or (VarType(curveExtrapolation) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid extrapolation in the cell " & m_Worksheet.Cells(5, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveExtrapolation = Replace$(UCase$(Trim$(curveExtrapolation)), " ", "_")
    
    If (curveExtrapolation = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid extrapolation in the cell " & m_Worksheet.Cells(5, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If

    Set curveExtrapolation = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.curve.interpolator.CurveExtrapolators", curveExtrapolation)

    Dim curveDaysCount As Variant: curveDaysCount = m_Worksheet.Cells(6, i + 1).Value2
    
    If IsEmpty(curveDaysCount) Or (VarType(curveDaysCount) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid days count convention in the cell " & m_Worksheet.Cells(6, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveDaysCount = UCase$(Trim$(curveDaysCount))
    
    If (curveDaysCount = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid days count convention in the cell " & m_Worksheet.Cells(6, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    Set curveDaysCount = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DayCount", "of", curveDaysCount)

    Dim rc As Long: rc = m_Worksheet.UsedRange.Rows.Count
    Dim curveNodes As Variant: Set curveNodes = m_Host.CreateObjectFromName("java.util.ArrayList")
    Dim curveNodesDays() As Long: ReDim curveNodesDays(rc - 8)

    Dim j As Long, curveNode As Variant

    For j = 8 To rc
    
        Dim nodeTenor As Variant: nodeTenor = m_Worksheet.Cells(j, i).Value2
        Dim nodeInstrument As Variant: nodeInstrument = m_Worksheet.Cells(j, i + 1).Value2
        Dim nodeQuote As Variant: nodeQuote = m_Worksheet.Cells(j, i + 2).Value2
        Dim nodeSpread As Variant: nodeSpread = m_Worksheet.Cells(j, i + 3).Value2
        
        If IsEmpty(nodeTenor) Or IsEmpty(nodeInstrument) Or IsEmpty(nodeQuote) Or IsEmpty(nodeSpread) Then
            Exit For
        End If
        
        If (VarType(nodeTenor) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid node tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        nodeTenor = UCase$(Trim$(nodeTenor))
            
        If (nodeTenor = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid node tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        If (VarType(nodeInstrument) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid node instrument in the cell " & m_Worksheet.Cells(j, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        nodeInstrument = UCase$(Trim$(nodeInstrument))

        If (nodeInstrument <> "DEP") And (nodeInstrument <> "OIS") Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid node instrument in the cell " & m_Worksheet.Cells(j, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        If Not IsNumeric(nodeQuote) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid node quote in the cell " & m_Worksheet.Cells(j, i + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        If Not IsNumeric(nodeSpread) Or (nodeSpread < 0) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid node spread in the cell " & m_Worksheet.Cells(j, i + 3).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        Dim nodeName As String: nodeName = CStr(ccy) & "-DT-" & nodeTenor
        Dim nodeId As Variant: Set nodeId = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.StandardId", "of", "CALIBRATION", nodeName)
        Dim nodeQuoteId As Variant: Set nodeQuoteId = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.observable.QuoteId", "of", nodeId)

        If (nodeTenor = "ON") Then
        
            If (j <> 8) Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " must define the overnight tenor in the first node.")
            End If
            
            If (nodeInstrument <> "DEP") Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " must associate the overnight tenor to a deposit instrument.")
            End If
            
            curveNodesDays(0) = 0
    
            Dim periodON As Variant: Set periodON = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofDays", 1)
            Dim daON As Variant: Set daON = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DaysAdjustment", "ofBusinessDays", 0, calendar)
            
            Dim conventionBuilderON As Variant: Set conventionBuilderON = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.deposit.type.ImmutableTermDepositConvention", "builder")
            Call m_Host.InvokeMethod(conventionBuilderON, "businessDayAdjustment", bda)
            Call m_Host.InvokeMethod(conventionBuilderON, "currency", ccy)
            Call m_Host.InvokeMethod(conventionBuilderON, "dayCount", dcc)
            Call m_Host.InvokeMethod(conventionBuilderON, "name", "DEP")
            Call m_Host.InvokeMethod(conventionBuilderON, "spotDateOffset", daON)
            Dim conventionON As Variant: Set conventionON = m_Host.InvokeMethod(conventionBuilderON, "build")
    
            Dim templateON As Variant: Set templateON = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.deposit.type.TermDepositTemplate", "of", periodON, conventionON)
            Set curveNode = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.node.TermDepositCurveNode", "of", templateON, nodeQuoteId, nodeSpread)
        
        Else
        
            Dim nodeTenorValue As Variant: nodeTenorValue = Left$(nodeTenor, Len(nodeTenor) - 1)
            
            If Not IsNumeric(nodeTenorValue) Or (InStr(nodeTenorValue, ",") > 0) Or (InStr(nodeTenorValue, ".") > 0) Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
            End If
            
            nodeTenorValue = CLng(nodeTenorValue)
            
            If (nodeTenorValue <= 0) Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " contains an invalid tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
            End If
        
            Dim nodeTenorScale As String: nodeTenorScale = Right$(nodeTenor, 1)
            
            If (nodeTenorScale <> "W") And (nodeTenorScale <> "M") And (nodeTenorScale <> "Y") Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " must define only weekly, monthly and yearly tenors.")
            End If

            Dim nodePeriod As Variant
            Dim nodeDays As Double
            
            If (nodeTenorScale = "W") Then
                Set nodePeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofWeeks", nodeTenorValue)
                nodeDays = nodeTenorValue * CDbl(7.5)
            ElseIf (nodeTenorScale = "M") Then
                Set nodePeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofMonths", nodeTenorValue)
                nodeDays = nodeTenorValue * CDbl(30)
            Else
                Set nodePeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofYears", nodeTenorValue)
                nodeDays = nodeTenorValue * CDbl(360)
            End If
            
            If (nodeDays > CDbl(21600)) Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " must cannot define tenors greater than 60 years.")
            End If

            If (j > 8) And (nodeDays <= curveNodesDays(j - 9)) Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " must be defined on an ascending sequence of tenors, with each tenor greater than the previous one.")
            End If

            curveNodesDays(j - 8) = nodeDays
            
            Set nodeTenor = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "of", nodePeriod)

            Dim daOIS As Variant: Set daOIS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DaysAdjustment", "ofBusinessDays", 2, calendar, bda)
            Dim frequencyOIS As Variant: Set frequencyOIS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "ofYears", 1)
            Dim periodOIS As Variant: Set periodOIS = m_Host.GetFieldStaticByName("java.time.Period", "ZERO")

            Dim conventionBuilderFixedOIS As Variant: Set conventionBuilderFixedOIS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.FixedRateSwapLegConvention", "builder")
            Call m_Host.InvokeMethod(conventionBuilderFixedOIS, "accrualBusinessDayAdjustment", bda)
            Call m_Host.InvokeMethod(conventionBuilderFixedOIS, "accrualFrequency", frequencyOIS)
            Call m_Host.InvokeMethod(conventionBuilderFixedOIS, "currency", ccy)
            Call m_Host.InvokeMethod(conventionBuilderFixedOIS, "dayCount", dcc)
            Call m_Host.InvokeMethod(conventionBuilderFixedOIS, "paymentDateOffset", daOIS)
            Dim conventionFixedOIS As Variant: Set conventionFixedOIS = m_Host.InvokeMethod(conventionBuilderFixedOIS, "build")

            Dim conventionBuilderFloatingOIS As Variant: Set conventionBuilderFloatingOIS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.OvernightRateSwapLegConvention", "builder")
            Call m_Host.InvokeMethod(conventionBuilderFloatingOIS, "accrualBusinessDayAdjustment", bda)
            Call m_Host.InvokeMethod(conventionBuilderFloatingOIS, "index", curveIndex)
            Call m_Host.InvokeMethod(conventionBuilderFloatingOIS, "paymentDateOffset", daOIS)
            Dim conventionFloatingOIS As Variant: Set conventionFloatingOIS = m_Host.InvokeMethod(conventionBuilderFloatingOIS, "build")

            Dim conventionBuilderOIS As Variant: Set conventionBuilderOIS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.ImmutableFixedOvernightSwapConvention", "builder")
            Call m_Host.InvokeMethod(conventionBuilderOIS, "fixedLeg", conventionFixedOIS)
            Call m_Host.InvokeMethod(conventionBuilderOIS, "floatingLeg", conventionFloatingOIS)
            Call m_Host.InvokeMethod(conventionBuilderOIS, "name", "OIS")
            Call m_Host.InvokeMethod(conventionBuilderOIS, "spotDateOffset", daOIS)
            Dim conventionOIS As Variant: Set conventionOIS = m_Host.InvokeMethod(conventionBuilderOIS, "build")
            
            Dim templateOIS As Variant: Set templateOIS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.FixedOvernightSwapTemplate", "of", periodOIS, nodeTenor, conventionOIS)
            Set curveNode = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.node.FixedOvernightSwapCurveNode", "of", templateOIS, nodeQuoteId, nodeSpread)
        
        End If

        Call m_Host.InvokeMethod(curveNodes, "add", curveNode)
        
        Dim nodeQuoteValue As Variant: Set nodeQuoteValue = m_Host.InvokeMethodStaticFromName("java.lang.Double", "valueOf", nodeQuote)
        Call m_Host.InvokeMethod(marketDataBuilder, "addValue", nodeQuoteId, nodeQuoteValue)
    
    Next j
    
    Dim curveNodesCount As Long: curveNodesCount = m_Host.InvokeMethod(curveNodes, "size")

    If (curveNodesCount < 10) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveDiscount", "The curve " & CStr(curveId) & " must contain at least 10 nodes (any empty cell stops the loop).")
    End If
    
    Dim curveName As Variant: Set curveName = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.CurveName", "of", CStr(curveIndex))
    Dim curveValuesX As Variant: Set curveValuesX = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.ValueType", "YEAR_FRACTION")
    Dim curveValuesY As Variant: Set curveValuesY = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.ValueType", "ZERO_RATE")
    
    Dim definitionBuilder As Variant: Set definitionBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.InterpolatedNodalCurveDefinition", "builder")
    Call m_Host.InvokeMethod(definitionBuilder, "dayCount", curveDaysCount)
    Call m_Host.InvokeMethod(definitionBuilder, "interpolator", curveInterpolation)
    Call m_Host.InvokeMethod(definitionBuilder, "extrapolatorLeft", curveExtrapolation)
    Call m_Host.InvokeMethod(definitionBuilder, "extrapolatorRight", curveExtrapolation)
    Call m_Host.InvokeMethod(definitionBuilder, "name", curveName)
    Call m_Host.InvokeMethod(definitionBuilder, "nodes", curveNodes)
    Call m_Host.InvokeMethod(definitionBuilder, "xValueType", curveValuesX)
    Call m_Host.InvokeMethod(definitionBuilder, "yValueType", curveValuesY)
    Dim definition As Variant: Set definition = m_Host.InvokeMethod(definitionBuilder, "build")
    
    Call m_Host.InvokeMethod(groupRatesBuilder, "addCurve", definition, ccy, curveIndex)

End Sub

' METHOD
' Parses a forward curve.

Private Sub ParseCurveForward(ByVal curveId As Long, ByVal groupRatesBuilder As Variant, ByVal marketDataBuilder As Variant, ByVal vd As Variant, ByVal ccy As Variant, ByVal dcc As Variant, ByVal calendar As Variant, ByVal bda As Variant)

    Dim i As Long: i = (curveId * 4) - 3

    Dim curveIndex As Variant: curveIndex = m_Worksheet.Cells(3, i + 1).Value2
    
    If IsEmpty(curveIndex) Or (VarType(curveIndex) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveIndex = Replace$(UCase$(Trim$(curveIndex)), "-", "_")

    If (curveIndex = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    Set curveIndex = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.IborIndices", curveIndex)

    Dim curveIndexTenor As Variant: Set curveIndexTenor = m_Host.InvokeMethod(curveIndex, "getTenor")
    Dim curveIndexTenorString As String: curveIndexTenorString = CStr(curveIndexTenor)
    Dim curveIndexPeriod As Variant: Set curveIndexPeriod = m_Host.InvokeMethod(curveIndexTenor, "getPeriod")
    
    If (m_Host.InvokeMethod(curveIndex, "getCurrency") <> ccy) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " is defined on an index whose currency is different from the referece one.")
    End If

    Dim curveInterpolation As Variant: curveInterpolation = m_Worksheet.Cells(4, i + 1).Value2
    
    If IsEmpty(curveInterpolation) Or (VarType(curveInterpolation) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid interpolation in the cell " & m_Worksheet.Cells(4, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveInterpolation = Replace$(UCase$(Trim$(curveInterpolation)), " ", "_")
    
    If (curveInterpolation = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid interpolation in the cell " & m_Worksheet.Cells(4, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    Set curveInterpolation = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.curve.interpolator.CurveInterpolators", curveInterpolation)

    Dim curveExtrapolation As Variant: curveExtrapolation = m_Worksheet.Cells(5, i + 1).Value2
    
    If IsEmpty(curveExtrapolation) Or (VarType(curveExtrapolation) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid extrapolation in the cell " & m_Worksheet.Cells(5, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveExtrapolation = Replace$(UCase$(Trim$(curveExtrapolation)), " ", "_")
    
    If (curveExtrapolation = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid extrapolation in the cell " & m_Worksheet.Cells(5, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If

    Set curveExtrapolation = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.curve.interpolator.CurveExtrapolators", curveExtrapolation)

    Dim curveDaysCount As Variant: curveDaysCount = m_Worksheet.Cells(6, i + 1).Value2
    
    If IsEmpty(curveDaysCount) Or (VarType(curveDaysCount) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid days count convention in the cell " & m_Worksheet.Cells(6, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    curveDaysCount = UCase$(Trim$(curveDaysCount))
    
    If (curveDaysCount = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid days count convention in the cell " & m_Worksheet.Cells(6, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
    End If
    
    Set curveDaysCount = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DayCount", "of", curveDaysCount)

    Dim rc As Long: rc = m_Worksheet.UsedRange.Rows.Count
    Dim curveNodes As Variant: Set curveNodes = m_Host.CreateObjectFromName("java.util.ArrayList")
    Dim curveNodesDays() As Long: ReDim curveNodesDays(rc - 8)

    Dim j As Long, curveNode As Variant

    For j = 8 To rc
    
        Dim nodeTenor As Variant: nodeTenor = m_Worksheet.Cells(j, i).Value2
        Dim nodeInstrument As Variant: nodeInstrument = m_Worksheet.Cells(j, i + 1).Value2
        Dim nodeQuote As Variant: nodeQuote = m_Worksheet.Cells(j, i + 2).Value2
        Dim nodeSpread As Variant: nodeSpread = m_Worksheet.Cells(j, i + 3).Value2
        
        If IsEmpty(nodeTenor) Or IsEmpty(nodeInstrument) Or IsEmpty(nodeQuote) Or IsEmpty(nodeSpread) Then
            Exit For
        End If
        
        If (VarType(nodeTenor) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid node tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        nodeTenor = UCase$(Trim$(nodeTenor))
            
        If (nodeTenor = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid node tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        If (VarType(nodeInstrument) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid node instrument in the cell " & m_Worksheet.Cells(j, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        nodeInstrument = UCase$(Trim$(nodeInstrument))

        If (nodeInstrument <> "FIX") And (nodeInstrument <> "FRA") And (nodeInstrument <> "FUT") And (nodeInstrument <> "IRS") Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid node instrument in the cell " & m_Worksheet.Cells(j, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        If Not IsNumeric(nodeQuote) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid node quote in the cell " & m_Worksheet.Cells(j, i + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        If Not IsNumeric(nodeSpread) Or (nodeSpread < 0) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid node spread in the cell " & m_Worksheet.Cells(j, i + 3).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        If (j = 8) And ((nodeTenor <> curveIndexTenorString) Or (nodeInstrument <> "FIX")) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must define the first node as a rate fixing whose tenor is equal to the tenor of its index.")
        End If
        
        Dim nodeName As String: nodeName = CStr(ccy) & "-" & curveIndexTenorString & "-" & nodeTenor
        Dim nodeId As Variant: Set nodeId = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.StandardId", "of", "CALIBRATION", nodeName)
        Dim nodeQuoteId As Variant: Set nodeQuoteId = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.observable.QuoteId", "of", nodeId)

        Dim nodeTenorValue As Variant: nodeTenorValue = Left$(nodeTenor, Len(nodeTenor) - 1)
        
        If Not IsNumeric(nodeTenorValue) Or (InStr(nodeTenorValue, ",") > 0) Or (InStr(nodeTenorValue, ".") > 0) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        nodeTenorValue = CLng(nodeTenorValue)
        
        If (nodeTenorValue <= 0) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " contains an invalid tenor in the cell " & m_Worksheet.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
    
        Dim nodeTenorScale As String: nodeTenorScale = Right$(nodeTenor, 1)
        
        If (nodeTenorScale <> "M") And (nodeTenorScale <> "Y") Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must define only monthly and yearly tenors.")
        End If
        
        Dim nodePeriod As Variant
        Dim nodeDays As Double

        If (nodeTenorScale = "M") Then
            Set nodePeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofMonths", nodeTenorValue)
            nodeDays = nodeTenorValue * CDbl(30)
        Else
            Set nodePeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofYears", nodeTenorValue)
            nodeDays = nodeTenorValue * CDbl(360)
        End If
            
        If (nodeDays > CDbl(21600)) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must cannot define tenors greater than 60 years.")
        End If

        If (j > 8) Then
            If (nodeDays <= curveNodesDays(j - 9)) Then
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must be defined on an ascending sequence of tenors, with each tenor greater than the previous one.")
            End If
        End If

        curveNodesDays(j - 8) = nodeDays

        Select Case nodeInstrument
        
            Case "FIX"
            
                Dim templateON As Variant: Set templateON = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.deposit.type.IborFixingDepositTemplate", "of", curveIndex)
                Set curveNode = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.node.IborFixingDepositCurveNode", "of", templateON, nodeQuoteId, nodeSpread)
        
            Case "FRA"

                If (nodeTenorScale <> "M") Then
                    Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must define FRA nodes on monthly tenors.")
                End If

                Dim periodFRA As Variant: Set periodFRA = m_Host.InvokeMethod(nodePeriod, "minus", curveIndexPeriod)
                Dim tenorFRA As Variant: Set tenorFRA = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "of", periodFRA)
                Dim comparisonFRA As Long: comparisonFRA = m_Host.InvokeMethod(tenorFRA, "compareTo", curveIndexTenor)
                
                If (comparisonFRA < 0) Then
                    Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must define FRA nodes on start periods greater than or equal to the tenor of its index.")
                End If
                
                Dim templateFRA As Variant: Set templateFRA = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.fra.type.FraTemplate", "of", periodFRA, curveIndex)
                Set curveNode = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.node.FraCurveNode", "of", templateFRA, nodeQuoteId, nodeSpread)
                
            Case "FUT"

                nodeQuote = nodeQuote / CDbl(100)

                Dim clashActionFUT As Variant: Set clashActionFUT = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.CurveNodeClashAction", "of", "DROP_THIS")
                Dim dateOrderFUT As Variant: Set dateOrderFUT = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.CurveNodeDateOrder", "of", 7, clashActionFUT)
                Dim dateSequenceFUT As Variant: Set dateSequenceFUT = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DateSequence", "of", "MONTHLY-IMM")
                Dim yearMonthDateFUT As Variant: Set yearMonthDateFUT = m_Host.InvokeMethod(m_Host.InvokeMethod(vd, "plus", nodePeriod), "minus", curveIndexPeriod)
                Dim yearMonthFUT As Variant: Set yearMonthFUT = m_Host.InvokeMethodStaticFromName("java.time.YearMonth", "from", yearMonthDateFUT)

                Dim conventionBuilderFUT As Variant: Set conventionBuilderFUT = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.index.type.ImmutableIborFutureConvention", "builder")
                Call m_Host.InvokeMethod(conventionBuilderFUT, "businessDayAdjustment", bda)
                Call m_Host.InvokeMethod(conventionBuilderFUT, "dateSequence", dateSequenceFUT)
                Call m_Host.InvokeMethod(conventionBuilderFUT, "index", curveIndex)
                Call m_Host.InvokeMethod(conventionBuilderFUT, "name", "FUT")
                Dim conventionFUT As Variant: Set conventionFUT = m_Host.InvokeMethod(conventionBuilderFUT, "build")

                Dim templateFUT As Variant: Set templateFUT = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.index.type.IborFutureTemplate", "of", yearMonthFUT, conventionFUT)

                Dim curveNodeBuilderFUT As Variant: Set curveNodeBuilderFUT = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.node.IborFutureCurveNode", "builder")
                Call m_Host.InvokeMethod(curveNodeBuilderFUT, "additionalSpread", nodeSpread)
                Call m_Host.InvokeMethod(curveNodeBuilderFUT, "dateOrder", dateOrderFUT)
                Call m_Host.InvokeMethod(curveNodeBuilderFUT, "label", nodeTenor)
                Call m_Host.InvokeMethod(curveNodeBuilderFUT, "rateId", nodeQuoteId)
                Call m_Host.InvokeMethod(curveNodeBuilderFUT, "template", templateFUT)
                Set curveNode = m_Host.InvokeMethod(curveNodeBuilderFUT, "build")
                
            Case "IRS"

                Set nodeTenor = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "of", nodePeriod)

                Dim daIRS As Variant: Set daIRS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DaysAdjustment", "ofBusinessDays", 2, calendar, bda)
                Dim frequencyIRS As Variant: Set frequencyIRS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "ofYears", 1)
                Dim periodIRS As Variant: Set periodIRS = m_Host.GetFieldStaticByName("java.time.Period", "ZERO")

                Dim conventionBuilderFixedIRS As Variant: Set conventionBuilderFixedIRS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.FixedRateSwapLegConvention", "builder")
                Call m_Host.InvokeMethod(conventionBuilderFixedIRS, "accrualBusinessDayAdjustment", bda)
                Call m_Host.InvokeMethod(conventionBuilderFixedIRS, "accrualFrequency", frequencyIRS)
                Call m_Host.InvokeMethod(conventionBuilderFixedIRS, "currency", ccy)
                Call m_Host.InvokeMethod(conventionBuilderFixedIRS, "dayCount", dcc)
                Call m_Host.InvokeMethod(conventionBuilderFixedIRS, "paymentDateOffset", daIRS)
                Dim conventionFixedIRS As Variant: Set conventionFixedIRS = m_Host.InvokeMethod(conventionBuilderFixedIRS, "build")

                Dim conventionBuilderFloatingIRS As Variant: Set conventionBuilderFloatingIRS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.IborRateSwapLegConvention", "builder")
                Call m_Host.InvokeMethod(conventionBuilderFloatingIRS, "accrualBusinessDayAdjustment", bda)
                Call m_Host.InvokeMethod(conventionBuilderFloatingIRS, "index", curveIndex)
                Call m_Host.InvokeMethod(conventionBuilderFloatingIRS, "paymentDateOffset", daIRS)
                Dim conventionFloatingIRS As Variant: Set conventionFloatingIRS = m_Host.InvokeMethod(conventionBuilderFloatingIRS, "build")

                Dim conventionBuilderIRS As Variant: Set conventionBuilderIRS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.ImmutableFixedIborSwapConvention", "builder")
                Call m_Host.InvokeMethod(conventionBuilderIRS, "fixedLeg", conventionFixedIRS)
                Call m_Host.InvokeMethod(conventionBuilderIRS, "floatingLeg", conventionFloatingIRS)
                Call m_Host.InvokeMethod(conventionBuilderIRS, "name", "IRS")
                Call m_Host.InvokeMethod(conventionBuilderIRS, "spotDateOffset", daIRS)
                Dim conventionIRS As Variant: Set conventionIRS = m_Host.InvokeMethod(conventionBuilderIRS, "build")
                
                Dim templateIRS As Variant: Set templateIRS = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.swap.type.FixedIborSwapTemplate", "of", periodIRS, nodeTenor, conventionIRS)
                Set curveNode = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.node.FixedIborSwapCurveNode", "of", templateIRS, nodeQuoteId, nodeSpread)

        End Select

        Call m_Host.InvokeMethod(curveNodes, "add", curveNode)
        
        Dim nodeQuoteValue As Variant: Set nodeQuoteValue = m_Host.InvokeMethodStaticFromName("java.lang.Double", "valueOf", nodeQuote)
        Call m_Host.InvokeMethod(marketDataBuilder, "addValue", nodeQuoteId, nodeQuoteValue)
    
    Next j
    
    Dim curveNodesCount As Long: curveNodesCount = m_Host.InvokeMethod(curveNodes, "size")

    If (curveNodesCount < 10) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurveForward", "The curve " & CStr(curveId) & " must contain at least 10 nodes (any empty cell stops the loop).")
    End If

    Dim curveName As Variant: Set curveName = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.CurveName", "of", CStr(curveIndex))
    Dim curveValuesX As Variant: Set curveValuesX = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.ValueType", "YEAR_FRACTION")
    Dim curveValuesY As Variant: Set curveValuesY = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.ValueType", "ZERO_RATE")
    
    Dim definitionBuilder As Variant: Set definitionBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.InterpolatedNodalCurveDefinition", "builder")
    Call m_Host.InvokeMethod(definitionBuilder, "dayCount", curveDaysCount)
    Call m_Host.InvokeMethod(definitionBuilder, "interpolator", curveInterpolation)
    Call m_Host.InvokeMethod(definitionBuilder, "extrapolatorLeft", curveExtrapolation)
    Call m_Host.InvokeMethod(definitionBuilder, "extrapolatorRight", curveExtrapolation)
    Call m_Host.InvokeMethod(definitionBuilder, "name", curveName)
    Call m_Host.InvokeMethod(definitionBuilder, "nodes", curveNodes)
    Call m_Host.InvokeMethod(definitionBuilder, "xValueType", curveValuesX)
    Call m_Host.InvokeMethod(definitionBuilder, "yValueType", curveValuesY)
    Dim definition As Variant: Set definition = m_Host.InvokeMethod(definitionBuilder, "build")

    Call m_Host.InvokeMethod(groupRatesBuilder, "addForwardCurve", definition, curveIndex)

End Sub

' METHOD
' Parses a set of cross-currency curves.

Private Sub ParseCurvesCrossCurrency()

    Dim form As New ReferenceDataCrossCurrency: form.Show
    Dim vd As Variant: vd = form.ReferenceValuationDate
    Dim ccyBase As Variant: ccyBase = form.ReferenceCurrencyBase
    Dim ccyCounter As Variant: ccyCounter = form.ReferenceCurrencyCounter
    Dim bdc As Variant: bdc = form.ReferenceBusinessDays
    Dim dcc As Variant: dcc = form.ReferenceDaysCount
    Call Unload(form)
    
    Set vd = m_Host.InvokeMethodStaticFromName("java.time.LocalDate", "of", Year(vd), Month(vd), Day(vd))
    Set ccyBase = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccyBase)
    Set ccyCounter = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccyCounter)
    Set bdc = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayConvention", "of", Replace$(bdc, " ", ""))
    Set dcc = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DayCount", "of", dcc)
    
    Dim calendarBase As Variant: Set calendarBase = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.HolidayCalendarId", "defaultByCurrency", ccyBase)
    Dim calendarCounter As Variant: Set calendarCounter = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.HolidayCalendarId", "defaultByCurrency", ccyCounter)
    Dim bdaBase As Variant: Set bdaBase = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayAdjustment", "of", bdc, calendarBase)
    Dim bdaCounter As Variant: Set bdaCounter = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayAdjustment", "of", bdc, calendarCounter)

    Dim groupName As Variant: Set groupName = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.CurveGroupName", "of", CStr(ccyBase) & "/" & CStr(ccyCounter) & "-CALIBRATED")
    Dim groupRatesBuilder As Variant: Set groupRatesBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.RatesCurveGroupDefinition", "builder")
    Call m_Host.InvokeMethod(groupRatesBuilder, "name", groupName)
    
    Dim marketDataBuilder As Variant: Set marketDataBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.data.ImmutableMarketData", "builder", vd)
    
    Dim cc As Long: cc = m_Worksheet.UsedRange.Columns.Count
    Dim curvesCount As Long: curvesCount = cc / 4
    Dim curves() As String: ReDim curves(curvesCount - 1)

    Dim curveId As Long: curveId = 1
    Dim discountCurveBaseParsed As Boolean: discountCurveBaseParsed = False
    Dim discountCurveCounterParsed As Boolean: discountCurveCounterParsed = False
    
    Dim i As Long
    
    For i = 1 To m_Worksheet.UsedRange.Columns.Count - 3 Step 4
        
        Dim curveType As Variant: curveType = m_Worksheet.Cells(2, i + 1).Value2
        
        If IsEmpty(curveType) Or (VarType(curveType) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The calibration sheet contains an invalid curve type in the cell " & m_Worksheet.Cells(2, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        curveType = UCase$(Trim$(curveType))
        
        Dim curveIndex As Variant: curveIndex = m_Worksheet.Cells(3, i + 1).Value2
        
        If IsEmpty(curveIndex) Or (VarType(curveIndex) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The calibration sheet contains an invalid curve index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        curveIndex = UCase$(Trim$(curveIndex))

        If (curveIndex = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The calibration sheet contains an invalid curve index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Select Case curveType
        
            Case "DISCOUNT"
            
                If discountCurveBaseParsed And discountCurveCounterParsed Then
                    Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The calibration sheet must contain only two discount curves, one for each referece currency.")
                End If

                Dim curveIndexDiscount As Variant: Set curveIndexDiscount = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.OvernightIndices", Replace$(curveIndex, "-", "_"))
                Dim curveCurrencyDiscount As Variant: Set curveCurrencyDiscount = m_Host.InvokeMethod(curveIndexDiscount, "getCurrency")
                
                If (curveCurrencyDiscount <> ccyBase) And (curveCurrencyDiscount <> ccyCounter) Then
                    Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The curve " & CStr(curveId) & " is defined on an index whose currency is different from the referece ones.")
                End If
                
                If (curveCurrencyDiscount = ccyBase) Then
                    Call ParseCurveDiscount(curveId, groupRatesBuilder, marketDataBuilder, vd, ccyBase, dcc, calendarBase, bdaBase)
                    discountCurveBaseParsed = True
                Else
                    Call ParseCurveDiscount(curveId, groupRatesBuilder, marketDataBuilder, vd, ccyCounter, dcc, calendarCounter, bdaCounter)
                    discountCurveCounterParsed = True
                End If
            
            Case "FORWARD"
            
                Dim curveIndexForward As Variant: Set curveIndexForward = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.IborIndices", Replace$(curveIndex, "-", "_"))
                Dim curveCurrencyForward As Variant: Set curveCurrencyForward = m_Host.InvokeMethod(curveIndexForward, "getCurrency")
                
                If (curveCurrencyForward <> ccyBase) And (curveCurrencyForward <> ccyCounter) Then
                    Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The curve " & CStr(curveId) & " is defined on an index whose currency is different from the referece ones.")
                End If
            
                If (curveCurrencyForward = ccyBase) Then
                    Call ParseCurveForward(curveId, groupRatesBuilder, marketDataBuilder, vd, ccyBase, dcc, calendarBase, bdaBase)
                Else
                    Call ParseCurveForward(curveId, groupRatesBuilder, marketDataBuilder, vd, ccyCounter, dcc, calendarCounter, bdaCounter)
                End If
                
            Case Else
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The calibration sheet contains an invalid curve type in the cell " & m_Worksheet.Cells(2, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        
        End Select
        
        curves(curveId - 1) = curveIndex
        curveId = curveId + 1
        
    Next i
    
    If Not discountCurveBaseParsed Or Not discountCurveCounterParsed Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesCrossCurrency", "The calibration sheet must contain a discount curve for each referece currency.")
    End If
    
    m_Curves = curves
    Set m_GroupRates = m_Host.InvokeMethod(groupRatesBuilder, "build")
    Set m_MarketData = m_Host.InvokeMethod(marketDataBuilder, "build")

End Sub

' METHOD
' Parses a set of single-currency curves.

Private Sub ParseCurvesSingleCurrency()

    Dim form As New ReferenceDataSingleCurrency: form.Show
    Dim vd As Variant: vd = form.ReferenceValuationDate
    Dim ccy As Variant: ccy = form.ReferenceCurrency
    Dim bdc As Variant: bdc = form.ReferenceBusinessDays
    Dim dcc As Variant: dcc = form.ReferenceDaysCount
    Call Unload(form)
    
    Set vd = m_Host.InvokeMethodStaticFromName("java.time.LocalDate", "of", Year(vd), Month(vd), Day(vd))
    Set ccy = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccy)
    Set bdc = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayConvention", "of", Replace$(bdc, " ", ""))
    Set dcc = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DayCount", "of", dcc)

    Dim calendar As Variant: Set calendar = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.HolidayCalendarId", "defaultByCurrency", ccy)
    Dim bda As Variant: Set bda = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayAdjustment", "of", bdc, calendar)
    
    Dim groupName As Variant: Set groupName = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.CurveGroupName", "of", CStr(ccy) & "-CALIBRATED")
    Dim groupRatesBuilder As Variant: Set groupRatesBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.RatesCurveGroupDefinition", "builder")
    Call m_Host.InvokeMethod(groupRatesBuilder, "name", groupName)

    Dim marketDataBuilder As Variant: Set marketDataBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.data.ImmutableMarketData", "builder", vd)

    Dim cc As Long: cc = m_Worksheet.UsedRange.Columns.Count
    Dim curvesCount As Long: curvesCount = cc / 4
    Dim curves() As String: ReDim curves(curvesCount - 1)

    Dim curveId As Long: curveId = 1
    Dim discountCurveParsed As Boolean: discountCurveParsed = False
    
    Dim i As Long
    
    For i = 1 To m_Worksheet.UsedRange.Columns.Count - 3 Step 4
        
        Dim curveType As Variant: curveType = m_Worksheet.Cells(2, i + 1).Value2
        
        If IsEmpty(curveType) Or (VarType(curveType) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesSingleCurrency", "The calibration sheet contains an invalid curve type in the cell " & m_Worksheet.Cells(2, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        curveType = UCase$(Trim$(curveType))
        
        Dim curveIndex As Variant: curveIndex = m_Worksheet.Cells(3, i + 1).Value2
        
        If IsEmpty(curveIndex) Or (VarType(curveIndex) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesSingleCurrency", "The calibration sheet contains an invalid curve index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        curveIndex = UCase$(Trim$(curveIndex))
        
        If (curveIndex = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesSingleCurrency", "The calibration sheet contains an invalid curve index in the cell " & m_Worksheet.Cells(3, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Select Case curveType
        
            Case "DISCOUNT"

                If discountCurveParsed Then
                    Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesSingleCurrency", "The calibration sheet must contain only one discount curve.")
                End If
    
                Call ParseCurveDiscount(curveId, groupRatesBuilder, marketDataBuilder, vd, ccy, dcc, calendar, bda)
                
                discountCurveParsed = True
            
            Case "FORWARD"
                Call ParseCurveForward(curveId, groupRatesBuilder, marketDataBuilder, vd, ccy, dcc, calendar, bda)
                
            Case Else
                Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesSingleCurrency", "The calibration sheet contains an invalid curve type in the cell " & m_Worksheet.Cells(2, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        
        End Select
        
        curves(curveId - 1) = curveIndex
        curveId = curveId + 1
        
    Next i
    
    If Not discountCurveParsed Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.ParseCurvesSingleCurrency", "The calibration sheet must contain a discount curve.")
    End If
    
    m_Curves = curves
    Set m_GroupRates = m_Host.InvokeMethod(groupRatesBuilder, "build")
    Set m_MarketData = m_Host.InvokeMethod(marketDataBuilder, "build")

End Sub

' METHOD
' Performs a cleaup of the instance members.

Private Sub PerformCleanup()

    Set m_CalibrationResult = Nothing
    Erase m_Curves
    Set m_GroupRates = Nothing
    Set m_Host = Nothing
    Set m_MarketData = Nothing
    m_SingleCurrency = Empty
    Set m_ValuationDate = Nothing
    m_ValuationDatePlain = Empty
    Set m_Worksheet = Nothing

    m_Initialized = False
    
End Sub

' METHOD
' Initializes an instance of the CurvesManager class.

Public Sub CalibrateCurves()

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.CalibrateCurves", "This instance has not been initialized yet.")
    End If

    Dim calibrator As Variant: Set calibrator = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.pricer.curve.RatesCurveCalibrator", "of", 0.000000001, 0.000000001, 100)
    Dim rd As Variant: Set rd = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.ReferenceData", "standard")
    
    Set m_CalibrationResult = m_Host.InvokeMethod(calibrator, "calibrate", m_GroupRates, m_MarketData, rd)
    
End Sub

' METHOD
' Initializes an instance of the CurvesManager class.

Public Sub Initialize(ByVal host As RuntimeHost, ByVal ws As Worksheet, ByVal SingleCurrency As Boolean)

    If m_Initialized Then
        Exit Sub
    End If

    If host Is Nothing Or Not host.Initialized Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.Initialize", "A valid Runtime Host instance must be provided.")
    End If

    If ws Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.Initialize", "A valid calibration sheet must be provided.")
    End If
    
    Dim rc As Long: rc = ws.UsedRange.Rows.Count
    
    If (rc <= 7) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.Initialize", "The calibration sheet contains an invalid number of rows (" & CStr(rc) & ").")
    End If
    
    Dim cc As Long: cc = ws.UsedRange.Columns.Count
    
    If ((cc Mod 4) <> 0) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.Initialize", "The calibration sheet contains an invalid number of columns (" & CStr(cc) & ").")
    End If

    If (cc = 0) Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.Initialize", "The calibration sheet contains no curves.")
    End If
    
    On Error GoTo ErrorHandler

    Set m_Host = host
    Set m_Worksheet = ws
    m_SingleCurrency = SingleCurrency
    
    If m_SingleCurrency Then
        Call ParseCurvesSingleCurrency
    Else
        Call ParseCurvesCrossCurrency
    End If
    
    m_Initialized = True

    Exit Sub

ErrorHandler:

    Call PerformCleanup
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    
End Sub

' METHOD
' Prepares the sheets containing the calibration result.

Public Sub PrepareResultsSheet()

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.PrepareResultSheets", "This instance has not been initialized yet.")
    End If
    
    If m_CalibrationResult Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "CurvesManager.PrepareResultSheets", "The curves have not been calibrated yet.")
    End If
    
    Dim wsResultsName As String
    
    If m_SingleCurrency Then
        wsResultsName = "Single-Currency Result"
    Else
        wsResultsName = "Cross-Currency Result"
    End If

    On Error Resume Next

    Dim wsResults As Worksheet: Set wsResults = ThisWorkbook.Sheets(wsResultsName)
    
    On Error GoTo 0

    If Not wsResults Is Nothing Then
        
        Dim alerts As Boolean: alerts = Application.DisplayAlerts
    
        Application.DisplayAlerts = False
        wsResults.Delete
        Application.DisplayAlerts = alerts

    End If

    Set wsResults = ThisWorkbook.Sheets.Add(After:=m_Worksheet)
    
    With wsResults.Cells
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Interior.Color = 15132391
        .Style.HorizontalAlignment = xlHAlignCenter
        .Style.VerticalAlignment = xlCenter
    End With
    
    wsResults.Name = wsResultsName
    wsResults.Tab.Color = 11573124

    Dim curvesMap As Variant: Set curvesMap = m_Host.InvokeMethod(m_CalibrationResult, "getCurves")
    Dim curvesMapValues As Variant: Set curvesMapValues = m_Host.InvokeMethod(curvesMap, "values")
    Dim curvesArray As Variant: curvesArray = m_Host.InvokeMethod(curvesMapValues, "toArray")

    Dim offsetLocal As Long: offsetLocal = 1
    Dim offsetExternal As Long: offsetExternal = 1
    Dim i As Long, j As Long
    
    For i = 0 To UBound(m_Curves)

        Dim curve As String: curve = m_Curves(i)
        
        Dim zeroRates As Variant
        
        For j = 0 To UBound(m_Curves)

            Dim curveInner As Variant: Set curveInner = curvesArray(j)
            Dim curveMetadata As Variant: Set curveMetadata = m_Host.InvokeMethod(curveInner, "getMetadata")
            Dim curveName As String: curveName = CStr(m_Host.InvokeMethod(curveMetadata, "getCurveName"))
            
            If (curveName = curve) Then
    
                Dim yValues As Variant: Set yValues = m_Host.InvokeMethod(curveInner, "getYValues")
                zeroRates = m_Host.InvokeMethod(yValues, "toArray")
                
                Exit For
                
            End If

        Next j
        
        wsResults.Columns(offsetLocal).ColumnWidth = 10
        wsResults.Columns(offsetLocal + 1).ColumnWidth = 17
        wsResults.Columns(offsetLocal + 2).ColumnWidth = 17

        With Application.Union(wsResults.Cells(1, offsetLocal), wsResults.Cells(1, offsetLocal + 1), wsResults.Cells(1, offsetLocal + 2))
            .Merge
            .Value2 = curve
            .Borders.Color = 0
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Bold = True
            .Font.Size = 13
            .Interior.Color = 11573124
        End With
        
        With Application.Union(wsResults.Cells(2, offsetLocal), wsResults.Cells(2, offsetLocal + 1), wsResults.Cells(2, offsetLocal + 2))
            .Borders.Color = 0
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = 13285804
        End With
        
        wsResults.Cells(2, offsetLocal).Value2 = "Tenor"
        wsResults.Cells(2, offsetLocal + 1).Value2 = "Market Quote"
        wsResults.Cells(2, offsetLocal + 2).Value2 = "Zero Coupon Rate"

        For j = 0 To UBound(zeroRates)
        
            Dim tenor As Variant: tenor = UCase$(Trim$(m_Worksheet.Cells(j + 8, offsetExternal).Value2))
            Dim instrument As Variant: instrument = UCase$(Trim$(m_Worksheet.Cells(j + 8, offsetExternal + 1).Value2))
            Dim marketQuote As Variant: marketQuote = m_Worksheet.Cells(j + 8, offsetExternal + 2).Value2
            Dim spread As Variant: spread = m_Worksheet.Cells(j + 8, offsetExternal + 3).Value2
        
            If IsEmpty(tenor) Or IsEmpty(instrument) Or IsEmpty(marketQuote) Or IsEmpty(spread) Then
                Exit For
            End If
            
            Dim marketQuoteTotal As Double
            
            If (instrument = "FUT") Then
                marketQuoteTotal = (CDbl(1) - (marketQuote / CDbl(100))) + spread
            Else
                marketQuoteTotal = marketQuote + spread
            End If

            With wsResults.Cells(j + 3, offsetLocal)
                .Value2 = tenor
                .NumberFormat = "@"
                .Borders.Color = 0
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Interior.Color = 16777215
            End With
            
            With wsResults.Cells(j + 3, offsetLocal + 1)
                .Value2 = marketQuoteTotal
                .NumberFormat = "0.0000000000"
                .Borders.Color = 0
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Interior.Color = 16777215
            End With

            With wsResults.Cells(j + 3, offsetLocal + 2)
                .Value2 = Round(zeroRates(j), 10)
                .NumberFormat = "0.0000000000"
                .Borders.Color = 0
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Interior.Color = 16777215
            End With

        Next j
        
        offsetLocal = offsetLocal + 3
        offsetExternal = offsetExternal + 4
    
    Next i

    m_Worksheet.Activate

End Sub
