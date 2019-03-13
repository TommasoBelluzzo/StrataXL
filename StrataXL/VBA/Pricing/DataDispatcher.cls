VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'====================================='
' Copyright (C) 2019 Tommaso Belluzzo '
'          Part of StrataXL           '
'====================================='

' SETTINGS

Option Explicit

' MEMBERS

Dim m_BusinessDays As Variant
Dim m_Calendar As Variant
Dim m_Currency As Variant
Dim m_DateFormatter As Variant
Dim m_Host As RuntimeHost
Dim m_Initialized As Boolean
Dim m_RatesProvider As Variant
Dim m_ReferenceData As Variant
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
' Gets the instance of the BusinessDaysConvention representing the reference business days convention.

Property Get BusinessDays() As Variant

    Set BusinessDays = m_BusinessDays

End Property

' PROPERTY
' Gets the instance of the HolidayCalendarId representing the reference calendar.

Property Get Calendar() As Variant

    Set Calendar = m_Calendar

End Property

' PROPERTY
' Gets the instance of the DateFormatter representing the reference date formatter.

Property Get DateFormatter() As Variant

    Set DateFormatter = m_DateFormatter

End Property

' PROPERTY
' Gets the instance of the RatesProvider representing the reference financial data provider.

Property Get RatesProvider() As Variant

    Set RatesProvider = m_RatesProvider

End Property

' PROPERTY
' Gets the instance of the ReferenceData representing the reference financial data.

Property Get ReferenceData() As Variant

    Set ReferenceData = m_ReferenceData

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
' Parses the setup sheet of curves (Setup-Curves).

Private Sub ParseWorksheetCurves(ByVal ws As Worksheet, ByVal ratesProviderBuilder As Variant)

    Dim cc As Long: cc = ws.UsedRange.Columns.Count

    If ((cc Mod 2) <> 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid number of columns (" & CStr(cc) & ").")
    End If
    
    Dim rc As Long: rc = ws.UsedRange.Rows.Count
    
    If (rc <= 6) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains too few rows (" & CStr(rc) & ").")
    End If
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim i As Long, j As Long

    For i = 1 To cc - 1 Step 2
    
        Dim curveType As Variant: curveType = ws.Cells(1, i).Value2
    
        If IsEmpty(curveType) Or (VarType(curveType) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid curve type in the cell " & ws.Cells(1, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        curveType = UCase$(Trim$(curveType))
        
        If (curveType <> "DISCOUNT") And (curveType <> "OVERNIGHT") And (curveType <> "FORWARD") Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid curve type in the cell " & ws.Cells(1, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Dim reference As Variant: reference = ws.Cells(2, i).Value2
        
        If IsEmpty(reference) Or (VarType(reference) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid reference in the cell " & ws.Cells(2, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        reference = UCase$(Trim$(reference))
        
        If (reference = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid reference in the cell " & ws.Cells(2, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Dim interpolation As Variant: interpolation = ws.Cells(3, i).Value2
        
        If IsEmpty(interpolation) Or (VarType(interpolation) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid interpolation in the cell " & ws.Cells(3, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        interpolation = Replace$(UCase$(Trim$(interpolation)), " ", "_")
        
        If (interpolation = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid interpolation in the cell " & ws.Cells(3, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Dim extrapolation As Variant: extrapolation = ws.Cells(4, i).Value2
        
        If IsEmpty(extrapolation) Or (VarType(extrapolation) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid extrapolation in the cell " & ws.Cells(4, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        extrapolation = Replace$(UCase$(Trim$(extrapolation)), " ", "_")
        
        If (extrapolation = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid extrapolation in the cell " & ws.Cells(4, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Dim dcc As Variant: dcc = ws.Cells(5, i).Value2
        
        If IsEmpty(dcc) Or (VarType(dcc) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid days count convention in the cell " & ws.Cells(5, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If
        
        dcc = UCase$(Trim$(dcc))
        
        If (dcc = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid days count convention in the cell " & ws.Cells(5, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
        End If

        Dim times() As Double: ReDim times(rc - 7)
        Dim rates() As Double: ReDim rates(rc - 7)
        Dim entriesCount As Long: entriesCount = 0

        For j = 7 To rc
        
            Dim time As Variant: time = ws.Cells(j, i).Value2
            Dim rate As Variant: rate = ws.Cells(j, i + 1).Value2
            
            If IsEmpty(time) Or IsEmpty(rate) Then
                Exit For
            End If
            
            If Not IsNumeric(time) Or (time < 0) Then
                Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid time in the cell " & ws.Cells(j, i).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
            End If
            
            If Not IsNumeric(rate) Then
                Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" contains an invalid rate in the cell " & ws.Cells(j, i + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ".")
            End If

            times(j - 7) = time
            rates(j - 7) = rate
            entriesCount = entriesCount + 1
        
        Next j
        
        If (entriesCount < 10) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetCurves", "The sheet ""Setup-Curves"" must contain at least 10 observations for " & reference & " (any empty cell stops the loop).")
        End If
        
        Set interpolation = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.curve.interpolator.CurveInterpolators", interpolation)
        Set extrapolation = m_Host.GetPropertyStaticFromName("com.opengamma.strata.market.curve.interpolator.CurveExtrapolators", extrapolation)
        Set dcc = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DayCount", "of", dcc)
        
        ReDim Preserve times(entriesCount - 1)
        Dim timesArray As Variant: Set timesArray = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.collect.array.DoubleArray", "ofUnsafe", times)
        
        ReDim Preserve rates(entriesCount - 1)
        Dim ratesArray As Variant: Set ratesArray = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.collect.array.DoubleArray", "ofUnsafe", rates)
        
        Select Case curveType
        
            Case "DISCOUNT"
            
                Dim discountCcy As Variant: Set discountCcy = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", reference)
                Dim discountFactors As Variant: Set discountFactors = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.Curves", "discountFactors", reference & "-DISCOUNT", dcc)
                Dim discountCurve As Variant: Set discountCurve = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.InterpolatedNodalCurve", "of", discountFactors, timesArray, ratesArray, interpolation, extrapolation, extrapolation)
                Call m_Host.InvokeMethod(ratesProviderBuilder, "discountCurve", discountCcy, discountCurve)
                
            Case "FORWARD"

                Dim forwardIndex As Variant: Set forwardIndex = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.IborIndices", Replace$(reference, "-", "_"))
                Dim forwardRates As Variant: Set forwardRates = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.Curves", "zeroRates", reference, dcc)
                Dim forwardCurve As Variant: Set forwardCurve = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.InterpolatedNodalCurve", "of", forwardRates, timesArray, ratesArray, interpolation, extrapolation, extrapolation)
                Dim forwardTimeSeriesPath As String: forwardTimeSeriesPath = ThisWorkbook.Path & "\Data\" & reference & ".csv"
        
                If Not fso.FileExists(forwardTimeSeriesPath) Then
                    Call m_Host.InvokeMethod(ratesProviderBuilder, "iborIndexCurve", forwardIndex, forwardCurve)
                Else
        
                    Dim forwardLocator As Variant: Set forwardLocator = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.collect.io.ResourceLocator", "of", "file:" & forwardTimeSeriesPath)
                    Dim forwardFixings As Variant: Set forwardFixings = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.loader.csv.FixingSeriesCsvLoader", "load", forwardLocator)
                    Dim forwardFixingsValues As Variant: Set forwardFixingsValues = m_Host.InvokeMethod(forwardFixings, "values")
                    Dim forwardFixingsValuesList As Variant: Set forwardFixingsValuesList = m_Host.InvokeMethod(forwardFixingsValues, "asList")

                    Dim forwardTimeSeries As Variant: Set forwardTimeSeries = m_Host.InvokeMethod(forwardFixingsValuesList, "get", 0)
                    Call m_Host.InvokeMethod(ratesProviderBuilder, "iborIndexCurve", forwardIndex, forwardCurve, forwardTimeSeries)
        
                End If
                
            Case "OVERNIGHT"
            
                Dim overnightIndex As Variant: Set overnightIndex = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.OvernightIndices", Replace$(reference, "-", "_"))
                Dim overnightRates As Variant: Set overnightRates = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.Curves", "zeroRates", reference, dcc)
                Dim overnightCurve As Variant: Set overnightCurve = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.market.curve.InterpolatedNodalCurve", "of", overnightRates, timesArray, ratesArray, interpolation, extrapolation, extrapolation)
                Dim overnightTimeSeriesPath As String: overnightTimeSeriesPath = ThisWorkbook.Path & "\Data\" & reference & ".csv"
        
                If Not fso.FileExists(overnightTimeSeriesPath) Then
                    Call m_Host.InvokeMethod(ratesProviderBuilder, "overnightIndexCurve", overnightIndex, overnightCurve)
                Else
        
                    Dim overnightLocator As Variant: Set overnightLocator = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.collect.io.ResourceLocator", "of", "file:" & overnightTimeSeriesPath)
                    Dim overnightFixings As Variant: Set overnightFixings = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.loader.csv.FixingSeriesCsvLoader", "load", overnightLocator)
                    Dim overnightFixingsValues As Variant: Set overnightFixingsValues = m_Host.InvokeMethod(overnightFixings, "values")
                    Dim overnightFixingsValuesList As Variant: Set overnightFixingsValuesList = m_Host.InvokeMethod(overnightFixingsValues, "asList")
        
                    Dim overnightTimeSeries As Variant: Set overnightTimeSeries = m_Host.InvokeMethod(overnightFixingsValuesList, "get", 0)
                    Call m_Host.InvokeMethod(ratesProviderBuilder, "overnightIndexCurve", overnightIndex, overnightCurve, overnightTimeSeries)
        
                End If
        
        End Select

    Next i
    
End Sub

' METHOD
' Parses the setup sheet of exchange rates (Setup-Exchange).

Private Sub ParseWorksheetExchange(ByVal ws As Worksheet, ByVal ratesProviderBuilder As Variant)

    Dim cs As Long: cs = ws.UsedRange.Columns.Count

    If (cs <> 3) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains an invalid number of columns (" & CStr(cs) & ").")
    End If
    
    Dim rs As Long: rs = ws.UsedRange.Rows.Count
    
    If (rs <= 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains an invalid number of rows (" & CStr(rs) & ").")
    End If

    Dim matrixBuilder As Variant: Set matrixBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.FxMatrix", "builder")

    Dim i As Long

    For i = 2 To rs
    
        Dim ccy1 As Variant: ccy1 = ws.Cells(i, 1).Value2
        
        If IsEmpty(ccy1) Or (VarType(ccy1) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains invalid data in the cell [" & CStr(i) & ", 1].")
        End If

        ccy1 = UCase$(Trim$(ccy1))
        
        If (Len(ccy1) <> 3) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains invalid data in the cell [" & CStr(i) & ", 1].")
        End If

        Dim ccy2 As Variant: ccy2 = ws.Cells(i, 2).Value2
        
        If IsEmpty(ccy2) Or (VarType(ccy2) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains invalid data in the cell [" & CStr(i) & ", 2].")
        End If
        
        ccy2 = UCase$(Trim$(ccy2))
        
        If (Len(ccy2) <> 3) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains invalid data in the cell [" & CStr(i) & ", 2].")
        End If
        
        If (ccy1 = ccy2) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains currency pairs defined on the same currency.")
        End If
        
        Dim rate As Variant: rate = ws.Cells(i, 3).Value2
        
        If IsEmpty(rate) Or Not IsNumeric(rate) Or (rate <= 0) Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.ParseWorksheetExchange", "The sheet ""Setup-Exchange"" contains invalid data in the cell [" & CStr(i) & ", 3].")
        End If

        Set ccy1 = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccy1)
        Set ccy2 = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccy2)
        Call m_Host.InvokeMethod(matrixBuilder, "addRate", ccy1, ccy2, rate)

    Next i

    Dim matrix As Variant: Set matrix = m_Host.InvokeMethod(matrixBuilder, "build")
    Call m_Host.InvokeMethod(ratesProviderBuilder, "fxRateProvider", matrix)

End Sub

' METHOD
' Performs a cleaup of the instance members.

Private Sub PerformCleanup()

    Set m_BusinessDays = Nothing
    Set m_Calendar = Nothing
    Set m_Currency = Nothing
    Set m_DateFormatter = Nothing
    Set m_Host = Nothing
    Set m_RatesProvider = Nothing
    Set m_ReferenceData = Nothing
    Set m_ValuationDate = Nothing
    m_ValuationDatePlain = Empty
    Set m_Worksheet = Nothing
    
    m_Initialized = False
    
End Sub

' METHOD
' Cleans the previous results from the current trades sheet.

Public Sub CleanCurrentTradesSheet(ByVal doubleHeader As Integer, ByVal doubleRows As Boolean)

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.CleanCurrentTradesSheet", "This instance has not been initialized yet.")
    End If
    
    Dim headers As Long
    
    If doubleHeader Then
        headers = 2
    Else
        headers = 1
    End If

    Dim rc As Long: rc = m_Worksheet.UsedRange.Rows.Count

    If (rc < headers) Then
        Call Err.Raise(vbObjectError + 1, "Pricing.CleanCurrentTradesSheet", "The trades sheet contains an invalid number of rows (" & CStr(rc) & ").")
    End If

    If (rc = headers) Then
        Exit Sub
    End If
    
    If doubleRows And (((rc - headers) Mod 2) <> 0) Then
        Call Err.Raise(vbObjectError + 1, "Pricing.CleanCurrentTradesSheet", "The trades sheet contains an invalid number of rows (" & CStr(rc) & ").")
    End If
    
    Dim cc As Long: cc = m_Worksheet.UsedRange.Columns.Count
    Dim i As Long, j As Long
        
    For i = headers + 1 To rc
        For j = cc - 1 To cc
            m_Worksheet.Cells(i, j).ClearContents
        Next j
    Next i

End Sub

' METHOD
' Creates a TradeInfo with the given identifier and trade date.

Public Function CreateTradeInfo(ByVal id As Long, ByVal tradeDate As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.CreateTradeInfo", "This instance has not been initialized yet.")
    End If

    If (id <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.CreateTradeInfo", "The identifier must be positive.")
    End If

    Dim idText As String: idText = CStr(id)
    Dim counterparty As Variant: Set counterparty = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.StandardId", "of", "StrataXL", "COUNTERPARTY" & idText)
    Dim identifier As Variant: Set identifier = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.StandardId", "of", "StrataXL", "ID" & idText)

    Dim tradeInfoBuilder As Variant: Set tradeInfoBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.TradeInfo", "builder")
    Call m_Host.InvokeMethod(tradeInfoBuilder, "counterparty", counterparty)
    Call m_Host.InvokeMethod(tradeInfoBuilder, "id", identifier)

    If Not IsMissing(tradeDate) Then
        Call m_Host.InvokeMethod(tradeInfoBuilder, "tradeDate", tradeDate)
    End If

    Set CreateTradeInfo = m_Host.InvokeMethod(tradeInfoBuilder, "build")

End Function

' METHOD
' Retrieves an amount from the given worksheet cell.

Public Function GetAmount(ByVal wsCell As String) As Double

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetAmount", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetAmount", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetAmount", "A valid cell must be provided.")
    End If
    
    Dim amount As Variant: amount = r.Value2
    
    If IsEmpty(amount) Or Not IsNumeric(amount) Or (amount <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetAmount", "The trades sheet contains an invalid amount in the cell " & wsCell & ".")
    End If
    
    GetAmount = amount

End Function

' METHOD
' Retrieves a BusinessDayConvention from the given worksheet cell.

Public Function GetBusinessDaysConvention(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetBusinessDaysConvention", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetBusinessDaysConvention", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetBusinessDaysConvention", "A valid cell must be provided.")
    End If
    
    Dim bdc As Variant: bdc = r.Value2

    If IsEmpty(bdc) Or (VarType(bdc) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetBusinessDaysConvention", "The trades sheet contains an invalid business days convention in the cell " & wsCell & ".")
    End If
    
    bdc = Replace$(UCase$(Trim$(bdc)), " ", "")
        
    If (bdc = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetBusinessDaysConvention", "The trades sheet contains an invalid business days convention in the cell " & wsCell & ".")
    End If

    Set GetBusinessDaysConvention = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayConvention", "of", bdc)

End Function

' METHOD
' Retrieves a HolidayCalendarId from the given worksheet cell.

Public Function GetCalendar(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCalendar", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCalendar", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCalendar", "A valid cell must be provided.")
    End If
    
    Dim Calendar As Variant: Calendar = r.Value2

    If IsEmpty(Calendar) Or (VarType(Calendar) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCalendar", "The trades sheet contains an invalid calendar in the cell " & wsCell & ".")
    End If
    
    Calendar = UCase$(Trim$(Calendar))
        
    If (Len(Calendar) <> 4) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrency", "The trades sheet contains an invalid calendar in the cell " & wsCell & ".")
    End If

    Set GetCalendar = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.HolidayCalendarId", "of", Calendar)

End Function

' METHOD
' Retrieves a Currency from the given worksheet cell.

Public Function GetCurrency(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrency", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrency", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrency", "A valid cell must be provided.")
    End If
    
    Dim ccy As Variant: ccy = r.Value2

    If IsEmpty(ccy) Or (VarType(ccy) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrency", "The trades sheet contains an invalid currency in the cell " & wsCell & ".")
    End If
    
    ccy = UCase$(Trim$(ccy))
        
    If (Len(ccy) <> 3) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrency", "The trades sheet contains an invalid currency in the cell " & wsCell & ".")
    End If

    Set GetCurrency = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccy)

End Function

' METHOD
' Retrieves a CurrencyAmount from the given worksheet cell.

Public Function GetCurrencyAmount(ByVal wsCells As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrencyAmount", "This instance has not been initialized yet.")
    End If

    wsCells = Trim$(wsCells)

    If (wsCells = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrencyAmount", "Two valid cells must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCells)
    
    If (r.Columns.Count <> 2) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrencyAmount", "Two valid cells must be provided.")
    End If
    
    Dim cellsArray() As String, cell As String

    Dim amount As Variant: amount = r.Item(1).Value2
    
    If IsEmpty(amount) Or Not IsNumeric(amount) Or (amount <= 0) Then
        cellsArray() = Split(wsCells, ":")
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrencyAmount", "The trades sheet contains an invalid amount in the cell " & cellsArray(0) & ".")
    End If
    
    Dim ccy As Variant: ccy = r.Item(2).Value2
    
    If IsEmpty(ccy) Or (VarType(ccy) <> vbString) Then
        cellsArray() = Split(wsCells, ":")
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrencyAmount", "The trades sheet contains an invalid currency in the cell " & cellsArray(1) & ".")
    End If
    
    ccy = UCase$(Trim$(ccy))
        
    If (Len(ccy) <> 3) Then
        cellsArray() = Split(wsCells, ":")
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetCurrencyAmount", "The trades sheet contains an invalid currency in the cell " & cellsArray(1) & ".")
    End If

    Set ccy = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccy)
    Set GetCurrencyAmount = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.CurrencyAmount", "of", ccy, amount)

End Function

' METHOD
' Retrieves a LocalDate from the given worksheet cell.

Public Function GetDate(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDate", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDate", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDate", "A valid cell must be provided.")
    End If
    
    Dim d As Variant: d = r.Value2
    
    If IsEmpty(d) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDate", "The trades sheet contains an invalid date in the cell " & wsCell & ".")
    End If

    Set GetDate = m_Host.InvokeMethodStaticFromName("java.time.LocalDate", "of", Year(d), Month(d), Day(d))

End Function

' METHOD
' Retrieves a DateSequence from the given worksheet cell.

Public Function GetDateSequence(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDateSequence", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDateSequence", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDateSequence", "A valid cell must be provided.")
    End If
    
    Dim dateSequence As Variant: dateSequence = r.Value2
    
    If IsEmpty(dateSequence) Or (VarType(dateSequence) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDateSequence", "The trades sheet contains an invalid date sequence in the cell " & wsCell & ".")
    End If
    
    dateSequence = Replace$(UCase$(Trim$(dateSequence)), " ", "-")

    If (dateSequence = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDateSequence", "The trades sheet contains an invalid date sequence in the cell " & wsCell & ".")
    End If

    Set GetDateSequence = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DateSequence", "of", dateSequence)

End Function

' METHOD
' Retrieves a DayCount from the given worksheet cell.

Public Function GetDaysCountConvention(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDaysCountConvention", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDaysCountConvention", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDaysCountConvention", "A valid cell must be provided.")
    End If
    
    Dim dcc As Variant: dcc = r.Value2

    If IsEmpty(dcc) Or (VarType(dcc) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDaysCountConvention", "The trades sheet contains an invalid days count convention in the cell " & wsCell & ".")
    End If
    
    dcc = UCase$(Trim$(dcc))

    If (dcc = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDaysCountConvention", "The trades sheet contains an invalid days count convention in the cell " & wsCell & ".")
    End If

    Set GetDaysCountConvention = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.DayCount", "of", dcc)

End Function

' METHOD
' Retrieves a BuySell from the given worksheet cell.

Public Function GetDirection(ByVal wsCell As String, ByVal asObject As Boolean) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDirection", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDirection", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDirection", "A valid cell must be provided.")
    End If
    
    Dim direction As Variant: direction = r.Value2

    If IsEmpty(direction) Or (VarType(direction) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDirection", "The trades sheet contains an invalid direction in the cell " & wsCell & ".")
    End If
    
    direction = UCase$(Trim$(direction))

    If (direction <> "BUY") And (direction <> "SELL") Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetDirection", "The trades sheet contains an invalid direction in the cell " & wsCell & ".")
    End If

    If asObject Then
        Set GetDirection = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.common.BuySell", "of", direction)
    Else
        GetDirection = direction
    End If

End Function

' METHOD
' Retrieves a Frequency from the given worksheet cell.

Public Function GetFrequency(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "A valid cell must be provided.")
    End If

    Dim frequency As Variant: frequency = r.Value2
    
    If IsEmpty(frequency) Or (VarType(frequency) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "The trades sheet contains an invalid frequency in the cell " & wsCell & ".")
    End If
    
    frequency = UCase$(Trim$(frequency))

    If (frequency = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "The trades sheet contains an invalid frequency in the cell " & wsCell & ".")
    End If
    
    If (frequency = "TERM") Then
        Set GetFrequency = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "TERM")
        Exit Function
    End If
    
    Dim frequencyValue As Variant: frequencyValue = Left$(frequency, Len(frequency) - 1)
    
    If Not IsNumeric(frequencyValue) Or (InStr(frequencyValue, ",") > 0) Or (InStr(frequencyValue, ".") > 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "The trades sheet contains an invalid frequency in the cell " & wsCell & ".")
    End If
    
    frequencyValue = CLng(frequencyValue)
    
    If (frequencyValue <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "The trades sheet contains an invalid frequency in the cell " & wsCell & ".")
    End If

    Dim frequencyPeriod As String: frequencyPeriod = Right$(frequency, 1)
    
    If (frequencyPeriod <> "D") And (frequencyPeriod <> "W") And (frequencyPeriod <> "M") And (frequencyPeriod <> "Y") Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "The trades sheet contains an invalid frequency in the cell " & wsCell & ".")
    End If
    
    Select Case frequencyPeriod
    
        Case "D"
            Set GetFrequency = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "ofDays", frequencyValue)
            
        Case "W"
            Set GetFrequency = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "ofWeeks", frequencyValue)
        
        Case "M"
            Set GetFrequency = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "ofMonths", frequencyValue)
        
        Case "Y"
            Set GetFrequency = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.schedule.Frequency", "ofYears", frequencyValue)
            
        Case Else
            Set GetFrequency = Nothing

    End Select

End Function

' METHOD
' Retrieves a future maturity from the given worksheet cell.

Public Function GetFutureMaturity(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFutureMaturity", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFutureMaturity", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFutureMaturity", "A valid cell must be provided.")
    End If
    
    Dim futureMaturity As Variant: futureMaturity = r.Value2

    If IsEmpty(futureMaturity) Or (VarType(futureMaturity) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFutureMaturity", "The trades sheet contains an invalid future maturity in the cell " & wsCell & ".")
    End If
    
    futureMaturity = UCase$(Trim$(futureMaturity))

    If (futureMaturity = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFutureMaturity", "The trades sheet contains an invalid future maturity in the cell " & wsCell & ".")
    End If
    
    Dim futureMaturityParts() As String: futureMaturityParts = Split(futureMaturity, "-")
    
    If (UBound(futureMaturityParts) <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFutureMaturity", "The trades sheet contains an invalid future maturity in the cell " & wsCell & ".")
    End If

    Dim m As String: m = futureMaturityParts(0)
    Dim y As String: y = futureMaturityParts(1)
    Dim d As Date: d = DateValue("01 " & m & " " & y)

    Set GetFutureMaturity = m_Host.InvokeMethodStaticFromName("java.time.YearMonth", "of", Year(d), Month(d))

End Function

' METHOD
' Retrieves an IborIndex from the given worksheet cell.

Public Function GetIborIndex(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetIborIndex", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetIborIndex", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetIborIndex", "A valid cell must be provided.")
    End If
    
    Dim index As Variant: index = r.Value2

    If IsEmpty(index) Or (VarType(index) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetIborIndex", "The trades sheet contains an invalid IBOR index in the cell " & wsCell & ".")
    End If
    
    index = Replace$(UCase$(Trim$(index)), "-", "_")

    If (index = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetIborIndex", "The trades sheet contains an invalid IBOR index in the cell " & wsCell & ".")
    End If

    Set GetIborIndex = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.IborIndices", index)

End Function

' METHOD
' Retrieves an OvernightAccrualMethod from the given worksheet cell.

Public Function GetOvernightAccrualMethod(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightAccrualMethod", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightAccrualMethod", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightAccrualMethod", "A valid cell must be provided.")
    End If
    
    Dim accrualMethod As Variant: accrualMethod = r.Value2

    If IsEmpty(accrualMethod) Or (VarType(accrualMethod) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightAccrualMethod", "The trades sheet contains an invalid overnight accrual method in the cell " & wsCell & ".")
    End If
    
    accrualMethod = Replace$(UCase$(Trim$(accrualMethod)), " ", "_")

    If (accrualMethod = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightAccrualMethod", "The trades sheet contains an invalid overnight accrual method in the cell " & wsCell & ".")
    End If

    Set GetOvernightAccrualMethod = m_Host.GetPropertyStaticFromName("com.opengamma.strata.product.swap.OvernightAccrualMethod", accrualMethod)

End Function

' METHOD
' Retrieves an OvernightIndex from the given worksheet cell.

Public Function GetOvernightIndex(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightIndex", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightIndex", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightIndex", "A valid cell must be provided.")
    End If
    
    Dim index As Variant: index = r.Value2

    If IsEmpty(index) Or (VarType(index) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightIndex", "The trades sheet contains an invalid overnight index in the cell " & wsCell & ".")
    End If
    
    index = Replace$(UCase$(Trim$(index)), "-", "_")

    If (index = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetOvernightIndex", "The trades sheet contains an invalid overnight index in the cell " & wsCell & ".")
    End If

    Set GetOvernightIndex = m_Host.GetPropertyStaticFromName("com.opengamma.strata.basics.index.OvernightIndices", index)

End Function

' METHOD
' Retrieves a PayReceive from the given worksheet cell.

Public Function GetPaymentDirection(ByVal wsCell As String, ByVal asObject As Boolean) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPaymentDirection", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPaymentDirection", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPaymentDirection", "A valid cell must be provided.")
    End If
    
    Dim paymentDirection As Variant: paymentDirection = r.Value2

    If IsEmpty(paymentDirection) Or (VarType(paymentDirection) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPaymentDirection", "The trades sheet contains an invalid payment direction in the cell " & wsCell & ".")
    End If
    
    paymentDirection = UCase$(Trim$(paymentDirection))

    If (paymentDirection <> "PAY") And (paymentDirection <> "RECEIVE") Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPaymentDirection", "The trades sheet contains an invalid payment direction in the cell " & wsCell & ".")
    End If
    
    If asObject Then
        Set GetPaymentDirection = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.common.PayReceive", "of", paymentDirection)
    Else
        GetPaymentDirection = paymentDirection
    End If

End Function

' METHOD
' Retrieves a Period from the given worksheet cell.

Public Function GetPeriod(ByVal wsCell As String, Optional ByVal restriction As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "A valid cell must be provided.")
    End If
    
    If (restriction <> vbNullString) Then
        If (restriction <> "D") And (restriction <> "W") And (restriction <> "M") And (restriction <> "Y") Then
            Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "A valid restriction must be provided.")
        End If
    Else
        restriction = "NONE"
    End If

    Dim period As Variant: period = r.Value2
    
    If IsEmpty(period) Or (VarType(period) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "The trades sheet contains an invalid period in the cell " & wsCell & ".")
    End If
    
    period = UCase$(Trim$(period))

    If (period = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "The trades sheet contains an invalid period in the cell " & wsCell & ".")
    End If
    
    If (period = "-") Then
        Set GetPeriod = m_Host.GetFieldStaticByName("java.time.Period", "ZERO")
        Exit Function
    End If
    
    Dim periodValue As Variant: periodValue = Left$(period, Len(period) - 1)
    
    If Not IsNumeric(periodValue) Or (InStr(periodValue, ",") > 0) Or (InStr(periodValue, ".") > 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "The trades sheet contains an invalid period in the cell " & wsCell & ".")
    End If
    
    periodValue = CLng(periodValue)
    
    If (periodValue <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "The trades sheet contains an invalid period in the cell " & wsCell & ".")
    End If

    Dim periodTenor As String: periodTenor = Right$(period, 1)
    
    If (periodTenor <> "D") And (periodTenor <> "W") And (periodTenor <> "M") And (periodTenor <> "Y") Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "The trades sheet contains an invalid period in the cell " & wsCell & ".")
    End If
    
    If (restriction <> "NONE") And (periodTenor <> restriction) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPeriod", "The trades sheet contains a forbidden period in the cell " & wsCell & ".")
    End If
    
    Select Case periodTenor
    
        Case "D"
            Set GetPeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofDays", periodValue)
            
        Case "W"
            Set GetPeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofWeeks", periodValue)
        
        Case "M"
            Set GetPeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofMonths", periodValue)
        
        Case "Y"
            Set GetPeriod = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofYears", periodValue)
            
        Case Else
            Set GetPeriod = Nothing

    End Select

End Function

' METHOD
' Retrieves a LongShort from the given worksheet cell.

Public Function GetPosition(ByVal wsCell As String, ByVal asObject As Boolean) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPosition", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPosition", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPosition", "A valid cell must be provided.")
    End If
    
    Dim position As Variant: position = r.Value2

    If IsEmpty(position) Or (VarType(position) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPosition", "The trades sheet contains an invalid position in the cell " & wsCell & ".")
    End If
    
    position = UCase$(Trim$(position))

    If (position <> "LONG") And (position <> "SHORT") Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPosition", "The trades sheet contains an invalid position in the cell " & wsCell & ".")
    End If
    
    If asObject Then
        Set GetPosition = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.product.common.LongShort", "of", position)
    Else
        GetPosition = position
    End If

End Function

' METHOD
' Retrieves a rate from the given worksheet cell.

Public Function GetPrice(ByVal wsCell As String) As Double

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPrice", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPrice", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPrice", "A valid cell must be provided.")
    End If
    
    Dim price As Variant: price = r.Value2
    
    If IsEmpty(price) Or Not IsNumeric(price) Or (price <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetPrice", "The trades sheet contains an invalid price in the cell " & wsCell & ".")
    End If
    
    GetPrice = price / CDbl(100)

End Function

' METHOD
' Retrieves a quantity from the given worksheet cell.

Public Function GetQuantity(ByVal wsCell As String) As Double

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetQuantity", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetQuantity", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetQuantity", "A valid cell must be provided.")
    End If
    
    Dim quantity As Variant: quantity = r.Value2
    
    If IsEmpty(quantity) Or Not IsNumeric(quantity) Or (quantity <> CLng(quantity)) Or (quantity <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetQuantity", "The trades sheet contains an invalid quantity in the cell " & wsCell & ".")
    End If

    GetQuantity = CDbl(quantity)

End Function

' METHOD
' Retrieves a rate from the given worksheet cell.

Public Function GetRate(ByVal wsCell As String) As Double

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetRate", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetRate", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetRate", "A valid cell must be provided.")
    End If
    
    Dim rate As Variant: rate = r.Value2
    
    If IsEmpty(rate) Or Not IsNumeric(rate) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetRate", "The trades sheet contains an invalid rate in the cell " & wsCell & ".")
    End If
    
    GetRate = rate

End Function

' METHOD
' Retrieves a Frequency from the given worksheet cell.

Public Function GetTenor(ByVal wsCell As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "This instance has not been initialized yet.")
    End If

    wsCell = Trim$(wsCell)

    If (wsCell = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "A valid cell must be provided.")
    End If
    
    Dim r As Range: Set r = m_Worksheet.Range(wsCell)
    
    If (r.Columns.Count <> 1) Or (r.Rows.Count <> 1) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "A valid cell must be provided.")
    End If

    Dim tenor As Variant: tenor = r.Value2
    
    If IsEmpty(tenor) Or (VarType(tenor) <> vbString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "The trades sheet contains an invalid tenor in the cell " & wsCell & ".")
    End If
    
    tenor = UCase$(Trim$(tenor))

    If (tenor = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "The trades sheet contains an invalid tenor in the cell " & wsCell & ".")
    End If
    
    Dim tenorValue As Variant: tenorValue = Left$(tenor, Len(tenor) - 1)
    
    If Not IsNumeric(tenorValue) Or (InStr(tenorValue, ",") > 0) Or (InStr(tenorValue, ".") > 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "The trades sheet contains an invalid tenor in the cell " & wsCell & ".")
    End If
    
    tenorValue = CLng(tenorValue)
    
    If (tenorValue <= 0) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetTenor", "The trades sheet contains an invalid tenor in the cell " & wsCell & ".")
    End If

    Dim tenorPeriod As String: tenorPeriod = Right$(tenor, 1)
    
    If (tenorPeriod <> "D") And (tenorPeriod <> "W") And (tenorPeriod <> "M") And (tenorPeriod <> "Y") Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.GetFrequency", "The trades sheet contains an invalid tenor in the cell " & wsCell & ".")
    End If
    
    Select Case tenorPeriod
    
        Case "D"
            Set GetTenor = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "ofDays", tenorValue)
            
        Case "W"
            Set GetTenor = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "ofWeeks", tenorValue)
        
        Case "M"
            Set GetTenor = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "ofMonths", tenorValue)
        
        Case "Y"
            Set GetTenor = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.Tenor", "ofYears", tenorValue)
            
        Case Else
            Set GetTenor = Nothing

    End Select

End Function

' METHOD
' Initializes an instance of the DataDispatcher class.

Public Sub Initialize(ByVal host As RuntimeHost, ByVal ws As Worksheet)

    If m_Initialized Then
        Exit Sub
    End If

    If host Is Nothing Or Not host.Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.Initialize", "A valid Runtime Host instance must be provided.")
    End If
    
    If ws Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.Initialize", "A valid trades sheet must be provided.")
    End If
    
    Dim cc As Long: cc = ws.UsedRange.Columns.Count
    
    If (cc <= 2) Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.Initialize", "The trades sheet contains an invalid number of columns (" & CStr(cc) & ").")
    End If
    
    On Error Resume Next

    Dim wsCurves As Worksheet: Set wsCurves = ThisWorkbook.Sheets("Setup-Curves")
    
    On Error GoTo 0
    
    If wsCurves Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.Initialize", "The sheet ""Setup-Curves"" could not be found.")
    End If
    
    On Error Resume Next
    
    Dim wsExchange As Worksheet: Set wsExchange = ThisWorkbook.Sheets("Setup-Exchange")
    
    On Error GoTo 0

    If wsExchange Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.Initialize", "The sheet ""Setup-Exchange"" could not be found.")
    End If
    
    On Error GoTo ErrorHandler
    
    Dim form As New ReferenceData: form.Show
    Dim vd As Date: vd = form.ReferenceValuationDate
    Dim bdc As String: bdc = form.ReferenceBusinessDays
    Dim ccy As Variant: ccy = form.ReferenceCurrency
    Call Unload(form)

    Set m_Host = host
    Set m_Worksheet = ws

    Dim vdo As Variant: Set vdo = m_Host.InvokeMethodStaticFromName("java.time.LocalDate", "of", Year(vd), Month(vd), Day(vd))

    Dim ratesProviderBuilder As Variant: Set ratesProviderBuilder = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.pricer.rate.ImmutableRatesProvider", "builder", vdo)
    Call ParseWorksheetCurves(wsCurves, ratesProviderBuilder)
    Call ParseWorksheetExchange(wsExchange, ratesProviderBuilder)
    Set m_RatesProvider = m_Host.InvokeMethod(ratesProviderBuilder, "build")
    
    Dim cbase As Variant:  Set cbase = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", "USD")
    Dim ccou As Variant:  Set ccou = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", "EUR")
    Dim cp As Variant: Set cp = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.CurrencyPair", "of", cbase, ccou)
    Dim fxfr As Variant: Set fxfr = m_Host.InvokeMethod(m_RatesProvider, "fxForwardRates", cp)
    Dim p9m As Variant: Set p9m = m_Host.InvokeMethodStaticFromName("java.time.Period", "ofMonths", 9)
    Dim vd9m As Variant: Set vd9m = m_Host.InvokeMethod(vdo, "plus", p9m)
    Dim r9m As Variant: r9m = m_Host.InvokeMethod(fxfr, "rate", cbase, vd9m)

    Set m_Currency = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.currency.Currency", "of", ccy)
    Set m_BusinessDays = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.BusinessDayConvention", "of", Replace$(bdc, " ", ""))
    Set m_Calendar = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.date.HolidayCalendarId", "defaultByCurrency", m_Currency)
    Set m_ReferenceData = m_Host.InvokeMethodStaticFromName("com.opengamma.strata.basics.ReferenceData", "standard")
    Set m_DateFormatter = m_Host.InvokeMethodStaticFromName("java.time.format.DateTimeFormatter", "ofPattern", "dd/MM/yyyy")
    Set m_ValuationDate = vdo
    m_ValuationDatePlain = vd
    
    m_Initialized = True

    Exit Sub

ErrorHandler:

    Call PerformCleanup
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    
End Sub

' METHOD
' Prepares the sheet containing the cash flows.

Public Function PrepareCashFlowsSheet(ByVal ws As Worksheet) As Worksheet

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "DataDispatcher.PrepareCashFlowsSheet", "This instance has not been initialized yet.")
    End If

    On Error Resume Next

    Dim wsCashFlows As Worksheet: Set wsCashFlows = ThisWorkbook.Sheets("Cash Flows")
    
    On Error GoTo 0

    If Not wsCashFlows Is Nothing Then
        
        Dim alerts As Boolean: alerts = Application.DisplayAlerts
    
        Application.DisplayAlerts = False
        wsCashFlows.Delete
        Application.DisplayAlerts = alerts

    End If

    Set wsCashFlows = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    With wsCashFlows.Cells
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Interior.Color = 15132391
        .Style.HorizontalAlignment = xlHAlignCenter
        .Style.VerticalAlignment = xlCenter
    End With
    
    wsCashFlows.Name = "Cash Flows"
    wsCashFlows.Tab.Color = 11573124

    ws.Activate

    Set PrepareCashFlowsSheet = wsCashFlows

End Function
