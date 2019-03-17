Attribute VB_Name = "Calibration"

'====================================='
' Copyright (C) 2019 Tommaso Belluzzo '
'          Part of StrataXL           '
'====================================='

Option Base 0
Option Explicit

Public Sub CalibrateCurvesCrossCurrency()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Cross-Currency Curves")
    Dim cc As Long: cc = ws.UsedRange.Columns.Count
    Dim rc As Long: rc = ws.UsedRange.Rows.Count
    
    Dim host As New RuntimeHost: Call host.Initialize
    Dim cm As New CurvesManager: Call cm.Initialize(host, ws, False)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call cm.CalibrateCurves
    Call cm.PrepareResultsSheet
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Call Err.Raise(Err.Number, Err.Source, Err.Description)

End Sub

Public Sub CalibrateCurvesSingleCurrency()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Single-Currency Curves")
    Dim cc As Long: cc = ws.UsedRange.Columns.Count
    Dim rc As Long: rc = ws.UsedRange.Rows.Count
    
    Dim host As New RuntimeHost: Call host.Initialize
    Dim cm As New CurvesManager: Call cm.Initialize(host, ws, True)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call cm.CalibrateCurves
    Call cm.PrepareResultsSheet
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Call Err.Raise(Err.Number, Err.Source, Err.Description)

End Sub
