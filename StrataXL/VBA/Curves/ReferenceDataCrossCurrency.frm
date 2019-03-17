VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReferenceDataCrossCurrency 
   Caption         =   "Reference Data"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "ReferenceDataCrossCurrency.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReferenceDataCrossCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'====================================='
' Copyright (C) 2019 Tommaso Belluzzo '
'          Part of StrataXL           '
'====================================='

' SETTINGS

Option Base 0
Option Explicit

' IMPORTS

#If Win64 Then

    Private Declare PtrSafe Function FindWindow Lib "User32.dll" Alias "FindWindowA" _
    ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String _
    ) As LongPtr
    
    Private Declare PtrSafe Function GetWindowLongPtr Lib "User32.dll" Alias "GetWindowLongPtrA" _
    ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long _
    ) As LongPtr
    
    Private Declare PtrSafe Function SetWindowLongPtr Lib "User32.dll" Alias "SetWindowLongPtrA" _
    ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr _
    ) As LongPtr

#Else

    Private Declare PtrSafe Function FindWindow Lib "User32.dll" Alias "FindWindowA" _
    ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String _
    ) As Long

    Private Declare PtrSafe Function GetWindowLongPtr Lib "User32.dll" Alias "GetWindowLongA" _
    ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long _
    ) As LongPtr
    
    Private Declare PtrSafe Function SetWindowLongPtr Lib "User32.dll" Alias "SetWindowLongA" _
    ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr _
    ) As LongPtr

#End If

' CONSTANTS

Const GWL_STYLE As Long = -16
Const WS_SYSMENU As Long = &H80000

' MEMBERS

Dim m_ReferenceBusinessDays As String
Dim m_ReferenceCurrencyForeign As String
Dim m_ReferenceCurrencyLocal As String
Dim m_ReferenceDaysCount As String
Dim m_ReferenceDaysOffset As Long
Dim m_ReferenceExchangeRate As Double
Dim m_ReferenceValuationDate As Date

' PROPERTY
' Gets the reference business days convention.

Property Get ReferenceBusinessDays() As String

    ReferenceBusinessDays = m_ReferenceBusinessDays

End Property

' PROPERTY
' Gets the reference foreign currency.

Property Get ReferenceCurrencyForeign() As String

    ReferenceCurrencyForeign = m_ReferenceCurrencyForeign

End Property

' PROPERTY
' Gets the reference local currency.

Property Get ReferenceCurrencyLocal() As String

    ReferenceCurrencyLocal = m_ReferenceCurrencyLocal

End Property

' PROPERTY
' Gets the reference days offset.

Property Get ReferenceDaysOffset() As Long

    ReferenceDaysOffset = m_ReferenceDaysOffset

End Property

' PROPERTY
' Gets the reference exchange rate.

Property Get ReferenceExchangeRate() As String

    ReferenceExchangeRate = m_ReferenceExchangeRate

End Property

' PROPERTY
' Gets the reference days count convention.

Property Get ReferenceDaysCount() As String

    ReferenceDaysCount = m_ReferenceDaysCount

End Property

' PROPERTY
' Gets the reference valuation date.

Property Get ReferenceValuationDate() As Variant

    ReferenceValuationDate = m_ReferenceValuationDate

End Property

' CONSTRUCTOR

Private Sub UserForm_Initialize()

    Dim handle As LongPtr
    
    If (Val(Application.Version) >= 9) Then
       handle = FindWindow("ThunderDFrame", Me.Caption)
    Else
       handle = FindWindow("ThunderXFrame", Me.Caption)
    End If
    
    Dim lStyle As LongPtr: lStyle = GetWindowLongPtr(handle, GWL_STYLE)
    Call SetWindowLongPtr(handle, GWL_STYLE, lStyle And Not WS_SYSMENU)

    FieldValuationDate.Text = "15/02/2019"
    
    With FieldCurrencyLocal
        .AddItem "AED"
        .AddItem "ARS"
        .AddItem "AUD"
        .AddItem "BGN"
        .AddItem "BHD"
        .AddItem "BRL"
        .AddItem "CAD"
        .AddItem "CHF"
        .AddItem "CLP"
        .AddItem "CNY"
        .AddItem "CNY"
        .AddItem "COP"
        .AddItem "CZK"
        .AddItem "DKK"
        .AddItem "EGP"
        .AddItem "EUR"
        .AddItem "GBP"
        .AddItem "HKD"
        .AddItem "HRK"
        .AddItem "HUF"
        .AddItem "IDR"
        .AddItem "ILS"
        .AddItem "INR"
        .AddItem "ISK"
        .AddItem "JPY"
        .AddItem "KRW"
        .AddItem "MXN"
        .AddItem "MYR"
        .AddItem "NOK"
        .AddItem "NZD"
        .AddItem "PEN"
        .AddItem "PHP"
        .AddItem "PKR"
        .AddItem "PLN"
        .AddItem "RON"
        .AddItem "RUB"
        .AddItem "SAR"
        .AddItem "SEK"
        .AddItem "SGD"
        .AddItem "THB"
        .AddItem "TRY"
        .AddItem "TWD"
        .AddItem "UAH"
        .AddItem "USD"
        .AddItem "ZAR"
        .ListIndex = 43
    End With
    
    With FieldCurrencyForeign
        .AddItem "AED"
        .AddItem "ARS"
        .AddItem "AUD"
        .AddItem "BGN"
        .AddItem "BHD"
        .AddItem "BRL"
        .AddItem "CAD"
        .AddItem "CHF"
        .AddItem "CLP"
        .AddItem "CNY"
        .AddItem "CNY"
        .AddItem "COP"
        .AddItem "CZK"
        .AddItem "DKK"
        .AddItem "EGP"
        .AddItem "EUR"
        .AddItem "GBP"
        .AddItem "HKD"
        .AddItem "HRK"
        .AddItem "HUF"
        .AddItem "IDR"
        .AddItem "ILS"
        .AddItem "INR"
        .AddItem "ISK"
        .AddItem "JPY"
        .AddItem "KRW"
        .AddItem "MXN"
        .AddItem "MYR"
        .AddItem "NOK"
        .AddItem "NZD"
        .AddItem "PEN"
        .AddItem "PHP"
        .AddItem "PKR"
        .AddItem "PLN"
        .AddItem "RON"
        .AddItem "RUB"
        .AddItem "SAR"
        .AddItem "SEK"
        .AddItem "SGD"
        .AddItem "THB"
        .AddItem "TRY"
        .AddItem "TWD"
        .AddItem "UAH"
        .AddItem "USD"
        .AddItem "ZAR"
        .ListIndex = 15
    End With
    
    FieldExchangeRate.Text = "0.887415"
    
    With FieldBusinessDays
        .AddItem "NO ADJUST"
        .AddItem "NEAREST"
        .AddItem "FOLLOWING"
        .AddItem "MODIFIED FOLLOWING"
        .AddItem "PRECEDING"
        .AddItem "MODIFIED PRECEDING"
        .ListIndex = 3
    End With
    
    With FieldDaysCount
        .AddItem "30/360 ISDA"
        .AddItem "30/360 PSA"
        .AddItem "30E/360"
        .AddItem "30E/360 ISDA"
        .AddItem "30E+/360"
        .AddItem "30U/360"
        .AddItem "30U/360 EOM"
        .AddItem "ACT/360"
        .AddItem "ACT/364"
        .AddItem "ACT/365.25"
        .AddItem "ACT/365A"
        .AddItem "ACT/365F"
        .AddItem "ACT/365L"
        .AddItem "ACT/ACT AFB"
        .AddItem "ACT/ACT ICMA"
        .AddItem "ACT/ACT ISDA"
        .AddItem "ACT/ACT AFB"
        .AddItem "ACT/ACT ICMA"
        .AddItem "ACT/ACT YEAR"
        .AddItem "NL/365"
        .ListIndex = 7
    End With
    
    FieldDaysOffset.Text = "2"

End Sub

' EVENT
' Raised when the OK button is clicked.

Private Sub ButtonOk_Click()

    Dim vd As String: vd = FieldValuationDate.Text
    Dim ccyLocal As String: ccyLocal = FieldCurrencyLocal.Text
    Dim ccyForeign As String: ccyForeign = FieldCurrencyForeign.Text
    Dim exchangeRate As String: exchangeRate = FieldExchangeRate.Text
    Dim offset As String: offset = FieldDaysOffset.Text
    
    Dim shouldExit As Boolean: shouldExit = False
    
    If Not IsDate(vd) Or (DateDiff("d", CDate(vd), Now()) < 0) Then
        FieldValuationDate.BackColor = RGB(247, 215, 215)
        FieldValuationDate.BorderColor = RGB(255, 0, 0)
        shouldExit = True
    End If
    
    If (ccyLocal = ccyForeign) Then
        FieldCurrencyLocal.BackColor = RGB(247, 215, 215)
        FieldCurrencyLocal.BorderColor = RGB(255, 0, 0)
        FieldCurrencyForeign.BackColor = RGB(247, 215, 215)
        FieldCurrencyForeign.BorderColor = RGB(255, 0, 0)
        shouldExit = True
    End If
    
    If (ccyLocal = ccyForeign) Then
        FieldCurrencyLocal.BackColor = RGB(247, 215, 215)
        FieldCurrencyLocal.BorderColor = RGB(255, 0, 0)
        FieldCurrencyForeign.BackColor = RGB(247, 215, 215)
        FieldCurrencyForeign.BorderColor = RGB(255, 0, 0)
        shouldExit = True
    End If
    
    If Not IsNumeric(exchangeRate) Or (CDbl(exchangeRate) <= 0) Then
        FieldExchangeRate.BackColor = RGB(247, 215, 215)
        FieldExchangeRate.BorderColor = RGB(255, 0, 0)
        shouldExit = True
    End If
    
    If (offset <> "0") And (offset <> "1") And (offset <> "2") And (offset <> "3") Then
        FieldDaysOffset.BackColor = RGB(247, 215, 215)
        FieldDaysOffset.BorderColor = RGB(255, 0, 0)
        shouldExit = True
    End If
    
    If shouldExit Then
        Exit Sub
    End If
    
    FieldValuationDate.BackColor = &H8000000F
    FieldValuationDate.BorderColor = &H80000012
    FieldCurrencyLocal.BackColor = &H8000000F
    FieldCurrencyLocal.BorderColor = &H80000012
    FieldCurrencyForeign.BackColor = &H8000000F
    FieldCurrencyForeign.BorderColor = &H80000012
    FieldExchangeRate.BackColor = &H8000000F
    FieldExchangeRate.BorderColor = &H80000012
    FieldDaysOffset.BackColor = &H8000000F
    FieldDaysOffset.BorderColor = &H80000012

    m_ReferenceBusinessDays = FieldBusinessDays.Text
    m_ReferenceCurrencyForeign = FieldCurrencyForeign.Text
    m_ReferenceCurrencyLocal = FieldCurrencyLocal.Text
    m_ReferenceDaysCount = FieldDaysCount.Text
    m_ReferenceDaysOffset = CLng(offset)
    m_ReferenceExchangeRate = CDbl(exchangeRate)
    m_ReferenceValuationDate = CDate(vd)

    Me.Hide

End Sub

' EVENT
' Raised when the form is closed.

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If (CloseMode = vbFormControlMenu) Then
        Cancel = True
    End If

End Sub
