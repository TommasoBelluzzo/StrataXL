VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReferenceData 
   Caption         =   "Reference Data"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "ReferenceData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReferenceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'====================================='
' Copyright (C) 2019 Tommaso Belluzzo '
'          Part of StrataXL           '
'====================================='

' SETTINGS

Option Explicit

' IMPORTS

Private Declare PtrSafe Function FindWindow Lib "User32.dll" Alias "FindWindowA" _
( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As LongPtr

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "User32.dll" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long _
    ) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "User32.dll" Alias "GetWindowLongA" _
    ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long _
    ) As LongPtr
#End If

#If Win64 Then
    Private Declare PtrSafe Function SetWindowLongPtr Lib "User32.dll" Alias "SetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr _
    ) As LongPtr
#Else
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
Dim m_ReferenceCurrency As String
Dim m_ReferenceValuationDate As Date

' PROPERTY
' Gets the reference currency.

Property Get ReferenceBusinessDays() As String

    ReferenceBusinessDays = m_ReferenceBusinessDays

End Property

' PROPERTY
' Gets the reference currency.

Property Get ReferenceCurrency() As String

    ReferenceCurrency = m_ReferenceCurrency

End Property

' PROPERTY
' Gets the reference valuation date.

Property Get ReferenceValuationDate() As Variant

    ReferenceValuationDate = m_ReferenceValuationDate

End Property

' CONSTRUCTOR

Private Sub UserForm_Initialize()

    Dim handle As Long, lStyle As Long
    
    If (Val(Application.Version) >= 9) Then
       handle = FindWindow("ThunderDFrame", Me.Caption)
    Else
       handle = FindWindow("ThunderXFrame", Me.Caption)
    End If
    
    Dim Style As Long: Style = GetWindowLongPtr(handle, GWL_STYLE)
    Call SetWindowLongPtr(handle, GWL_STYLE, Style And Not WS_SYSMENU)

    FieldValuationDate.Text = "15/02/2019"
    
    With FieldCurrency
        .AddItem "CHF"
        .AddItem "EUR"
        .AddItem "GBP"
        .AddItem "JPY"
        .AddItem "USD"
        .ListIndex = 1
    End With
    
    With FieldBusinessDays
        .AddItem "NO ADJUST"
        .AddItem "NEAREST"
        .AddItem "FOLLOWING"
        .AddItem "MODIFIED FOLLOWING"
        .AddItem "PRECEDING"
        .AddItem "MODIFIED PRECEDING"
        .ListIndex = 3
    End With

End Sub

' EVENT
' Raised when the OK button is clicked.

Private Sub ButtonOk_Click()

    Dim vd As String: vd = FieldValuationDate.Text
    
    If Not IsDate(vd) Then

        FieldValuationDate.BackColor = RGB(247, 215, 215)
        FieldValuationDate.BorderColor = RGB(255, 0, 0)

        Exit Sub

    End If
    
    FieldValuationDate.BackColor = &H8000000F
    FieldValuationDate.BorderColor = &H80000012
    
    m_ReferenceBusinessDays = FieldBusinessDays.Text
    m_ReferenceCurrency = FieldCurrency.Text
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
