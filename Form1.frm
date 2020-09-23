VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printer"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Control"
      Height          =   2055
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   2175
      Begin VB.Label Secl 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Init 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Auto 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Strobe 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   2055
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2175
      Begin VB.Label Er 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Sel 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Pap 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Bus 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Ack 
         Caption         =   "label"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   240
         Top             =   960
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label DA 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Inp Lib "Out.dll" Alias "_Inp@4" (ByVal port As Integer) As Double

Const Data = &H378 ' decimal 888
Const Ctrl = &H37A ' decimal 890
Const Status = &H379 ' decimal 889
Dim Dat As Integer
Dim Ctr As Integer
Dim Stat As Integer

Dim OldDat As Integer
Dim OldCtr As Integer
Dim OldStat As Integer

Function ToBin(Number) As String
Dim Min As Integer
For I = 0 To 7
Select Case I
Case 0
Min = 128
Case 1
Min = 64
Case 2
Min = 32
Case 3
Min = 16
Case 4
Min = 8
Case 5
Min = 4
Case 6
Min = 2
Case 7
Min = 1
End Select
If Number - Min >= 0 Then
Number = Number - Min
ToBin = ToBin & "1"
Else
ToBin = ToBin & "0"
End If
Next
End Function

Function GoDat(Number As Integer)
Dim text As String
text = ToBin(Number)
For I = 0 To 7
DA(I).Caption = "Data " & CStr(I + 1) & " = " & Mid(text, I + 1, 1)
Next
End Function

Function GoCtrl(Number As Integer)
Dim text As String
text = ToBin(Number)
For I = 1 To 4
X = Mid(text, 8 - I, 1)

Select Case I
Case 1
Strobe.Caption = "Strobe = " & CStr(X)
Case 2
Auto.Caption = "Auto Feed = " & CStr(X)
Case 3
Init.Caption = "Initialize = " & CStr(X)
Case 4
If X = "0" Then
Secl.Caption = "Printer Selected"
Else
Secl.Caption = "Printer Not Selected"
End If
End Select
Next

End Function

Function GoStat(Number As Integer)
Dim text As String
text = ToBin(Number)
For I = 0 To 4
X = Mid(text, I + 1, 1)
Select Case I
Case 0
If X = "0" Then
Bus.Caption = "Printer Busy"
Else
Bus.Caption = "Printer Ready"
End If

Case 1
Ack.Caption = "Acknowledge = " & CStr(X)

Case 2
If X = "1" Then
Pap.Caption = "Paper Emty"
Else
Pap.Caption = "Paper full"
End If

Case 3
If X = "0" Then
Sel.Caption = "Printer Not Select"
Else
Sel.Caption = "Printer Select"
End If

Case 4
If X = "1" Then
Er.Caption = "No Error"
Else
Er.Caption = "Error"
End If
End Select
Next
End Function

Private Sub Timer1_Timer()
Dat = Inp(Data) ' Get the data from the printer port
Ctr = Inp(Ctrl) ' Get the Control from the printer port
Stat = Inp(Status) ' Get the status of the printer from the printer port
If Dat <> OldDat Then 'Only when Data is changed then GoDat
OldDat = Dat
GoDat Dat
End If
If Ctr <> OldCtr Then 'Do only GoCtrl if Control on printer port is changed
OldCtr = Ctr
GoCtrl Ctr
End If
If Stat <> OldStat Then 'Do only gostat if the status is changed
OldStat = Stat
GoStat Stat
End If
End Sub
