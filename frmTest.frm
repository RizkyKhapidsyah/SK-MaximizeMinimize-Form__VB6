VERSION 5.00
Object = "{C208CB66-02A2-11D4-90CD-9369BFCF0C5B}#3.0#0"; "ssMaxMin.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ssMaxMin.MaxMin MaxMin1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      MaxWidth        =   400
   End
   Begin VB.Label Label1 
      Caption         =   "Try to resize the form..."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormExtender1_Activate()
    'This will not fire the first time the Form is activated because at
    'that time the FormExtender control has not been initialized.
    Cls
    Print "Active"
End Sub

Private Sub FormExtender1_Deactivate()
    Cls
    Print "Inactive"
End Sub
