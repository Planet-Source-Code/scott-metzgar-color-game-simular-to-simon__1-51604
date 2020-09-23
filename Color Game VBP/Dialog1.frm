VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Your Skill Level"
   ClientHeight    =   3405
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Dialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Begin Game"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Select a Number  (1-255)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   $"Dialog1.frx":0CCA
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Check, Check1 As Variant
Option Explicit

Private Sub OKButton_Click()
Check1 = 0
For Check = 1 To 255
If Text1.Text <> Check Then Check1 = Check1 + 1
Next Check
If Check1 > 254 Then Text1.Text = 15
Form1.Times = Int(Text1.Text)
Form1.Show
Dialog1.Hide

End Sub

