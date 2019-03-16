VERSION 5.00
Begin VB.Form TestForm 
   Caption         =   "Test Form"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Call Net Standard Library Method"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim oMathMethods As IMathUtils
    Set oMathMethods = New NetStandardLibrary.MathUtils

    MsgBox ("Result from complex method is: " & oMathMethods.CalculateComplexMethod())
End Sub
