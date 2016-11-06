VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Test"
   ClientHeight    =   1470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "cmdConsulta"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   
   Dim oSefaz As New SefazClass
   
   oSefaz.NFeConsultaProtocolo "351610xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", "nomecertificado", "1"
   ShowXml oSefaz.cXmlEnvio
   ShowXml oSefaz.cXmlSoap
   ShowXml oSefaz.cXmlRetorno

End Sub


Function ShowXml(ByVal cText)
   
   Dim ctext2 As String
   
   ctext2 = ""
   Do While Len(cText) > 0
      ctext2 = ctext2 & Left(cText, 50) & vbCrLf
      cText = Mid(cText, 51)
   Loop
   MsgBox ctext2

End Function

