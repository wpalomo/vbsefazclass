VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Test"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Text            =   "Nome certificado"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "35170100000000000000000000000000000000000000"
      Top             =   840
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consulta na Sefaz"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "CN do Certificado"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Chave da Nota Fiscal Eletrônica"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   3255
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
   
   oSefaz.NFeConsultaProtocolo Text1.Text, "nomecertificado", "1"
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

