VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SefazClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const WS_NFE_CONSULTAPROTOCOLO = 12

Const WS_AMBIENTE_HOMOLOGACAO = "2"
Const WS_AMBIENTE_PRODUCAO = "1"

Const WS_PROJETO_NFE = "nfe"
Const WS_PROJETO_CTE = "cte"
Const WS_PROJETO_MDFE = "mdfe"

Public cChave As String
Public cCertificado As String
Public cSoapAction As String
Public cSoapWs As String
Public cServico As String
Public cXmlDados As String
Public cXmlEnvelope As String
Public cXmlRetorno As String
Public cAmbiente As String
Public cUF As String
Public cVersao As String

Private Sub Class_initialize()
   
   cUF = "SP"
   cVersao = "3.10"
   cAmbiente = WS_AMBIENTE_PRODUCAO

End Sub

Function nfeConsultaProtocolo(xChave, Optional xCertificado As String, Optional xAmbiente As String)

   If Not IsMissing(xAmbiente) Then cAmbiente = xAmbiente
   If Not IsMissing(xCertificado) Then cCertificado = xCertificado
   cUF = UFSigla(Mid(xChave, 1, 2))
   cServico = "http://www.portalfiscal.inf.br/nfe/wsdl/NfeConsulta2"
   cSoapAction = "NfeConsulta2"
   SetSoapWs WS_NFE_CONSULTAPROTOCOLO
   cXmlDados = "<consSitNFe versao='" & cVersao & "' xmlns='http://www.portalfiscal.inf.br/nfe'>" & _
     XmlTag("tpAmb", cAmbiente) & _
     XmlTag("xServ", "CONSULTAR") & _
     XmlTag("chNFe", cChave) & _
     "</consSitNFe>"
   XmlSoapEnvelope
   MicrosoftXmlSoapPost

End Function


Function XmlSoapEnvelope()

   cXmlEnvelope = "<?xml version='1.0' encoding='UTF-8'?>" & _
      "<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
      " xmlns:xsd='http:/www.w3.org/2001/XMLSchema'" & _
      " xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>" & _
      "<soap12:Header>" & _
      "<nfeCabecMsg xmlns='" & cServico & "'>" & _
      XmlTag("cUF", UFCodigo(cUF)) & _
      XmlTag("versaoDados", cVersao) & _
      "</nfeCabecMsg>" & _
      "</soap12:Header>" & _
      "<soap12:Body>" & _
      "<nfeDadosMsg xmlns='" & cServico & "'>" & _
      cXmlDados & _
      "</nfeDadosMsg>" & _
      "</soap12:Body>" & _
      "</soap12:Envelope>"
   cXmlEnvelope = Replace(cXmlEnvelope, "'", Chr(34))

End Function

Function MicrosoftXmlSoapPost()

   Dim bRetorno() As Byte, oElement As Variant, oServer As Object

   Set oServer = CreateObject("MSXML2.ServerXMLHTTP")
   oServer.setOption 3, "CURRENT_USER\MY\" & cCertificado
   oServer.open "POST", cSoapWs, False
   oServer.setrequestheader "SOAPAction", cSoapAction
   oServer.setrequestheader "Content-Type", "application/soap+xml;charset=utf-8"
   oServer.send cXmlEnvelope
   oServer.waitforresponse 500
   bRetorno = oServer.responsebody
   cXmlRetorno = ""
   For Each oElement In bRetorno
      cXmlRetorno = cXmlRetorno & Chr(oElement)
    Next

End Function

Function SetSoapWs(nWsServico As Integer)

   Select Case cUF
      Case "SP"
         Select Case cAmbiente
         Case WS_AMBIENTE_PRODUCAO
            If WS_NFE_CONSULTAPROTOCOLO Then cSoapWs = "https://nfe.fazenda.sp.gov.br/ws/nfeconsulta2.asmx"
         End Select
   End Select
End Function


Function XmlTag(cTag As String, cValue As String)

   XmlTag = "<" & cTag & ">" & cValue & "</" & cTag & ">"

End Function

Function UFCodigo(cUF As String)
   
   Select Case cUF
   Case "SP": UFCodigo = "35"
   End Select

End Function

Function UFSigla(cUF As String)
   
   Select Case cUF
   Case "35": UFSigla = "SP"
   End Select

End Function
