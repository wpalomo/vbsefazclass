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

Const WS_CTE_CONSULTACADASTRO = 1
Const WS_CTE_CONSULTAPROTOCOLO = 2
Const WS_CTE_INUTILIZACAO = 3
Const WS_CTE_RECEPCAO = 4
Const WS_CTE_RECEPCAOEVENTO = 5
Const WS_CTE_RETRECEPCAO = 6
Const WS_CTE_STATUSSERVICO = 7

Const WS_NFE_AUTORIZACAO = 8
Const WS_NFE_RETAUTORIZACAO = 9
Const WS_NFE_CANCELAMENTO = 10
Const WS_NFE_CONSULTACADASTRO = 11
Const WS_NFE_CONSULTAPROTOCOLO = 12
Const WS_NFE_INUTILIZACAO = 13
Const WS_NFE_RECEPCAO = 14
Const WS_NFE_RECEPCAOEVENTO = 15
Const WS_NFE_RETRECEPCAO = 16
Const WS_NFE_STATUSSERVICO = 17

Const WS_MDFE_RECEPCAO = 18
Const WS_MDFE_RETRECEPCAO = 19
Const WS_MDFE_RECEPCAOEVENTO = 20
Const WS_MDFE_CONSULTA = 21
Const WS_MDFE_STATUSSERVICO = 22
Const WS_MDFE_CONSNAOENC = 23

Const WS_NFE_DISTRIBUICAODFE = 24
Const WS_MDFE_DISTRIBUICAODFE = 25
Const WS_NFE_DOWNLOADNF = 26
Const WS_NFE_CONSULTADEST = 27

Const WS_AMBIENTE_HOMOLOGACAO = "2"
Const WS_AMBIENTE_PRODUCAO = "1"

Const WS_PROJETO_NFE = "nfe"
Const WS_PROJETO_CTE = "cte"
Const WS_PROJETO_MDFE = "mdfe"

Const INDSINC_RETORNA_PROTOCOLO = "1"
Const INDSINC_RETORNA_RECIBO = "0"

' configuração
Public cProjeto        As String
Public cAmbiente       As String
Public cVersao         As String
Public cScan           As String
Public cUF             As String
Public cCertificado    As String
Public cIndSinc        As String
Public nTempoEspera    As Integer
Public cUFTimeZone     As String
' xmls de cada etapa
Public cXmlDocumento   As String
Public cXmlEnvio       As String
Public cXmlSoap        As String
Public cXmlRetorno     As String
Public cXmlRecibo      As String
Public cXmlProtocolo   As String
Public cXmlAutorizado As String
Public cStatus         As String
Public cRecibo         As String
Public cMotivo         As String
' uso interno
Public cVersaoXml      As String
Public cSoapService    As String
Public cSoapAction     As String
Public cSoapURL        As String

Private Sub Class_initialize()

   cProjeto = WS_PROJETO_NFE
   cAmbiente = WS_AMBIENTE_PRODUCAO
   cVersao = "3.10"
   cScan = "N"
   cUF = "SP"
   cIndSinc = INDSINC_RETORNA_RECIBO
   nTempoEspera = 7
   cUFTimeZone = "SP"
   cXmlRetorno = "Erro Desconhecido"

   End Sub

Function CTeConsultaProtocolo(xChave As String, Optional xCertificado As String, Optional xAmbiente As String)

   Setup UFSigla(Mid(xChave, 1, 2)), xCertificado, xAmbiente, WS_CTE_CONSULTAPROTOCOLO

   cVersaoXml = "2.00"
   cXmlEnvio = "<consSitCTe versao=" & Quoted2(cVersaoXml) & " xmlns=" & Quoted2("http://www.portalfiscal.inf.br/cte") & ">" & _
                  XmlTag("tpAmb", cAmbiente) & _
                  XmlTag("xServ", "CONSULTAR") & _
                  XmlTag("chCTe", xChave) & _
                  "</consSitCTe>"
   If Mid(xChave, 21, 2) <> "57" Then
      cXmlRetorno = "*ERRO* Chave não se refere a CTE"
   Else
      XmlSoapPost
   End If
   CTeConsultaProtocolo = cXmlRetorno
End Function

Function CTeConsultaRecibo(Optional xRecibo As String, Optional xUF As String, Optional xCertificado As String, Optional xAmbiente As String)

   If Not IsMissing(xRecibo) Then cRecibo = xRecibo

   Setup xUF, xCertificado, xAmbiente, WS_CTE_RETRECEPCAO

   cVersaoXml = "2.00"
   cXmlEnvio = "<consReciCTe versao=" & Quoted2(cVersaoXml) & " xmlns=" & Quoted2("http://www.portalfiscal.inf.br/cte") & ">" & _
      XmlTag("tpAmb", cAmbiente) & _
      XmlTag("nRec", cRecibo) & _
      "</consReciCTe>"
   XmlSoapPost
   cXmlProtocolo = cXmlRetorno
   cMotivo = xmlnode(xmlnode(cXmlRetorno, "infProt"), "xMotivo")
   CTeConsultaRecibo = cXmlRetorno
End Function

Function CTeEventoCancela(xChave As String, xSequencia As Integer, xProt As String, xJust As String, Optional xCertificado As String, Optional xAmbiente As String)

   Setup UFSigla(Mid(xChave, 1, 2)), xCertificado, xAmbiente, WS_CTE_RECEPCAOEVENTO

   cXmlDocumento = "<eventoCTe xmlns=" & Quoted2("http://www.portalfiscal.inf.br/cte") & " versao=" & Quoted2("2.00") & ">" & _
      "<infEvento Id=" & Quoted2("ID110111" & xChave & StrZero(xSequencia, 2)) & ">" & _
      XmlTag("cOrgao", Mid(xChave, 1, 2)) & _
      XmlTag("tpAmb", cAmbiente) & _
      XmlTag("CNPJ", Mid(xChave, 7, 14)) & _
      XmlTag("chCTe", xChave) & _
      XmlTag("dhEvento", DateTimeXml(, , False)) & _
      XmlTag("tpEvento", "110111") & _
      XmlTag("nSeqEvento", LTrim(Str(xSequencia))) & _
      "<detEvento versaoEvento='2.00'>" & _
      "<evCancCTe>" & _
      XmlTag("descEvento", "Cancelamento") & _
      XmlTag("nProt", xProt) & _
      XmlTag("xJust", xJust) & _
      "</evCancCTe>" & _
      "</detEvento>" & _
      "</infEvento>" & _
      "</eventoCTe>"
   If assinaxml() = "OK" Then
      cXmlEnvio = cXmlDocumento
      XmlSoapPost
      cXmlProtocolo = cXmlRetorno
      CTEGeraAutorizado cXmlDocumento, cXmlProtocolo
   End If
   CTeEventoCancela = cXmlRetorno
End Function

Function NFeConsultaProtocolo(xChave As String, Optional xCertificado As String, Optional xAmbiente As String)

   Setup UFSigla(Mid(xChave, 1, 2)), xCertificado, xAmbiente, WS_CTE_CONSULTAPROTOCOLO

   cXmlEnvio = "<consSitNFe versao=" & Quoted2(cVersao) & " xmlns=" & Quoted2("http://www.portalfiscal.inf.br/nfe") & ">" & _
     XmlTag("tpAmb", cAmbiente) & _
     XmlTag("xServ", "CONSULTAR") & _
     XmlTag("chNFe", xChave) & _
     "</consSitNFe>"
   If InStr("55,65", Mid(xChave, 21, 2)) = 0 Then
      cXmlRetorno = "*ERRO* Chave não se refere a NFE"
   Else
      XmlSoapPost
   End If
   NFeConsultaProtocolo = cXmlRetorno

End Function

Function XmlSoapPost()

   If Len(cSoapURL) = 0 Then
      cXmlRetorno = "Erro SOAP: Não há endereço de webservice"
      Exit Function
   ElseIf Len(cSoapService) = 0 Then
      cXmlRetorno = "Erro SOAP: Não há nome de serviço"
      Exit Function
   ElseIf Len(cSoapAction) = 0 Then
      cXmlRetorno = "Erro SOAP: Não há endereço de SOAP Action"
      Exit Function
    End If
    XmlSoapEnvelope
    MicrosoftXmlSoapPost

End Function

Function XmlSoapEnvelope()

   cXmlSoap = "<?xml version='1.0' encoding='UTF-8'?>" & _
      "<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
      " xmlns:xsd='http:/www.w3.org/2001/XMLSchema'" & _
      " xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>" & _
      "<soap12:Header>" & _
      "<nfeCabecMsg xmlns='" & cSoapService & "'>" & _
      XmlTag("cUF", UFCodigo(cUF)) & _
      XmlTag("versaoDados", cVersao) & _
      "</nfeCabecMsg>" & _
      "</soap12:Header>" & _
      "<soap12:Body>" & _
      "<nfeDadosMsg xmlns='" & cSoapService & "'>" & _
      cXmlEnvio & _
      "</nfeDadosMsg>" & _
      "</soap12:Body>" & _
      "</soap12:Envelope>"
   cXmlSoap = Replace(cXmlSoap, "'", Chr(34))
   XmlSoapEnvelope = cXmlSoap
End Function

Function MicrosoftXmlSoapPost()

   Dim bRetorno() As Byte, oElement As Variant, oServer As Object

   Set oServer = CreateObject("MSXML2.ServerXMLHTTP")
   oServer.setOption 3, "CURRENT_USER\MY\" & cCertificado
   oServer.open "POST", cSoapURL, False
   oServer.setrequestheader "SOAPAction", cSoapAction
   oServer.setrequestheader "Content-Type", "application/soap+xml;charset=utf-8"
   oServer.send cXmlSoap
   oServer.waitforresponse 500
   bRetorno = oServer.responsebody
   cXmlRetorno = ""
   For Each oElement In bRetorno
      cXmlRetorno = cXmlRetorno & Chr(oElement)
    Next

End Function


Function XmlTag(cTag As String, cValue As String)

   XmlTag = "<" & cTag & ">" & cValue & "</" & cTag & ">"

End Function

Function UFCodigo(cUF As String)

   Dim cUFs As String, nPosicao As Integer
   cUFs = "AC,12,AL,27,AM,13,AP,16,BA,29,CE,23,DF,53,ES,32,GO,52,MG,31,MS,50,MT,51,MA,21,PA,15,PB,25,PE,26,PI,22,PR,41,RJ,33,RO,11,RN,24,RR,14,RS,43,SC,42,SE,28,SP,35,TO,17,"
   nPosicao = InStr(cUFs, cUF)
   If nPosicao < 1 Then
      UFCodigo = "99"
   Else
      UFCodigo = Mid(cUFs, nPosicao + 3, 2)
   End If

End Function

Function UFSigla(cUF As String)

   Dim cUFs As String, nPosicao As Integer

   cUFs = "AC,12,AL,27,AM,13,AP,16,BA,29,CE,23,DF,53,ES,32,GO,52,MG,31,MS,50,MT,51,MA,21,PA,15,PB,25,PE,26,PI,22,PR,41,RJ,33,RO,11,RN,24,RR,14,RS,43,SC,42,SE,28,SP,35,TO,17,"
   nPosicao = InStr(cUFs, cUF)
   If nPosicao < 1 Then
      UFSigla = "XX"
   Else
      UFSigla = Mid(cUFs, nPosicao - 3, 2)
   End If

End Function


Function Setup(xUF As String, xCertificado As String, xAmbiente As String, nWsServico As Integer)

   If Not IsMissing(xUF) Then cUF = xUF
   If Not IsMissing(xCertificado) Then cCertificado = xCertificado
   If Not IsMissing(xAmbiente) Then cAmbiente = xAmbiente

   If nWsServico = WS_CTE_CONSULTAPROTOCOLO Then
      cProjeto = WS_PROJETO_CTE
      cSoapAction = "cteConsultaCT"
      cSoapService = "http://www.portalfiscal.inf.br/cte/wsdl/CteConsulta"
   End If
   SetSoapURL nWsServico
   End Function

Function SetSoapURL(nWsServico As Integer)

   If cProjeto = WS_PROJETO_NFE Then
      cSoapURL = UrlWebService(cUF, cAmbiente, nWsServico, cVersao)
   End If
   If Len(cSoapURL) = 0 Then
      If cProjeto = WS_PROJETO_NFE Then
         If InStr("AC,AL,AP,DF,ES,PB,RJ,RN,RO,RR,SC,SE,TO", cUF) = 0 Then
            cSoapURL = UrlWebService("SVRS", cAmbiente, nWsServico, cVersao)
         ElseIf InStr("MA,PA,PI", cUF) = 0 Then
            cSoapURL = UrlWebService("SVAN", cAmbiente, nWsServico, cVersao)
         End If
         If Len(cSoapURL) = 0 Then
            cSoapURL = UrlWebService("AN", cAmbiente, nWsServico, cVersao)
         End If
      End If
    End If

End Function

Function UrlWebService(cUF As String, cAmbiente As String, nWsServico As Integer, cVersao As String)

   Select Case cUF
   Case "AC"
      UrlWebService = UrlWebService("SVRS", cAmbiente, nWsServico, cVersao)
   Case "AM"
      Select Case cAmbiente
      Case WS_AMBIENTE_PRODUCAO
         If WS_NFE_CONSULTAPROTOCOLO Then UrlWebService = "https://nfe.sefaz.am.gov.br/services2/services/NfeConsulta2"
      End Select
   Case "SP"
      Select Case cAmbiente
      Case WS_AMBIENTE_PRODUCAO
         If WS_NFE_CONSULTAPROTOCOLO Then UrlWebService = "https://nfe.fazenda.sp.gov.br/ws/nfeconsulta2.asmx"
      End Select
   End Select
End Function

Function xmlnode(cXml As String, cNode As String) As String
' apenas pra não dar erro por faltar a função
End Function

Function StrZero(nValue As Integer, nLen As Integer)
   StrZero = Format(nValue, String("0", nLen))
End Function

Function DateTimeXml(Optional dDate As Date, Optional cTime As String, Optional cUF As String, Optional lUTC As Boolean)
If IsMissing(dDate) Then dDate = Now()
If IsMissing(cTime) Then cTime = Time()
If IsMissing(cUF) Then cUF = "SP"
If IsMissing(lUTC) Then lUTC = True
DateTimeXml = Format(Year(dDate), "0000") & "-" & Format(Month(dDate), "00") & "-" & Format(Day(dDate), "00") & "T" & Time() & "-02"
End Function

Function assinaxml()
  ' vai assinar cXmlDocumento usando cCertificado
End Function

Function CTEGeraAutorizado(cXmlDocumento, cXmlProtocolo)

End Function

Function Quoted2(ByVal cTexto As String) As String
   Quoted2 = Chr(34) & cTexto & Chr(34)
End Function
