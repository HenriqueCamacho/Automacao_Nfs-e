Imports CClass
Imports NFSeCore
Imports System.Xml
Imports System.Web.Services.Protocols
Imports System.Security.Cryptography.X509Certificates
Imports System.text.RegularExpressions
Imports System.IO
Imports CertificadoDigital

Namespace c2931350
    Public Class ConsultarNfseRps
        Inherits AConsultarNfseRps

        Public dtsNfse As DtsNfse = New DtsNfse
        Private isExternal As Boolean = False

        Public Sub New(ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
        End Sub

        Public Sub New(ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar, _
                       ByVal isExternal As Boolean)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            Me.isExternal = isExternal
        End Sub

        Public Overrides Function processar(ByVal conn As IDbConnection, ByVal trans As IDbTransaction, ByVal tipoRps As Integer, ByVal serieRps As String, ByVal numeroRps As Integer, _
                                  ByVal numrCnpj As Long, ByVal idEmpresa As Integer) As String

            Dim propXml As IPropriedadesXML = Nothing
            Dim cidadeRn As NfseCidadesRN = Nothing
            Dim ambiente As TipoAmbiente = Nothing
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim capa As NfseCapaData = Nothing
            Dim empresaData As NfseEmpresasData = GetIntegracaoEmpresa(idEmpresa)
            Dim cidadeData As NfseCidadesData = GetCidade(empresaData.codgCidade)
            capa = GetIntegracaoCapa(tipoRps, serieRps, numeroRps, numrCnpj)
            Dim xmlRetorno As String = Nothing
            Dim responseSOAP As New XmlDocument()
            Dim certificado As X509Certificate = Nothing
            Dim url As String = Nothing
            Dim result As Object
            Me.conn = conn
            Me.trans = trans

            Try

                ' obter o ambiente producao ou homologacao
                ambiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")

                ' obter implementacao de consulta de nfse por rps da cidade selecionada
                propXml = FactoryCore.GetPropriedadesXML(empresaData.codgCidade)
                Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(empresaData.codgCidade)
                propXml = FactoryCore.GetPropriedadesXML(cidadeData.codgCidade)
                Dim xmlConsultaPorRpsEnvio As String
                'If ambiente = TipoAmbiente.HOMOLOGACAO Then
                '    xmlConsultaPorRpsEnvio = propXml.GetXmlConsultaNfseRps(CStr(tipoRps), _
                '                                  CStr(numeroRps), _
                '                                  serieRps, 1, _
                '                                  "02974456001230", _
                '                                  "9096")
                'Else
                LogInfo("ConsultarNfseRps xmlConsultaPorRpsEnvio 001: ")
                xmlConsultaPorRpsEnvio = propXml.GetXmlConsultaNfseRps(CStr(tipoRps), _
                                                                  CStr(numeroRps), _
                                                                  serieRps, 1, _
                                                                  numrCnpj.ToString("00000000000000"), _
                                                                  Regex.Replace(capa.inscMunicipal, _
                                                                  "[^0-9a-zA-Z]", ""))
                'End If
                LogInfo("ConsultarNfseRps xmlConsultaPorRpsEnvio 1: ")
                xmlConsultaPorRpsEnvio = "<Envelope xmlns=""http://schemas.xmlsoap.org/soap/envelope/"">" + _
                                            "<Header/>" + _
                                            "<Body>" + _
                                                "<ConsultarNfsePorRps xmlns=""http://nfse.eunapolis.ba.gov.br/webrun/webservices/NFEServices.jws"">" + _
                                                    "<Nfsecabecmsg><![CDATA[" + propXml.GetCabecalho() + "]]></Nfsecabecmsg>" + _
                                                    "<Nfsedadosmsg><![CDATA[" + xmlConsultaPorRpsEnvio + "]]></Nfsedadosmsg>" + _
                                                "</ConsultarNfsePorRps>" + _
                                            "</Body>" + _
                                         "</Envelope>"

                LogInfo("ConsultarNfseRps xmlConsultaPorRpsEnvio 2: " + xmlConsultaPorRpsEnvio)
                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, empresaData.codgCidade)
                LogInfo(idEmpresa, "ConsultarNfseRps xmlConsultaPorRpsEnvio 3")
                'If ambiente = TipoAmbiente.PRODUCAO Then
                certificado = propCertificado.GetCertificadoTransmissao(idEmpresa, numrCnpj)
                'Else
                '    certificado = propCertificado.GetCertificadoTransmissao(idEmpresa, 2974456000188)
                'End If
                LogInfo(idEmpresa, "ConsultarNfseRps xmlConsultaPorRpsEnvio 3.1")
                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para CONSULTA DE NFSE POR RPS, através do nº do documento informado : " + numrCnpj.ToString)
                End If
                LogInfo(idEmpresa, "ConsultarNfseRps xmlConsultaPorRpsEnvio 3.2")
                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = ParametroUtil.GetParametroCidade(empresaData.codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    url = ParametroUtil.GetParametroCidade(empresaData.codgCidade, "URL_AMBIENTE_PRODUCAO")
                End If

                Dim doc As XmlDocument = New XmlDocument()
                doc.PreserveWhitespace = True
                doc.LoadXml(xmlConsultaPorRpsEnvio)
                xmlRetorno = HTTPSoapTextRequest(certificado, doc, url, idEmpresa, "")
                responseSOAP.LoadXml(xmlRetorno)
                LogInfo("ConsultarNfseRps xmlConsultaPorRpsEnvio 4: " + xmlRetorno)
                Dim node As XmlNode = responseSOAP.SelectSingleNode(GetXpathFromString("Envelope, Body, ConsultarNfsePorRpsResponse, ConsultarNfsePorRpsReturn "))
                node.InnerXml = node.InnerXml.Replace("<![CDATA[", "").Replace("]]>", "").Replace("&lt;", "<").Replace("&gt;", ">")

                If node Is Nothing Then
                    Throw New FalhaException("Retorno da consulta de NFS-e por RPS da prefeitura não pode ser processado: objeto veio nulo. Retorno: " + xmlRetorno)
                End If

                result = node.InnerXml

                xmlRetorno = result
            Catch ex As Exception
                Throw ex
            End Try

            Return result
        End Function

        Public Overrides Sub gravarXmlNfse(ByVal conn As IDbConnection, ByVal trans As IDbTransaction, ByVal xmlNfse As String, ByVal codgMunicipio As Integer, ByVal idEmpresa As Integer)
            Dim propXml As IPropriedadesXML = Nothing
            Dim numeroRps As Integer = 0
            Dim serieRps As String = 0
            Dim tipoRps As Integer = 0
            Dim numrCnpj As Long = 0
            Dim xmldoc As XmlDocument = Nothing
            Dim xmlIdentificacaoRps As XmlDocument = Nothing
            Dim movimentoRn As NfseMovimentoRN = Nothing
            Dim dtsNfse As DtsNfse = Nothing
            Dim ambiente As TipoAmbiente = TipoAmbiente.HOMOLOGACAO
            Dim movimentoData As NfseMovimentoData = Nothing
            ambiente = CInt(ParametroUtil.GetParametro(idEmpresa, "AMBIENTE"))
            Try
                ' obter o document xml 
                xmldoc = New XmlDocument()
                xmldoc.LoadXml(xmlNfse)

                ' obter o document da tag identificacaorps
                xmlIdentificacaoRps = XmlUtil.getDocByTag(xmldoc, "IdentificacaoRps")

                ' obter o numero do RPS
                If XmlUtil.getValorTag(xmlIdentificacaoRps, "Numero").Trim.Length > 0 Then
                    numeroRps = CInt(XmlUtil.getValorTag(xmlIdentificacaoRps, "Numero"))
                Else
                    Throw New FalhaException("Erro ao obter o número do RPS no xml informado")
                End If

                ' obter a serie do rps
                If XmlUtil.getValorTag(xmlIdentificacaoRps, "Serie").Trim.Length > 0 Then
                    serieRps = XmlUtil.getValorTag(xmlIdentificacaoRps, "Serie")
                Else
                    Throw New FalhaException("Erro ao obter o número do RPS no xml informado")
                End If

                ' obter o tipo do rps
                If XmlUtil.getValorTag(xmlIdentificacaoRps, "Tipo").Trim.Length > 0 Then
                    tipoRps = CInt(XmlUtil.getValorTag(xmlIdentificacaoRps, "Tipo"))
                Else
                    Throw New FalhaException("Erro ao obter o número do RPS no xml informado")
                End If

                ' obter o CNPJ da nfse
                'If ambiente = TipoAmbiente.HOMOLOGACAO Then
                '    numrCnpj = CLng(2348447000181)
                'Else
                If XmlUtil.getValorTag(xmldoc, "Cnpj").Trim.Length > 0 Then
                    numrCnpj = CLng(XmlUtil.getValorTag(xmldoc, "Cnpj"))
                Else
                    Throw New FalhaException("Erro ao obter o número do CNPJ do prestador no xml informado")
                End If
                'End If



                ' obter o ambiente


                ' obter o movimento do RPS
                dtsNfse = New DtsNfse()
                movimentoRn = New NfseMovimentoRN(factory, publicVar, dtsNfse)
                movimentoRn.filtro = " TIPO_RPS = :TIPO_RPS AND SERIE_RPS = :SERIE_RPS AND NUMERO_RPS = :NUMERO_RPS AND NUMR_CNPJ = :NUMR_CNPJ AND TIPO_AMBIENTE = :TIPO_AMBIENTE"
                movimentoRn.params.Add(New ParamDB("TIPO_RPS", tipoRps, DbType.Int64))
                movimentoRn.params.Add(New ParamDB("SERIE_RPS", serieRps, DbType.AnsiString))
                movimentoRn.params.Add(New ParamDB("NUMERO_RPS", numeroRps, DbType.Int64))
                movimentoRn.params.Add(New ParamDB("NUMR_CNPJ", numrCnpj, DbType.Int64))
                movimentoRn.params.Add(New ParamDB("TIPO_AMBIENTE", ambiente, DbType.Int64))
                movimentoRn.Listar()

                ' alterar o movimento
                If dtsNfse.NFSE_MOVIMENTO.Count > 0 Then
                    movimentoData = SetMovimentoRowToData(dtsNfse.NFSE_MOVIMENTO(0))
                    movimentoData.xmlNota = xmlNfse
                    movimentoData.numrNfse = CLng(XmlUtil.getValorTag(xmldoc, "Numero"))
                    movimentoData.codgVerificacao = XmlUtil.getValorTag(xmldoc, "CodigoVerificacao")
                    If XmlUtil.getDocByTag(xmldoc, "NfseCancelamento") Is Nothing Then
                        movimentoData.status = StatusProcessamento.AUTORIZADA
                    Else
                        movimentoData.status = StatusProcessamento.CANCELADA
                    End If
                    MyBase.AtualizaVendas_WEB_DELPHI(conn, trans, movimentoData.serieRps, movimentoData.numeroRps, idEmpresa, movimentoData.numrNfse)
                    movimentoRn.alterar(conn, trans, movimentoData)
                Else
                    Throw New FalhaException("Nenhum registro encontrado com os dados do RPS informado")
                End If
            Catch ex As Exception
                If TypeOf ex Is FalhaException Then
                    Throw ex
                Else
                    Throw New FalhaException("Erro ao gravar NFS-e", ex)
                End If
            End Try
        End Sub

        Protected Overrides Function isRetornoProcessamentoOk(ByVal xml As String) As Boolean
            xml = CClass.Util.TString(xml)

            If xml = "ERRO PREFEITURA" Then
                Return True
            Else
                Dim xmldoc As XmlDocument = New XmlDocument

                Try
                    xmldoc.LoadXml(xml)
                Catch ex As Exception
                    Throw New FalhaException("Erro ao ler XML de retorno: " + ex.Message + _
                                             " - conteudo do retorno: " + xml)
                End Try

                Dim retorno As Boolean

                retorno = CClass.XmlUtil.getDocByTag(xmldoc, "CompNfse") IsNot Nothing

                If retorno = False Then
                    retorno = CClass.XmlUtil.getDocByTag(xmldoc, "ListaMensagemRetorno") IsNot Nothing
                End If

                Return retorno
            End If
        End Function
    End Class
End Namespace
