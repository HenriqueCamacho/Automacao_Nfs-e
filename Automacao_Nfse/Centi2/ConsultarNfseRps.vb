Imports CClass
Imports NFSeCore
Imports System.Xml
Imports System.Web.Services.Protocols
Imports System.Security.Cryptography.X509Certificates
Imports System.Text.RegularExpressions
Imports System.IO
Imports CertificadoDigital
Imports System.Text

Namespace cCenti2
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
            Dim usuario As String = ""
            Dim senha As String = ""
            Me.conn = conn
            Me.trans = trans

            Try

                ambiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")
                propXml = FactoryCore.GetPropriedadesXML(empresaData.codgCidade)

                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = "https://api.centi.com.br/nfe/consultar/homologacao/rps/GO/rioverde"
                    usuario = "fiscal@carpal.com.br"
                    senha = "carpal@2016"
                Else
                    url = "https://api.centi.com.br/nfe/consultar/" + empresaData.uf.ToUpper + "/" + empresaData.municipio.Replace(" ", "").ToLower
                    usuario = ParametroUtil.GetParametro(empresaData.idEmpresa, "USUARIO")
                    senha = ParametroUtil.GetParametro(empresaData.idEmpresa, "SENHA")
                End If

                Dim xmlConsultaPorRpsEnvio As New StringBuilder()

                'xmlConsultaPorRpsEnvio.Append("<?xml version=""1.0"" encoding=""UTF-8""?>")
                xmlConsultaPorRpsEnvio.Append("<ConsultarNfseRpsEnvio>")
                xmlConsultaPorRpsEnvio.Append("<IdentificacaoRps>")
                xmlConsultaPorRpsEnvio.Append("<Numero>" & numeroRps & "</Numero>")
                xmlConsultaPorRpsEnvio.Append("<Serie>" & serieRps & "</Serie>")
                xmlConsultaPorRpsEnvio.Append("<Tipo>" & tipoRps & "</Tipo>")
                xmlConsultaPorRpsEnvio.Append("</IdentificacaoRps>")
                xmlConsultaPorRpsEnvio.Append("<Prestador>")
                xmlConsultaPorRpsEnvio.Append("<Cnpj>" & empresaData.numrCnpj & "</Cnpj>")
                xmlConsultaPorRpsEnvio.Append("<InscricaoMunicipal>" & empresaData.inscricaoMunicipal & "</InscricaoMunicipal>")
                xmlConsultaPorRpsEnvio.Append("</Prestador>")
                xmlConsultaPorRpsEnvio.Append("</ConsultarNfseRpsEnvio>")


                Dim jsonEnvio As String = "{""usuario"": """ + usuario + """, ""senha"": """ + senha + """, ""xml"": """ + xmlConsultaPorRpsEnvio.ToString().Replace("""", "'") + """}"

                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, empresaData.codgCidade)
                certificado = propCertificado.GetCertificadoTransmissao(idEmpresa, numrCnpj)

                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para CONSULTA DE NFSE POR RPS, através do nº do documento informado : " + numrCnpj.ToString)
                End If


                xmlRetorno = HTTPPostRequestJson(url, jsonEnvio, idEmpresa, empresaData.numrCnpj)

                xmlRetorno = (CClass.StringUtil.tirarAcentos(xmlRetorno))
                xmlRetorno = xmlRetorno.Replace("<?xml version=""1.0"" encoding=""utf-8""?>", "")

                result = xmlRetorno

            Catch ex As Exception
                Throw ex
            End Try

            Return result
        End Function

        Public Overrides Sub gravarXmlNfse(ByVal conn As IDbConnection, ByVal trans As IDbTransaction, ByVal xmlNfse As String, ByVal codgMunicipio As Integer, ByVal idEmpresa As Integer)
            Dim numeroRps As Integer = 0
            Dim serieRps As String = 0
            Dim tipoRps As Integer = 0
            Dim numrCnpjPrestador As Long = 0
            Dim xmldoc As XmlDocument = Nothing
            Dim xmlIdentificacaoRps As XmlDocument = Nothing
            Dim movimentoRn As NfseMovimentoRN = Nothing
            Dim dtsNfse As DtsNfse = Nothing
            Dim ambiente As TipoAmbiente = TipoAmbiente.HOMOLOGACAO
            Dim movimentoData As NfseMovimentoData = Nothing
            Dim prestador As XmlDocument = Nothing
            ambiente = CInt(ParametroUtil.GetParametro(idEmpresa, "AMBIENTE"))
            Try
                ' obter o document xml 
                xmldoc = New XmlDocument()
                xmldoc.LoadXml(xmlNfse)

                ' obter o document da tag identificacaorps
                xmlIdentificacaoRps = XmlUtil.getDocByTag(xmldoc, "IdentificacaoRps")
                'Identificação do Prestador
                prestador = XmlUtil.getDocByTag(xmldoc, "PrestadorServico")

                ' obter o numero do RPS

                If XmlUtil.getValorTag(xmlIdentificacaoRps, "Numero").Trim.Length > 0 Then
                    numeroRps = CInt(XmlUtil.getValorTag(xmlIdentificacaoRps, "Numero"))
                Else
                    Throw New FalhaException("Erro ao obter o número do RPS no xml informado")
                End If

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
                If XmlUtil.getValorTag(prestador, "Cnpj").Trim.Length > 0 Then
                    numrCnpjPrestador = CLng(XmlUtil.getValorTag(prestador, "Cnpj"))
                ElseIf XmlUtil.getValorTag(prestador, "Cpf").Trim.Length > 0 Then
                    numrCnpjPrestador = CLng(XmlUtil.getValorTag(prestador, "Cpf"))
                Else
                    Throw New FalhaException("Erro ao obter o número do CNPJ do prestador no xml informado")
                End If

                ' obter o movimento do RPS                
                dtsNfse = New DtsNfse()
                movimentoRn = New NfseMovimentoRN(factory, publicVar, dtsNfse)

                movimentoRn.filtro = " TIPO_RPS = :TIPO_RPS AND SERIE_RPS = :SERIE_RPS AND NUMERO_RPS = :NUMERO_RPS AND NUMR_CNPJ = :NUMR_CNPJ AND TIPO_AMBIENTE = :TIPO_AMBIENTE"
                movimentoRn.params.Add(New ParamDB("TIPO_RPS", tipoRps, DbType.Int64))
                movimentoRn.params.Add(New ParamDB("SERIE_RPS", serieRps, DbType.AnsiString))
                movimentoRn.params.Add(New ParamDB("NUMERO_RPS", numeroRps, DbType.Int64))
                movimentoRn.params.Add(New ParamDB("NUMR_CNPJ", numrCnpjPrestador, DbType.Int64))
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
                    Throw New FalhaException("Erro ao gravar NFS-e: " + ex.Message, ex)
                End If
            End Try
        End Sub

    End Class
End Namespace
