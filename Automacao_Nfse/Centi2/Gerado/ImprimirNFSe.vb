Imports CClass
Imports CClass.comum
Imports GrapeCity.ActiveReports.Document
Imports GrapeCity.ActiveReports.Extensibility
Imports System.Xml
Imports System.Text.RegularExpressions
Imports NFSeCore
Imports System.Windows.Forms
Imports System.Net

Namespace cCenti3
    Public Class ImprimirNFSe
        Inherits AImprimirNFSe

        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub

        Public Overrides Sub imprimir(ByVal codgCidade As Integer, _
                                      ByVal movimentoData As NfseMovimentoData, _
                                      ByVal tpImpressao As TipoImpressao)
            Dim tpSistema As TipoSistema = CInt(ParametroUtil.GetParametro(ParametroUtil.PARAM_GERAL, "TIPO_SISTEMA"))

            If tpSistema = TipoSistema.DELPHI Then
                imprimirDelphi(codgCidade, movimentoData, tpImpressao)
            Else
                imprimirWeb(codgCidade, movimentoData, tpImpressao)
            End If
        End Sub

        Protected Overrides Sub imprimirRps(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData)
            Dim document As SectionDocument = GetDocumentRps(codgCidade, movimentoData)

            document.Printer.PrinterName = movimentoData.nomeImpressora
            GrapeCity.ActiveReports.PrintExtension.Print(document, False, False)
        End Sub

        Private Function GetDocumentRps(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData) As SectionDocument
            ' obter uma instancia do datatable
            Dim dtRps As DataTable = ImpressaoUtil.criarTblImpressaoRPS()
            Dim row As DataRow = dtRps.NewRow()
            ' gerar um objeto xml com os dados do rps
            Dim rps As New XmlDocument()
            Dim prestador As XmlDocument = Nothing
            Dim tomador As XmlDocument = Nothing
            Dim servico As XmlDocument = Nothing
            Dim servicoData As NfseServicoData = Nothing
            Dim tomadorData As NfseTomadorData = Nothing
            Dim empresaData As NfseEmpresasData = Nothing
            Dim document As SectionDocument = Nothing
            Dim report As GrapeCity.ActiveReports.SectionReport = Nothing

            Try
                LogInfo("GetDocumentRps Inicio")
                rps.LoadXml(movimentoData.xmlNota)
                prestador = XmlUtil.getDocByTag(rps, "Prestador")
                tomador = XmlUtil.getDocByTag(rps, "Tomador")
                If prestador Is Nothing Or tomador Is Nothing Then
                    prestador = XmlUtil.getDocByTag(rps, "PrestadorServico")
                    tomador = XmlUtil.getDocByTag(rps, "TomadorServico")
                End If
                servico = XmlUtil.getDocByTag(rps, "Servico")
                servicoData = GetIntegracaoServico(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                tomadorData = GetIntegracaoTomador(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                empresaData = GetIntegracaoEmpresa(CInt(movimentoData.idEmpresa))

                ' preencher o datatable com os dados necessarios
                row("NUMR_RPS") = servicoData.numeroRps.ToString
                row("DATA_EMISSAO") = XmlUtil.getValorTag(rps, "DataEmissao").ToString()
                row("CODIGO_VERIFICACAO") = XmlUtil.getValorTag(rps, "CodigoVerificacao")
                Try
                    row("SERIE_RPS") = XmlUtil.getValorTag(rps, "Serie")
                Catch ex As Exception
                    row("SERIE_RPS") = ""
                End Try

                If XmlUtil.getValorTag(prestador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_PRESTADOR") = CLng(XmlUtil.getValorTag(prestador, "Cnpj").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("00\.000\.000/0000-00")
                Else
                    row("CNPJ_CPF_PRESTADOR") = CLng(XmlUtil.getValorTag(prestador, "Cpf").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("000\.000\.000-00")
                End If

                row("INSC_MUNC_PRESTADOR") = XmlUtil.getValorTag(prestador, "InscricaoMunicipal")
                row("RAZAO_SOCIAL_PRESTADOR") = empresaData.nomeEmpresa
                row("ENDERECO_PRESTADOR") = empresaData.endereco
                row("NUM_END_PRESTADOR") = empresaData.numrEnd
                row("BAIRRO_END_PRESTADOR") = empresaData.bairroEnd
                row("CEP_END_PRESTADOR") = empresaData.cepEnd
                Try
                    row("EMAIL_PRESTADOR") = empresaData.email.ToLower
                Catch ex As Exception
                    row("EMAIL_PRESTADOR") = ""
                End Try
                row("MUNICIPIO_PRESTADOR") = empresaData.municipio
                row("UF_PRESTADOR") = empresaData.uf
                row("RAZAO_SOCIAL_TOMADOR") = XmlUtil.getValorTag(tomador, "RazaoSocial")

                If XmlUtil.getValorTag(tomador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_TOMADOR") = CLng(XmlUtil.getValorTag(tomador, "Cnpj").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("00\.000\.000/0000-00")
                Else
                    row("CNPJ_CPF_TOMADOR") = CLng(XmlUtil.getValorTag(tomador, "Cpf").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("000\.000\.000-00")
                End If

                row("INSC_MUNC_TOMADOR") = tomadorData.inscMunicipal
                row("ENDERECO_TOMADOR") = tomadorData.endereco
                row("NUM_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Numero")
                row("BAIRRO_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Bairro")
                row("CEP_END_TOMADOR") = tomadorData.cep
                row("MUNICIPIO_TOMADOR") = tomadorData.nomeMunicipio
                row("UF_TOMADOR") = tomadorData.uf
                row("DISCRIMINACAO") = XmlUtil.getValorTag(servico, "Discriminacao").Replace("|", Chr(13) + Chr(10))
                Try
                    row("EMAIL_TOMADOR") = tomadorData.email
                Catch ex As Exception
                    row("EMAIL_TOMADOR") = ""
                End Try
                Try
                    row("VALOR_SERVICO") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_SERVICO") = XmlUtil.getValorTag(servico, "ValorServicos")
                End Try

                Try
                    row("VALOR_DEDUCOES") = servicoData.valrDeducoes.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_DEDUCOES") = servicoData.valrDeducoes.ToString
                End Try

                Try
                    row("BASE_CALCULO") = servicoData.baseCalculo.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("BASE_CALCULO") = servicoData.baseCalculo.ToString
                End Try

                Try
                    row("ALIQUOTA") = CDec(servicoData.aliquota.ToString.Replace(".", ","))
                Catch ex As Exception
                    row("ALIQUOTA") = servicoData.aliquota.ToString
                End Try

                Try
                    row("VALOR_ISS") = servicoData.valrIss.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_ISS") = servicoData.valrIss.ToString
                End Try

                Try
                    row("VALR_DESCONTOS") = servicoData.valrDescIncondicionado.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALR_DESCONTOS") = servicoData.valrDescIncondicionado
                End Try

                Try
                    row("valr_liquido_nfse") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ","))
                Catch ex As Exception
                    row("valr_liquido_nfse") = XmlUtil.getValorTag(servico, "ValorServicos")
                End Try

                Try
                    row("VALOR_PIS") = servicoData.valrPis.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_PIS") = servicoData.valrPis
                End Try

                Try
                    row("VALOR_COFINS") = servicoData.valrCofins.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_COFINS") = servicoData.valrCofins
                End Try

                Try
                    row("VALOR_INSS") = servicoData.valrInss.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_INSS") = servicoData.valrInss
                End Try
                LogInfo("GetDocumentRps 09")
                Try
                    row("VALOR_IRRF") = servicoData.valrIr.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_IRRF") = servicoData.valrIr
                End Try

                Try
                    row("VALOR_CSLL") = servicoData.valrCsll.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_CSLL") = servicoData.valrCsll
                End Try

                Try
                    row("VALOR_OUTRAS_RETENCOES") = servicoData.outrasRetencoes.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_OUTRAS_RETENCOES") = servicoData.outrasRetencoes
                End Try

                If XmlUtil.getValorTag(servico, "IssRetido") = "1" Then 'é retido
                    Try
                        row("VALOR_ISSQN") = CDec(XmlUtil.getValorTag(servico, "ValorIss").Replace(".", ","))
                    Catch ex As Exception
                        row("VALOR_ISSQN") = XmlUtil.getValorTag(servico, "ValorIss")
                    End Try
                Else
                    row("VALOR_ISSQN") = 0
                End If

                Try
                    row("COD_DESC_SERVICO") = XmlUtil.getValorTag(servico, "ItemListaServico") + " - " + GetAtividadeDescricao(CInt(XmlUtil.getValorTag(servico, "ItemListaServico")))
                Catch ex As Exception
                    row("COD_DESC_SERVICO") = servicoData.codgServico + " - " + GetAtividadeDescricao(servicoData.codgServico)
                End Try
                row("VALOR_CREDITO") = DBNull.Value
                row("OUTRAS_OBSERVACAO") = ""
                row("VALOR_TOTAL_NOTA") = servicoData.valrServicos - servicoData.valrDescIncondicionado



                SetDadosRelatorio(row, movimentoData)
                dtRps.Rows.Add(row)
                report = GetReportRps(movimentoData)
                report.DataSource = dtRps
                document = report.Document

                report.Run()

                Return document
            Finally
                rps = Nothing
                prestador = Nothing
                tomador = Nothing
                servico = Nothing
                servicoData = Nothing
                tomadorData = Nothing
                empresaData = Nothing
                report = Nothing
            End Try
        End Function

        Protected Overrides Sub imprimirNfse(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData)
            Dim document As SectionDocument = Nothing
            document = GetDocumentNfse(codgCidade, movimentoData)

            document.Printer.PrinterName = movimentoData.nomeImpressora
            GrapeCity.ActiveReports.PrintExtension.Print(document, False, False)
        End Sub

        Public Overrides Function GetDocumentNfse(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData) As SectionDocument
            ' obter uma instancia do datatable
            Dim dtNfse As DataTable = ImpressaoUtil.criarTblImpressaoNFSe()
            Dim row As DataRow = dtNfse.NewRow()
            ' gerar um objeto xml com os dados do rps
            Dim nfse As New XmlDocument()
            Dim prestador As XmlDocument = Nothing
            Dim tomador As XmlDocument = Nothing
            Dim servico As XmlDocument = Nothing
            Dim valores As XmlDocument = Nothing
            Dim servicoData As NfseServicoData = Nothing
            Dim tomadorData As NfseTomadorData = Nothing
            Dim empresaData As NfseEmpresasData = Nothing
            Dim document As SectionDocument = Nothing
            Dim report As GrapeCity.ActiveReports.SectionReport = Nothing
            Dim cidadeData As NfseCidadesData = Nothing
            'NAO IRA GERAR UM DOCUMENTO NFSE LEGITIMO
            'O RETORNO DESSA PREFEITURA EH HORRIVEL, NAO TEM DADOS O SUFICIENTE PARA GERAR NFSE 
            'ALGUMS CLIENTES ESTAVAM RECLAMANDO QUE ESTAVA IMPRIMINDO SO RPS            
            Try
                nfse.LoadXml(movimentoData.xmlNota)
                LogInfo("GetDocumentNFSe Inicio")
                nfse.LoadXml(movimentoData.xmlNota)
                prestador = XmlUtil.getDocByTag(nfse, "Prestador")
                tomador = XmlUtil.getDocByTag(nfse, "Tomador")
                valores = XmlUtil.getDocByTag(nfse, "ValoresNfse")
                If prestador Is Nothing Or tomador Is Nothing Then
                    prestador = XmlUtil.getDocByTag(nfse, "PrestadorServico")
                    tomador = XmlUtil.getDocByTag(nfse, "TomadorServico")
                End If
                servico = XmlUtil.getDocByTag(nfse, "Servico")
                servicoData = GetIntegracaoServico(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                tomadorData = GetIntegracaoTomador(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                empresaData = GetIntegracaoEmpresa(CInt(movimentoData.idEmpresa))

                ' preencher o datatable com os dados necessarios
                row("NUMR_RPS") = servicoData.numeroRps.ToString
                row("NUMR_NFSE") = XmlUtil.getValorTag(nfse, "Numero")
                row("DATA_EMISSAO") = XmlUtil.getValorTag(nfse, "DataEmissao").ToString()
                row("CODIGO_VERIFICACAO") = XmlUtil.getValorTag(nfse, "CodigoVerificacao")
                Try
                    row("SERIE_RPS") = XmlUtil.getValorTag(nfse, "Serie")
                Catch ex As Exception
                    row("SERIE_RPS") = ""
                End Try

                If XmlUtil.getValorTag(prestador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_PRESTADOR") = CLng(XmlUtil.getValorTag(prestador, "Cnpj").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("00\.000\.000/0000-00")
                Else
                    row("CNPJ_CPF_PRESTADOR") = CLng(XmlUtil.getValorTag(prestador, "Cpf").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("000\.000\.000-00")
                End If

                'INSC_EST_TOMADOR
                row("INSC_MUNC_PRESTADOR") = empresaData.inscricaoMunicipal
                row("INSC_EST_TOMADOR") = tomadorData.inscricaoEstadual

                row("RAZAO_SOCIAL_PRESTADOR") = empresaData.nomeEmpresa
                row("ENDERECO_PRESTADOR") = empresaData.endereco
                row("NUM_END_PRESTADOR") = empresaData.numrEnd
                row("BAIRRO_END_PRESTADOR") = empresaData.bairroEnd
                row("CEP_END_PRESTADOR") = empresaData.cepEnd
                Try
                    row("EMAIL_PRESTADOR") = empresaData.email.ToLower
                Catch ex As Exception
                    row("EMAIL_PRESTADOR") = ""
                End Try
                row("MUNICIPIO_PRESTADOR") = empresaData.municipio
                row("UF_PRESTADOR") = empresaData.uf
                row("RAZAO_SOCIAL_TOMADOR") = XmlUtil.getValorTag(tomador, "RazaoSocial")

                If XmlUtil.getValorTag(tomador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_TOMADOR") = CLng(XmlUtil.getValorTag(tomador, "Cnpj").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("00\.000\.000/0000-00")
                Else
                    row("CNPJ_CPF_TOMADOR") = CLng(XmlUtil.getValorTag(tomador, "Cpf").Replace(".", "").Replace("-", "").Replace("/", "")).ToString("000\.000\.000-00")
                End If

                row("INSC_MUNC_TOMADOR") = tomadorData.inscMunicipal
                row("ENDERECO_TOMADOR") = tomadorData.endereco
                row("NUM_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Numero")
                row("BAIRRO_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Bairro")
                row("CEP_END_TOMADOR") = tomadorData.cep
                row("MUNICIPIO_TOMADOR") = tomadorData.nomeMunicipio
                row("UF_TOMADOR") = tomadorData.uf
                row("DISCRIMINACAO") = XmlUtil.getValorTag(servico, "Discriminacao").Replace("|", Chr(13) + Chr(10))
                Try
                    row("EMAIL_TOMADOR") = tomadorData.email
                Catch ex As Exception
                    row("EMAIL_TOMADOR") = ""
                End Try
                Try
                    row("VALOR_SERVICO") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_SERVICO") = XmlUtil.getValorTag(servico, "ValorServicos")
                End Try

                Try
                    row("VALOR_DEDUCOES") = servicoData.valrDeducoes.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_DEDUCOES") = servicoData.valrDeducoes.ToString
                End Try

                Try
                    row("BASE_CALCULO") = servicoData.baseCalculo.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("BASE_CALCULO") = servicoData.baseCalculo.ToString
                End Try

                Try
                    row("ALIQUOTA") = CDec(servicoData.aliquota.ToString.Replace(".", ","))
                Catch ex As Exception
                    row("ALIQUOTA") = servicoData.aliquota.ToString
                End Try

                Dim valorIss As Decimal
                If servicoData.issRetido = TipoIss.RETIDO Then
                    valorIss = servicoData.valrIssRetido
                Else
                    valorIss = servicoData.valrIss
                End If

                Try
                    row("VALOR_ISS") = valorIss.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_ISS") = valorIss.ToString()
                End Try

                Try
                    row("VALR_DESCONTOS") = servicoData.valrDescIncondicionado.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALR_DESCONTOS") = servicoData.valrDescIncondicionado
                End Try

                Try
                    row("valr_liquido_nfse") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ",")) - servicoData.valrIr
                Catch ex As Exception
                    row("valr_liquido_nfse") = servicoData.valrLiquidoNfse - servicoData.valrIr
                End Try

                Try
                    row("VALOR_PIS") = CDec(XmlUtil.getValorTag(valores, "ValorPis").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_PIS") = XmlUtil.getValorTag(valores, "ValorPis")
                End Try

                Try
                    row("VALOR_COFINS") = CDec(XmlUtil.getValorTag(valores, "ValorCofins").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_COFINS") = XmlUtil.getValorTag(valores, "ValorCofins")
                End Try

                Try
                    row("VALOR_INSS") = servicoData.valrInss.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_INSS") = servicoData.valrInss
                End Try
                LogInfo("GetDocumentRps 09")
                Try
                    row("VALOR_IRRF") = servicoData.valrIr.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_IRRF") = servicoData.valrIr
                End Try

                Try
                    row("VALOR_CSLL") = servicoData.valrCsll.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_CSLL") = servicoData.valrCsll
                End Try

                Try
                    row("VALOR_OUTRAS_RETENCOES") = servicoData.outrasRetencoes.ToString.Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_OUTRAS_RETENCOES") = servicoData.outrasRetencoes
                End Try

                If XmlUtil.getValorTag(servico, "IssRetido") = "1" Then 'é retido
                    Try
                        row("VALOR_ISSQN") = CDec(XmlUtil.getValorTag(servico, "ValorIss").Replace(".", ","))
                    Catch ex As Exception
                        row("VALOR_ISSQN") = XmlUtil.getValorTag(servico, "ValorIss")
                    End Try
                Else
                    row("VALOR_ISSQN") = 0
                End If

                'Try
                'row("COD_DESC_SERVICO") = XmlUtil.getValorTag(servico, "ItemListaServico") + " - " + GetAtividadeDescricao(CInt(XmlUtil.getValorTag(servico, "ItemListaServico")))
                'Catch ex As Exception
                row("COD_DESC_SERVICO") = servicoData.codgServico + " - " + GetAtividadeDescricao(servicoData.codgServico)
                'End Try
                row("VALOR_CREDITO") = DBNull.Value
                row("OUTRAS_OBSERVACAO") = "Numero RPS.: " + servicoData.numeroRps.ToString + ";"
                row("VALOR_TOTAL_NOTA") = servicoData.valrServicos - servicoData.valrDescIncondicionado

                report = New NfseReport()

                Dim destaque As String

                Try
                    destaque = ParametroUtil.GetParametro(CInt(movimentoData.idEmpresa), "DESTAQUE_IMP_NFSE_CLIENTE")
                Catch ex As Exception
                    destaque = "N"
                End Try

                If destaque = "S" Then
                    CType(report, NfseReport).lblTitulo1Cliente.Visible = True
                    CType(report, NfseReport).lblTitulo2Cliente.Visible = True
                    CType(report, NfseReport).lblTitulo3Cliente.Visible = True
                    CType(report, NfseReport).lblDataRecebimentoCliente.Visible = True
                    CType(report, NfseReport).lblIdentificacaoAssinaturaCliente.Visible = True
                    CType(report, NfseReport).lblNfseCliente.Visible = True
                    CType(report, NfseReport).txbNumeroNotaCliente.Visible = True
                    CType(report, NfseReport).shapeDestaqueCliente.Visible = True
                    CType(report, NfseReport).line1DestaqueCliente.Visible = True
                    CType(report, NfseReport).line2DestaqueCliente.Visible = True
                    CType(report, NfseReport).line3DestaqueCliente.Visible = True
                End If

                'Imprimir o campo valor liquido
                Dim mostrarValrLiq As String

                Try
                    mostrarValrLiq = ParametroUtil.GetParametro(CInt(movimentoData.idEmpresa), "VALR_LIQ_VISIBLE")
                Catch ex As Exception
                    mostrarValrLiq = "S"
                End Try

                If mostrarValrLiq = "N" Then
                    CType(report, NfseReport).Label37.Visible = False
                    CType(report, NfseReport).TextBox34.Visible = False
                    row("OUTRAS_OBSERVACAO") = "RETENÇÃO DE IMPOSTOS POR CONTA DO DESTINATÁRIO!"
                Else
                    CType(report, NfseReport).Label37.Visible = True
                    CType(report, NfseReport).TextBox34.Visible = True
                End If

                dtNfse.Rows.Add(row)

                report.DataSource = dtNfse
                document = report.Document

                ' nao é obrigatorio o uso de logomarca por parte da empresa... 
                ' mas nao custa nada avisar que ele nao esta usando
                Try
                    CType(report, NfseReport).LogoEmpresa.Image = ImpressaoUtil.CarregarImagem(ParametroUtil.GetParametro(movimentoData.idEmpresa, "LOGOMARCA"))
                Catch ex As Exception
                    MensagemUtil.GravarMsg(conn, trans, factory, publicVar, movimentoData.idMovimento, "AVISO: nao foi possivel usar logo marca na impressão: " + ex.Message)
                End Try

                '---Inicio logomarca prefeitura
                Dim logoPrefeitura As String = ""

                Try
                    logoPrefeitura = ParametroUtil.GetParametro(movimentoData.idEmpresa, "BRASAO_PREFEITURA")
                Catch ex As Exception
                    logoPrefeitura = ""
                End Try

                If logoPrefeitura <> "" Then
                    Try
                        CType(report, NfseReport).logoPrefeitura.Image = ImpressaoUtil.CarregarImagem(logoPrefeitura)
                    Catch ex As Exception
                        MensagemUtil.GravarMsg(conn, trans, factory, publicVar, movimentoData.idMovimento, "AVISO: nao foi possivel usar logo marca da prefeitura na impressão: " + ex.Message)
                    End Try
                End If
                '---Fim logomarca prefeitura

                ' CType(report, NfseReport).lblTituloNfse.Text = "RECIBO PROVISÓRIO DE SERVIÇO - RPS"

                'quando o status do movimento for cancelado, aparecerá uma marca d'água de "Cancelado"
                If movimentoData.status = StatusProcessamento.CANCELADA Then
                    Try
                        CType(report, NfseReport).imCancelado.Visible = True
                    Catch ex As Exception
                        MensagemUtil.GravarMsg(conn, trans, factory, publicVar, movimentoData.idMovimento, "AVISO: nao foi possivel acrescentar marca d'agua 'Cancelado': " + ex.Message)
                    End Try
                End If

                '---Inicio titulo e subtitulo
                Dim titulo As String = ""
                Dim subtitulo As String = ""

                Try
                    titulo = ParametroUtil.GetParametro(movimentoData.idEmpresa, "TITULO_IMPRESSAO")
                Catch ex As Exception
                    titulo = ""
                End Try

                Try
                    subtitulo = ParametroUtil.GetParametro(movimentoData.idEmpresa, "SUBTITULO_IMPRESSAO")
                Catch ex As Exception
                    subtitulo = ""
                End Try

                CType(report, NfseReport).lblNomePrefeitura.Text = titulo
                CType(report, NfseReport).lblSubTitulo.Text = subtitulo
                '---Fim titulo e subtitulo

                report.Run()

                Return document
            Finally
                nfse = Nothing
                prestador = Nothing
                tomador = Nothing
                servico = Nothing
                servicoData = Nothing
                tomadorData = Nothing
                empresaData = Nothing
                report = Nothing
            End Try
        End Function

        Public Overrides Function GetDocument(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData, ByVal tpImpressao As TipoImpressao)
            LogInfo("GetDocument Codg. Cidade: " & codgCidade)

            If tpImpressao = TipoImpressao.NFSE Then
                Return GetDocumentNfse(codgCidade, movimentoData)
            Else
                Return GetDocumentRps(codgCidade, movimentoData)
            End If

        End Function

    End Class
End Namespace
