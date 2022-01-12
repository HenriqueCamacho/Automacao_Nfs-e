Imports CClass
Imports System.Security.Cryptography.X509Certificates
Imports CertificadoDigital
Imports NFSeCore

Namespace cCenti2
    Public Class PropriedadeCertificado
        Inherits APropriedadeCertificado

        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub

        Public Overrides Function assinarXMLRps(ByVal xml As String, _
                             ByVal idEmpresa As Integer, _
                             ByVal codgCidade As Integer, Optional ByVal assinaRps As Boolean = True, Optional ByVal gerarKeyName As Boolean = False, Optional ByVal usarUriRef As Boolean = True) _
                             As String
            Return assinarXML1(Nothing, xml, idEmpresa, codgCidade, False, assinaRps, True, usarUriRef)
        End Function

        Private Function assinarXML1(ByRef certificado As X509Certificate, _
                                ByVal xml As String, _
                                  ByVal idEmpresa As Integer, _
                                  ByVal codgCidade As Integer, _
                                  ByVal isCancelamento As Boolean, _
                                  Optional ByVal assinaRps As Boolean = True, Optional ByVal gerarKeyName As Boolean = False, Optional ByVal usarUriRef As Boolean = True) _
                                  As String
            Dim empresa As NfseEmpresasData = GetIntegracaoEmpresa(idEmpresa)
            Dim propXml As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgCidade)
            Dim tpCertificado As TipoCertificado = CInt(ParametroUtil.GetParametro(idEmpresa, "TIPO_CERT_ASSINATURA"))
            Dim usaCertEspecifico As Boolean = ParametroUtil.getParametroCheck(idEmpresa, "USAR_CERT_ESPEC_ASSINATURA")
            Dim numrDocumento As Long = empresa.numrCnpj
            Dim tipoDocumento As Integer = CertificadoUtil.CNPJ
            Dim ambiente As TipoAmbiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")

            If usaCertEspecifico Then
                tipoDocumento = CInt(ParametroUtil.GetParametro(idEmpresa, "TIPO_CERT_ASSINATURA"))
                numrDocumento = CLng(ParametroUtil.GetParametro(idEmpresa, "NUMR_DOC_CERT_ASSINATURA"))
            End If
            LogInfo("Certificado numer:" & numrDocumento)
            LogInfo("Certificado tipo:" & tipoDocumento)
            LogInfo("Certificado usa:" & usaCertEspecifico)
            'LogInfo("Certificado tp:" + tpCertificado)

            If certificado Is Nothing Then
                LogInfo("Certificado esta vazio")
            End If

            If tpCertificado = TipoCertificado.A1 Then
                If Not isCancelamento Then
                    ' assinar o rps
                    If assinaRps Then
                        LogInfo("cCenti2.PropriedadeCertificado.assinarXML1(" & numrDocumento & ", " & xml & ", " & propXml.GetFirstTagRps & ", " & propXml.GetSecondTagRps)

                        If certificado Is Nothing Then
                            certificado = CertificadoUtil.obterCertificadoPorCnpj(numrDocumento)
                        End If

                        If certificado Is Nothing Then
                            LogInfo("########## Não encontrado certificado para o CNPJ " & numrDocumento)
                        End If

                        xml = CertificadoUtil.assinarXML(certificado, xml, propXml.GetFirstTagRps, propXml.GetSecondTagRps, False, False)
                        'xml = CertificadoUtil.assinarXML(CertificadoUtil.obterCertificadoPorCnpj(numrDocumento), xml, propXml.GetFirstTagRps, propXml.GetSecondTagRps, False, False)

                    End If
                Else
                    If certificado Is Nothing Then
                        ' assinar o cancelamento
                        xml = CertificadoUtil.assinarXML(numrDocumento, xml, propXml.GetFTagCanc, propXml.GetSTagCanc, False, True)
                    Else
                        ' assinar o cancelamento
                        xml = CertificadoUtil.assinarXML(certificado, xml, propXml.GetFTagCanc, propXml.GetSTagCanc, False, True)
                    End If
                End If
            Else
                If Not isCancelamento Then
                    ' assinar o rps
                    If assinaRps Then
                        SyncLock SemafaroA3.ocupado
                            If Not certificadoA3.IsOpen Then
                                certificadoA3.Open(CertificadoUtil.CNPJ, numrDocumento)
                                certificadoA3.OpenSession()
                            End If

                            xml = certificadoA3.assinarXML(xml, propXml.GetXPathFTagRps, propXml.GetXPathSTagRps, propXml.GetNamespaceMap())
                            ' os componentes de xml usados para assinar usando A3 
                            ' devolvem o string com encoding UTF-16, por isso o replace aqui
                            xml = xml.Replace("utf-16", "UTF-8")
                        End SyncLock
                    End If

                    SyncLock SemafaroA3.ocupado
                        If Not certificadoA3.IsOpen Then
                            certificadoA3.Open(CertificadoUtil.CNPJ, numrDocumento)
                            certificadoA3.OpenSession()
                        End If
                        ' assinar o lote
                        xml = certificadoA3.assinarXML(xml, propXml.GetXPathFTagLote(), propXml.GetXPathSTagLote(), propXml.GetNamespaceMap())

                        xml = xml.Replace("utf-16", "UTF-8")
                    End SyncLock
                Else
                    SyncLock SemafaroA3.ocupado
                        If Not certificadoA3.IsOpen Then
                            certificadoA3.Open(CertificadoUtil.CNPJ, numrDocumento)
                            certificadoA3.OpenSession()
                        End If

                        ' assinar o cancelamento
                        xml = certificadoA3.assinarXML(xml, propXml.GetXPathFTagCanc(), propXml.GetXPathSTagCanc(), _
                                    propXml.GetNamespaceMap())

                        ' os componentes de xml usados para assinar usando A3 
                        ' devolvem o string com encoding UTF-16, por isso o replace aqui
                        xml = xml.Replace("utf-16", "UTF-8")
                    End SyncLock
                End If
            End If
            xml = xml.Replace("Rps1", "Rps")

            Return xml
        End Function
    End Class
End Namespace
