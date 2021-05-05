Imports System.IO
Imports CrystalDecisions.Shared
Imports SAPbobsCOM

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports
Namespace SAP.ImpresionEtiquetas
    Public Class CRPrinting
        Public Shared Property GetError As String

        Protected Friend Shared Function ExecutePDFReportHANA(oCompany As SAPbobsCOM.Company, ByRef Name As String, ByVal LayoutCode As String, ByVal ParamName As String(), ByVal ParamValue As String(), Optional ByVal PathCR As String = "", Optional ByVal defaultPrinterSetting As System.Drawing.Printing.PrinterSettings = Nothing, Optional LOG As EXO_Log.EXO_Log = Nothing) As String

            Dim oCRReport As CrystalDecisions.CrystalReports.Engine.ReportDocument = Nothing
            Dim ExportCR As SAP.ImpresionEtiquetas.ExportCRHelper = Nothing
            Dim File As String = ""
            Dim eDriver As String = "{B1CRHPROXY}"
            Dim Cripto As New EXO_DIAPI.EXO_Cripto()
            Dim rs As SAPbobsCOM.Recordset

            'Dim usuarioHANA As String = System.Configuration.ConfigurationManager.AppSettings("usuarioHANA")
            'Dim pwdHANA As String = System.Configuration.ConfigurationManager.AppSettings("pwdHANA")


            Dim BDUS As String = System.Configuration.ConfigurationManager.AppSettings("usuarioHANA")
            Dim BDPW As String = System.Configuration.ConfigurationManager.AppSettings("pwdHANA")

            Try


                'ObjGlobal.SBOApp.StatusBar.SetText("Preparando información de impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

                'UDO_Config = New EXO_DIAPI.EXO_UDOEntity(ObjGlobal.refDi.comunes, "EXO_OGEN")
                'If UDO_Config.GetByKey("EXO_KERNEL") = False Then
                '    ObjGlobal.SBOApp.MessageBox("No se encontró la configuración de conexión!")
                '    Return ""
                'End If

                'Dim sql As String = "SELECT ""U_EXO_BDUS"",""U_EXO_BDPW"" FROM ""@EXO_OGEN"""

                'rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'rs.DoQuery(sql)

                'If rs.RecordCount > 0 Then
                '    rs.MoveFirst()
                '    BDUS = rs.Fields.Item("U_EXO_BDUS").Value.ToString
                '    BDPW = rs.Fields.Item("U_EXO_BDPW").Value.ToString
                'End If

                'Descargar el fichero en la Temp
                If PathCR = "" Then
                    LOG.escribeMensaje("Descargando impreso en TMP...Espere por favor")
                    ExportCR = New SAP.ImpresionEtiquetas.ExportCRHelper(oCompany)
                    File = ExportCR.ExportReport(LayoutCode)
                Else
                    File = PathCR
                End If
                If File = "" Then
                    LOG.escribeMensaje("No se pudo recuperar el impreso por defecto!")
                    Return "No se pudo recuperar el impreso por defecto!"
                End If

                LOG.escribeMensaje("Cargando impreso por defecto...Espere por favor")
                Try
                    oCRReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
                    oCRReport.Load(File)
                    LOG.escribeMensaje("Aplicando datos de conexión al impreso por defecto...Espere por favor")
                    If ApplySetConnection(oCompany, oCRReport, eDriver, Split(oCompany.Server, ":")(0), oCompany.CompanyDB, BDUS, BDPW, Split(oCompany.Server, ":")(1), LOG) = False Then
                        Return ""
                    End If
                    'Cripto.desencripta(BDPW)
                    If ApplyNewServerNew(oCompany, oCRReport, eDriver, Split(oCompany.Server, ":")(0), oCompany.CompanyDB, BDUS, BDPW, Split(oCompany.Server, ":")(1), LOG) = False Then
                        Return ""
                    End If
                    'ObjGlobal.SBOApp.StatusBar.SetText("Llamando al modelo de impresión por defecto con los valores de los parámetros...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    For i = 0 To ParamValue.Length - 1
                        ' ObjGlobal.SBOApp.StatusBar.SetText("Parámetro CR:" & oCRReport.ParameterFields(i).Name, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'ObjGlobal.SBOApp.StatusBar.SetText(ParamName(i) & ":" & ParamValue(i), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If ParamValue(i).Contains("'") Then
                            oCRReport.SetParameterValue(ParamName(i), ParamValue(i).Replace("'", ""))
                        Else
                            oCRReport.SetParameterValue(ParamName(i), Integer.Parse(ParamValue(i)))
                        End If

                    Next
                    'oCRReport.Refresh()
                    If defaultPrinterSetting Is Nothing Then
                        'ObjGlobal.SBOApp.StatusBar.SetText("Salvando a disco el PDF:" & Name, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'oCRReport.SaveAs("D:\EO\CR.rpt")
                        oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, Name)
                        Return Name
                    Else
                        'ObjGlobal.SBOApp.StatusBar.SetText("Imprimiendo:" & Name, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oCRReport.PrintToPrinter(defaultPrinterSetting, defaultPrinterSetting.DefaultPageSettings, False)
                        Return "Impreso"
                    End If

                Catch ex As Exception
                    Throw New Exception(ex.Message)
                Finally
                    If Not oCRReport Is Nothing Then
                        oCRReport.Close()
                        oCRReport = Nothing
                    End If
                End Try
                Return Name

            Catch ex As Exception
                'ObjGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                Return ex.Message
            Finally
                If Not oCRReport Is Nothing Then
                    oCRReport.Close()
                End If
                oCRReport = Nothing
                ExportCR = Nothing
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(rs, Object))
            End Try
        End Function

        Private Shared Function ApplyNewServerNew(ocompany As SAPbobsCOM.Company, ByRef rd As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal Driver As String, ByVal Server As String, ByVal ServerNamespace As String, ByVal Username As String, ByVal Password As String, ByVal Port As String, LOG As EXO_Log.EXO_Log) As Boolean
            Dim ConnectionInfo As ConnectionInfo = Nothing
            Try
                ConnectionInfo = CreateConnection(ocompany, Driver, Server, Port, ServerNamespace, Username, Password, LOG)
                If ConnectionInfo Is Nothing Then
                    LOG.escribeMensaje("Imposible crear string de conexión!")
                    Return False
                End If
                If SetDBLogonForReport(ocompany, ConnectionInfo, rd, LOG) = False Then
                    Return False
                End If
                If SetDBLogonForSubreports(ocompany, ConnectionInfo, rd, LOG) = False Then
                    Return False
                End If
                Return True
            Catch ex As Exception
                LOG.escribeMensaje("ApplyNewServerNew " + ex.Message)
                Return False
            Finally
                If Not ConnectionInfo Is Nothing Then
                    ConnectionInfo = Nothing
                End If
            End Try
        End Function
        Private Shared Function CreateConnection(ByRef oCompany As SAPbobsCOM.Company, ByVal Driver As String, ByVal Server As String, ByVal Port As String, ByVal ServerNamespace As String, ByVal Username As String, ByVal Password As String, LOG As EXO_Log.EXO_Log) As ConnectionInfo
            Dim ConnectionInfo As ConnectionInfo = Nothing
            Try
                ConnectionInfo = New ConnectionInfo()
                If Server.Contains("@") Then
                    Server = Split(Server, "@")(1)
                End If
                If Port = "30013" Then
                    Port = "30015"
                End If
                Dim connString As String = String.Format("DRIVER={0};SERVERNODE={1};DATABASE={2}", Driver, Server & ":" & Port, ServerNamespace)
                ConnectionInfo.IntegratedSecurity = False
                ConnectionInfo.UserID = Username
                ConnectionInfo.AllowCustomConnection = True
                ConnectionInfo.Password = Password
                ConnectionInfo.ServerName = Server & ":" & Port
                ConnectionInfo.DatabaseName = ServerNamespace
                ConnectionInfo.Type = ConnectionInfoType.CRQE
                Return ConnectionInfo
            Catch ex As Exception
                LOG.escribeMensaje("CreateConnection: " + ex.Message)
                'ObjGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion, EXO_UIAPI.EXO_UIAPI.EXO_TipoSalidaMensaje.MessageBox)
                Return Nothing
            End Try
        End Function
        Private Shared Function SetDBLogonForReport(ByRef oCompany As SAPbobsCOM.Company, ByRef ConnectionInfo As ConnectionInfo, ByRef rd As CrystalDecisions.CrystalReports.Engine.ReportDocument, LOG As EXO_Log.EXO_Log) As Boolean
            Try
                For Each table As CrystalDecisions.CrystalReports.Engine.Table In rd.Database.Tables
                    Dim tableLogonInfo As TableLogOnInfo = table.LogOnInfo
                    tableLogonInfo.ConnectionInfo = ConnectionInfo
                    table.ApplyLogOnInfo(tableLogonInfo)
                Next
                Return True
            Catch ex As Exception
                LOG.escribeMensaje("SetDBLogonForReport: " + ex.Message)
                'ObjGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion, EXO_UIAPI.EXO_UIAPI.EXO_TipoSalidaMensaje.MessageBox)
                Return False
            End Try
        End Function
        Private Shared Function SetDBLogonForSubreports(ByRef oCompany As SAPbobsCOM.Company, ByRef connectionInfo As ConnectionInfo, ByRef rd As CrystalDecisions.CrystalReports.Engine.ReportDocument, LOG As EXO_Log.EXO_Log) As Boolean
            Try
                For Each section As Engine.Section In rd.ReportDefinition.Sections
                    For Each reportObject As ReportObject In section.ReportObjects
                        If reportObject.Kind = ReportObjectKind.SubreportObject Then
                            Dim SubreportObject As SubreportObject = CType(reportObject, SubreportObject)
                            Dim subReportDocument As ReportDocument = SubreportObject.OpenSubreport(SubreportObject.SubreportName)
                            If SetDBLogonForReport(oCompany, connectionInfo, subReportDocument, LOG) = False Then
                                Return False
                            End If
                        End If
                    Next
                Next
                Return True
            Catch ex As Exception
                LOG.escribeMensaje("SetDBLogonForSubreports: " + ex.Message)
                'ObjGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion, EXO_UIAPI.EXO_UIAPI.EXO_TipoSalidaMensaje.MessageBox)
                Return False
            End Try
        End Function
        Private Shared Function ApplySetConnection(ByRef oCompany As SAPbobsCOM.Company, ByRef rd As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal Driver As String, ByVal Server As String, ByVal ServerNamespace As String, ByVal Username As String, ByVal Password As String, ByVal Port As String, LOG As EXO_Log.EXO_Log) As Boolean

            Try
                For i = 0 To rd.DataSourceConnections.Count - 1
                    Dim LogonProperties As NameValuePairs2 = rd.DataSourceConnections.Item(i).LogonProperties
                    For j = 0 To rd.DataSourceConnections.Item(i).LogonProperties.Count - 1
                        Dim LogonPropertie As NameValuePair2 = CType(rd.DataSourceConnections.Item(i).LogonProperties.Item(j), NameValuePair2)
                        Select Case LogonPropertie.Name.ToString
                            Case "Connection String"
                                If Server.Contains("@") Then
                                    Server = Split(Server, "@")(1)
                                End If
                                If Port = "30013" Then
                                    Port = "30015"
                                End If
                                LogonPropertie.Value = String.Format("DRIVER={0};SERVERNODE={1};DATABASE={2}", Driver, Server & ":" & Port, ServerNamespace)
                            Case "Database"
                                LogonPropertie.Value = ServerNamespace
                            Case "Provider"
                                LogonPropertie.Value = Driver
                            Case "Server"
                                If Server.Contains("@") Then
                                    Server = Split(Server, "@")(1)
                                End If
                                If Port = "30013" Then
                                    Port = "30015"
                                End If
                                LogonPropertie.Value = Server & ":" & Port
                        End Select
                    Next
                    rd.DataSourceConnections.Item(i).SetLogonProperties(LogonProperties)
                    'For Each LogonPropertie As NameValuePair2 In LogonProperties

                    'Next

                Next
                'Connection String
                'Database
                'Provider
                'Server
                Return True
            Catch ex As Exception
                LOG.escribeMensaje("ApplySetConnection " + ex.Message)
                Return False
            Finally

            End Try
        End Function
        Private Shared Function Get_ReportLayout(ByRef oCompany As SAPbobsCOM.Company, ByVal LayoutCode As String) As ReportLayout

            Dim CmpSrv As SAPbobsCOM.CompanyService = Nothing
            Dim srv As ReportLayoutsService = Nothing
            Dim oParams As ReportLayoutParams = Nothing
            Dim oReportLayout As ReportLayout = Nothing

            Try

                CmpSrv = oCompany.GetCompanyService
                srv = CType(CmpSrv.GetBusinessService(ServiceTypes.ReportLayoutsService), ReportLayoutsService)
                oParams = CType(srv.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), ReportLayoutParams)
                oParams.LayoutCode = LayoutCode
                oReportLayout = srv.GetReportLayout(oParams)
                If Not oReportLayout Is Nothing Then
                    Return oReportLayout
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                'ObjGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                Return Nothing
            Finally
            End Try
        End Function
    End Class
End Namespace