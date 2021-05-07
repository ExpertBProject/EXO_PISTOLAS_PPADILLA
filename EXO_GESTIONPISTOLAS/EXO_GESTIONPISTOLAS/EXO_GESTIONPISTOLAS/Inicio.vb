
Public Class Inicio
    Inherits EXO_UIAPI.EXO_DLLBase

    Public Sub New(general As EXO_UIAPI.EXO_UIAPI, actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(general, actualizar, False, idAddOn)



        If actualizar Then
            cargaCampos()
        End If

    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        'Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        'filtrosXML.LoadXml(objGlobal.Functions.leerEmbebido(Me.GetType(), "FiltrosMDFR.xml"))
        'Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        'filtro.LoadFromXML(filtrosXML.OuterXml)
        'Return filtro
        Return Nothing
    End Function

    Public Overrides Function menus() As Xml.XmlDocument
        'Dim menuXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "MenuMDFRportes.xml")
        'Dim menu As Xml.XmlDocument = New Xml.XmlDocument
        'menu.LoadXml(menuXML)
        ' Return menu
        Return Nothing
    End Function

    Public Sub cargaCampos()

        If objGlobal.refDi.comunes.esAdministrador() Then
            objGlobal.SBOApp.StatusBar.SetText("El usuario es administrador", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim contenidoXML As String

            Try

                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UTs_EXO_GP_TRACOM.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UTs_EXO_GP_PEDCOM.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UTs_EXO_GP_PEDCOM.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UTs_EXO_GP_PEDCOM.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OWTR.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UDFs_OWTR.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)



                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_WTR1.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UDFs_WTR1.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_GP.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UDFs_GP.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OPKL.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UDFs_OPKL.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OINC.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)
                objGlobal.SBOApp.StatusBar.SetText("Validado UDFs_OINC.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


                Dim oXML As String = ""
                Dim udoObj As EXO_Generales.EXO_UDO = Nothing

                Dim sError As String = ""
                objGlobal.SBOApp.StatusBar.SetText("Comprobando UDO_EXO_OGPPA.xml", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_OGPPA.xml")
                If objGlobal.refDi.comunes.LoadBDFromXML(oXML, sError) = False Then
                    objGlobal.SBOApp.MessageBox("Error al actualizar UDO_EXO_OGPPA.xml:" & sError)
                    'Return False
                End If



                'Dim sql As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "exo_gp_propongo_lote.sql")
                'objGlobal.SQL.executeNonQuery(sql)

                'sql = objGlobal.Functions.leerEmbebido(Me.GetType(), "exo_gp_propongo_ubicacion.sql")
                'objGlobal.SQL.executeNonQuery(sql)

                'sql = objGlobal.Functions.leerEmbebido(Me.GetType(), "exo_gp_gestionar_en_pick.sql")
                'objGlobal.SQL.executeNonQuery(sql)

                'sql = objGlobal.Functions.leerEmbebido(Me.GetType(), "exo_gp_trabajo_lista_picking.sql")
                'objGlobal.SQL.executeNonQuery(sql)

            Catch ex As Exception

            End Try
        Else
            objGlobal.SBOApp.StatusBar.SetText("El usuario NO es administrador")
        End If
    End Sub


    'Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
    '    Dim res As Boolean = True
    '    Dim tipoForm As String = ""
    '    Try
    '        tipoForm = infoEvento.FormTypeEx
    '        Select Case tipoForm
    '            Case "40014"
    '                res = eventosEXO_40014.SBOApp_ItemEvent(infoEvento)
    '            Case "149"
    '                res = eventosEXO_149.SBOApp_ItemEvent(infoEvento)
    '            Case "139"
    '                res = eventosEXO_139.SBOApp_ItemEvent(infoEvento)
    '        End Select
    '        Return res
    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
    '        Return False
    '    Catch ex As Exception
    '        objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
    '        Return False
    '    End Try
    'End Function
    'Public Overrides Function SBOApp_FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    Dim sItemCode As String = ""
    '    Dim res As Boolean = True

    '    Try
    '        If infoEvento.BeforeAction = True Then
    '            Select Case infoEvento.FormTypeEx
    '                Case "40014"
    '                    Select Case infoEvento.EventType

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

    '                    End Select
    '                Case "149"
    '                    Select Case infoEvento.EventType

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

    '                    End Select
    '                Case "139"
    '                    Select Case infoEvento.EventType

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

    '                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

    '                    End Select
    '            End Select
    '        End If

    '        Return res

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
    '        Return False
    '    Catch ex As Exception
    '        objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
    '        Return False
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
    '    End Try
    'End Function


End Class

