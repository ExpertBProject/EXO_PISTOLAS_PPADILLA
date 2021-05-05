Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oComp As SAPbobsCOM.Company
        oComp = New SAPbobsCOM.Company

        'oComp = New EXO_DIAPI.EXO_DIAPI()
        Try

            'Dim servidorSBO As String = System.Configuration.ConfigurationManager.AppSettings("servidorSBO")
            '    Dim servidorLicencias As String = System.Configuration.ConfigurationManager.AppSettings("servidorLicencias")
            '    Dim BDSBO As String = System.Configuration.ConfigurationManager.AppSettings("BDSBO")
            '    Dim usuarioSBO As String = System.Configuration.ConfigurationManager.AppSettings("usuarioSBO")
            '    Dim pwdSBO As String = System.Configuration.ConfigurationManager.AppSettings("pwdSBO")
            '    Dim usuarioHANA As String = System.Configuration.ConfigurationManager.AppSettings("usuarioHANA")
            '    Dim pwdHANA As String = System.Configuration.ConfigurationManager.AppSettings("pwdHANA")

            oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                oComp.UseTrusted = False
            oComp.CompanyDB = "PD_PPADILLA"
            'oComp.UserName = usuarioSBO
            'oComp.Password = pwdSBO
            oComp.UserName = "manager"
            oComp.Password = "Exo3x0$.1"
            oComp.Server = "TS1@xper-hanades02.hanab1.local:30013"
            'oComp.LicenseServer = servidorLicencias
            'oComp.DbUserName = "B1SQLUSER"
            'oComp.DbPassword = "12629iYk"


            If oComp.Connect() <> 0 Then
                    Dim algo As String = oComp.GetLastErrorDescription()



            Else

            End If




        Finally

        End Try

    End Sub
End Class
