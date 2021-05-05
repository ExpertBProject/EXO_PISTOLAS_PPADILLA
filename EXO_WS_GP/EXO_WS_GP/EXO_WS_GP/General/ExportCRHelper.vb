Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports SAPbobsCOM

Namespace SAP.ImpresionEtiquetas
    Public Class ExportCRHelper
        Private _company As SAPbobsCOM.Company = Nothing

        Private Const STR_EMPTY As String = ""
        Private Const STR_B1RptTokenPrefix As String = "{"
        Private Const STR_B1RptTokenSuffix As String = "}"

        Private Const STR_REPORT_TYPE As String = "RCRI"
        Private Const STR_NODENAME As String = "node"
        Private Const STR_RPTEXTENSION As String = ".rpt"

        Private Const STR_RDOC_TableName As String = "RDOC"
        Private Const STR_Blob_FieldName As String = "Template"
        Private Const STR_DocCode_KeySegmentName As String = "DocCode"
        Private Const STR_LangCode_KeySegmentName As String = "LangCode"

        Public Sub New(ByRef company As SAPbobsCOM.Company)
            _company = company
        End Sub

        ''' <summary>
        ''' Export content report from B1 Server to local
        ''' </summary>
        ''' <param name="reportCode"></param>
        Public Function ExportReport(ByVal reportCode As String) As String
            Dim rptFilePath As String = String.Empty
            Dim oCompanyService As CompanyService = Nothing
            Dim oLayoutService As ReportLayoutsService = Nothing
            Dim oParams As ReportLayoutParams = Nothing
            Dim oReportLayout As ReportLayout = Nothing
            Dim oBlobParams As BlobParams = Nothing
            Dim oKeySegment As BlobTableKeySegment = Nothing
            Dim oFile As FileStream = Nothing
            Dim oBlob As Blob = Nothing
            Dim buf As Byte() = Nothing

            'If the application is not valid, we will throw the exception.
            If Me._company Is Nothing Or Not Me._company.Connected Then
                Return String.Empty
            End If

            rptFilePath = Me.GetTempFolder() & reportCode & STR_B1RptTokenPrefix & Guid.NewGuid().ToString & STR_B1RptTokenSuffix & STR_RPTEXTENSION

            Try
                oCompanyService = _company.GetCompanyService()
                oLayoutService = CType(oCompanyService.GetBusinessService(ServiceTypes.ReportLayoutsService), ReportLayoutsService)

                oParams = CType(oLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), ReportLayoutParams)
                oParams.LayoutCode = reportCode

                'Get the ReportLayout object by reportCode.
                oReportLayout = oLayoutService.GetReportLayout(oParams)

                'Specify the table and field to update.
                oBlobParams = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), BlobParams)
                oBlobParams.Table = STR_RDOC_TableName
                oBlobParams.Field = STR_Blob_FieldName

                ' Specify the record whose blob field is to be set 
                oKeySegment = oBlobParams.BlobTableKeySegments.Add()
                oKeySegment.Name = STR_DocCode_KeySegmentName
                oKeySegment.Value = reportCode

                ' Get a blob field
                oBlob = CType(oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob), Blob)
                oBlob = oCompanyService.GetBlob(oBlobParams)

                'Get the content of the blob.
                buf = Convert.FromBase64String(oBlob.Content)

                'Put the buffer into the rpt file.
                oFile = New FileStream(rptFilePath, System.IO.FileMode.Create)

                oFile.Write(buf, 0, buf.Length)
                oFile.Close()

            Catch ex As System.Exception
            End Try

            Return rptFilePath
        End Function

        Private Function GetTempFolder() As String
            Dim tempDir As String = IO.Path.GetTempPath 'Environment.GetEnvironmentVariable("TEMP")

            If String.IsNullOrEmpty(tempDir) Then
                Return String.Empty
            End If

            'If Not tempDir.EndsWith("\\") Then
            '    tempDir = tempDir + "\\"
            'End If

            Return tempDir
        End Function
    End Class
End Namespace

