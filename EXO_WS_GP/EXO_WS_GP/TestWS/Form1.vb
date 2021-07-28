Imports System.IO
Imports System.Runtime.Serialization

Imports System.Runtime.Serialization.Json


Public Class Form1

    'Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient

    End Sub

    Private Function ValidarCertificado(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        'System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.CompruebaArticulo_busqueda("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "M2Z", "")
        MessageBox.Show(respuestas)

    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.CompruebaUbicacion_busqueda("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "350,PL", "")
        MessageBox.Show(respuestas)

    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.ping()
        MessageBox.Show(respuestas)



    End Sub


    'Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
    '    System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.BasesDeDatos()
    '    MessageBox.Show(respuestas)

    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.LoginUsuario("PD_PPADILLA", "mperiz", "M@rt1nN1c01")
        MessageBox.Show(respuestas)

    End Sub





    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '    System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.UbicacionesDelAlmacen("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
    '    MessageBox.Show(respuestas)
    'End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        'respuestas = cliente.RecepcionMaterialesBuscador("PD_PPADILLA", "manager", "Exp3rt0n3$", "", "", "21000016", "")
        respuestas = cliente.RecepcionMaterialesBuscador("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "", "", "", "")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oCab As WS_GP.ListaPedidoCompraRegistrarLinea = New WS_GP.ListaPedidoCompraRegistrarLinea

        Dim oLineas As List(Of WS_GP.PedidoCompraRegistrarLinea) = New List(Of WS_GP.PedidoCompraRegistrarLinea)

        Dim oRegLinea As WS_GP.PedidoCompraRegistrarLinea = New WS_GP.PedidoCompraRegistrarLinea

        oRegLinea.NumInterno = "26"
        oRegLinea.NumLinea = 0
        oRegLinea.Proveedor = "P002681"
        oRegLinea.Codigo = "MHB0RCZZZZZZZ"

        oRegLinea.Lote = "21000013-001-001-350-P002681-210429"
        oRegLinea.CantidadReal = 100
        oRegLinea.CantidadSeleccionada = 100
        oRegLinea.Ubicacion = "350,PL"

        oRegLinea.QDES = 95
        oRegLinea.SCOF = "Y"
        oRegLinea.UOMO = "KGS"
        oRegLinea.UOMD = "UDS"
        oRegLinea.ORIG = "Y"
        oRegLinea.RATIO = 0.95

        oLineas.Add(oRegLinea)

        oCab.Lineas = oLineas.ToArray


        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oCab.GetType)
        js.WriteObject(str, oCab)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        respuestas = cliente.PedidoCompraRegistrarLinea2("PD_PPADILLA", "mperiz", "M@rt1nN1c0", JSON)
        MessageBox.Show(respuestas)
    End Sub



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.PedioCompraResumenFinalizar("PD_PPADILLA", "mperiz", "M@rt1nN1c0")

        'respuestas = cliente.PedidoCompraGenerar("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.ListasPicking("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.UbicacionesDelAlmacenBahias("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.DesglosePicking("PD_PPADILLA", "mperiz", "M@rt1nN1c0", 2)
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oTraslado As WS_GP.Traslado = New WS_GP.Traslado

        'oTraslado.CodigoArticulo = "500703530601"
        'oTraslado.Cantidad = 14
        'oTraslado.Lote = "C0002E17080002-0008"
        'oTraslado.Almacen = "02"
        'oTraslado.UbicacionOrigen = "02.15.PLAYA"
        'oTraslado.UbicacionDestino = "02.12.10.23.C"

        oTraslado.CodigoArticulo = "111001089900007"
        oTraslado.Cantidad = 500
        oTraslado.Lote = "19122017004"
        oTraslado.Almacen = "01LANDE"
        oTraslado.UbicacionOrigen = "01LANDEA010C"
        oTraslado.UbicacionDestino = "01LANDEA001B"
        oTraslado.NumeroPicking = 31
        oTraslado.PickingLinea = 0

        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oTraslado.GetType)
        js.WriteObject(str, oTraslado)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        respuestas = cliente.OperacionesTraslado(JSON, "ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oPicking As WS_GP.GenerarPicking = New WS_GP.GenerarPicking

        Dim oLinea As WS_GP.LineasPicking = New WS_GP.LineasPicking
        Dim oLineas As List(Of WS_GP.LineasPicking) = New List(Of WS_GP.LineasPicking)

        Dim oBulto As WS_GP.BultosPicking = New WS_GP.BultosPicking
        Dim oBultos As List(Of WS_GP.BultosPicking) = New List(Of WS_GP.BultosPicking)

        Dim oPalet As WS_GP.PaletsPicking = New WS_GP.PaletsPicking
        Dim oPalets As List(Of WS_GP.PaletsPicking) = New List(Of WS_GP.PaletsPicking)



        oPicking.NumeroPicking = 3
        oPicking.Ubicacion = "02BAHIA"
        oPicking.Resultado = ""

        oLinea = New WS_GP.LineasPicking
        oLinea.Articulo = "000003"
        oLinea.Cantidad = 1
        oLinea.Lote = ""
        oLinea.PickingLinea = 1
        oLineas.Add(oLinea)

        oLinea = New WS_GP.LineasPicking
        oLinea.Articulo = "000002"
        oLinea.Cantidad = 1
        oLinea.Lote = "L02"
        oLinea.PickingLinea = 0
        oLineas.Add(oLinea)

        oPicking.Lineas = oLineas.ToArray

        oBulto = New WS_GP.BultosPicking
        oBulto.Articulo = "110001110100001"
        oBulto.Cantidad = 2000
        oBulto.Lote = ""
        oBulto.Bulto = 1
        oBulto.LineaPicking = 1
        oBultos.Add(oBulto)


        oPicking.Bultos = oBultos.ToArray

        oPalet = New WS_GP.PaletsPicking
        oPalet.Tipo = "europalet"
        oPalet.Palet = 1
        oPalet.Peso = 1
        oPalet.Volumen = 0.96
        oPalet.Altura = 1
        oPalets.Add(oPalet)

        oPicking.Palets = oPalets.ToArray

        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oPicking.GetType)
        js.WriteObject(str, oPicking)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        'Dim path As String = "E:\Desarrollo\Usuarios\mperiz\picking.txt"
        'Dim readText As String = File.ReadAllText(path)
        'JSON = readText

        respuestas = cliente.GenerarPicking2(JSON, "DEMO_SBO", "mperiz", "M@rt1nN1c0")


        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.ListasPickingTraslado("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.DesglosePickingTraslado("PD_PPADILLA", "mperiz", "M@rt1nN1c0", 13)
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oTraslado As WS_GP.Traslado = New WS_GP.Traslado

        'oTraslado.CodigoArticulo = "500703530601"
        'oTraslado.Cantidad = 14
        'oTraslado.Lote = "C0002E17080002-0008"
        'oTraslado.Almacen = "02"
        'oTraslado.UbicacionOrigen = "02.15.PLAYA"
        'oTraslado.UbicacionDestino = "02.12.10.23.C"

        oTraslado.CodigoArticulo = "M2Z010E070M68"
        oTraslado.Cantidad = 95
        oTraslado.Lote = "L3"
        oTraslado.Almacen = "350"
        oTraslado.UbicacionOrigen = "350,PL"
        oTraslado.UbicacionDestino = "350,03,00,00,12"
        oTraslado.NumeroPicking = 11
        oTraslado.PickingLinea = 0
        oTraslado.Motivo = "67"

        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oTraslado.GetType)
        js.WriteObject(str, oTraslado)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        respuestas = cliente.OperacionesTraslado(JSON, "PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oPicking As WS_GP.GenerarTraslado = New WS_GP.GenerarTraslado

        Dim oLinea As WS_GP.LineasPicking = New WS_GP.LineasPicking
        Dim oLineas As List(Of WS_GP.LineasPicking) = New List(Of WS_GP.LineasPicking)

        Dim oBulto As WS_GP.BultosPicking = New WS_GP.BultosPicking
        Dim oBultos As List(Of WS_GP.BultosPicking) = New List(Of WS_GP.BultosPicking)

        Dim oPalet As WS_GP.PaletsPicking = New WS_GP.PaletsPicking
        Dim oPalets As List(Of WS_GP.PaletsPicking) = New List(Of WS_GP.PaletsPicking)



        oPicking.NumeroTraslado = 13
        oPicking.Ubicacion = "350,03,00,00,12"
        oPicking.Resultado = ""

        oLinea = New WS_GP.LineasPicking
        oLinea.Articulo = "M2Z010E070M68"
        oLinea.Cantidad = 95
        oLinea.Lote = "L3"
        oLinea.PickingLinea = 0
        oLineas.Add(oLinea)



        oPicking.Lineas = oLineas.ToArray

        oBulto = New WS_GP.BultosPicking
        oBulto.Articulo = "M2Z010E070M68"
        oBulto.Cantidad = 95
        oBulto.Lote = "L3"
        oBulto.Bulto = 1
        oBulto.LineaPicking = 0
        oBultos.Add(oBulto)


        oPicking.Bultos = oBultos.ToArray

        oPalet = New WS_GP.PaletsPicking
        oPalet.Tipo = "europalet"
        oPalet.Palet = 1
        oPalet.Peso = 20
        oPalet.Volumen = 0.96
        oPalet.Altura = 1
        oPalets.Add(oPalet)

        oPicking.Palets = oPalets.ToArray

        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oPicking.GetType)
        js.WriteObject(str, oPicking)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        'Dim path As String = "E:\Desarrollo\Usuarios\mperiz\picking.txt"
        'Dim readText As String = File.ReadAllText(path)
        'JSON = readText

        respuestas = cliente.GenerarPickingTraslado(JSON, "PD_PPADILLA", "mperiz", "M@rt1nN1c0")


        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.LeerQR("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "P002681;21000006;MHB0RCZZZZZZZ;350,PL;100.00;105.00;BOBINA 205x2,5 DECAPADO ;22/03/2021;21000012-001-001-350-P002681-210429")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.ListadoImprimir("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "64", "20")
        MessageBox.Show(respuestas)

    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oPicking As WS_GP.ListaLotesImprimir = New WS_GP.ListaLotesImprimir

        Dim oLinea As WS_GP.LotesImprimir = New WS_GP.LotesImprimir
        Dim oLineas As List(Of WS_GP.LotesImprimir) = New List(Of WS_GP.LotesImprimir)

        oPicking.Resultado = ""

        oLinea = New WS_GP.LotesImprimir
        oLinea.TipoEtiqueta = 1
        oLinea.Lote = "21000019-001-020-01-A000018-210602"
        oLinea.CodArticulo = "MHB0RCZZZZZZZ"
        oLinea.SysNumber = 613
        oLinea.Ubicacion = "350,04,00,00,14"
        oLinea.Impresora = "Microsoft Print to PDF"
        oLineas.Add(oLinea)

        oPicking.LotesImprimir = oLineas.ToArray

        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oPicking.GetType)
        js.WriteObject(str, oPicking)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        'Dim path As String = "E:\Desarrollo\Usuarios\mperiz\picking.txt"

        'If File.Exists(path) Then
        '    Dim algo As String = ""
        'End If

        'Dim readText As String = File.ReadAllText(path)
        'JSON = readText

        respuestas = cliente.LanzoImprimir(JSON, "PD_PPADILLA", "mperiz", "M@rt1nN1c0")


        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.RecepcionTrasladoListado("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)

    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.RecepcionTrasladosBuscador("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "20")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        Dim oCab As WS_GP.ListaPedidoCompraRegistrarLinea = New WS_GP.ListaPedidoCompraRegistrarLinea

        Dim oLineas As List(Of WS_GP.PedidoCompraRegistrarLinea) = New List(Of WS_GP.PedidoCompraRegistrarLinea)

        Dim oRegLinea As WS_GP.PedidoCompraRegistrarLinea = New WS_GP.PedidoCompraRegistrarLinea

        oRegLinea.NumInterno = "20"
        oRegLinea.NumLinea = 0
        oRegLinea.Proveedor = "C015449"
        oRegLinea.Codigo = "M2Z010E070M68"

        oRegLinea.Lote = "L3"
        oRegLinea.CantidadReal = 95
        oRegLinea.CantidadSeleccionada = 95
        oRegLinea.Ubicacion = "555-UBICACIÓN-DE-SISTEMA"

        oRegLinea.QDES = 95
        oRegLinea.SCOF = ""
        oRegLinea.UOMO = ""
        oRegLinea.UOMD = ""
        oRegLinea.ORIG = "Y"
        oRegLinea.RATIO = 1

        oLineas.Add(oRegLinea)

        oCab.Lineas = oLineas.ToArray


        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oCab.GetType)
        js.WriteObject(str, oCab)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        respuestas = cliente.RecepcionTrasladoRegistrarLinea2("PD_PPADILLA", "mperiz", "M@rt1nN1c0", JSON)
        MessageBox.Show(respuestas)

    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.RecepcionTrasladoResumenFinalizar("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click

        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        respuestas = cliente.RecepcionTrasladoGenerar("PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub



    'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    '    System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.ComprobarExisteArticulo("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "018435043130414")
    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
    '    System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.ComprobarArticuloSalida("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "111001089900001", "", 500, "01_LANDE-A-A-1")
    '    MessageBox.Show(respuestas)
    'End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""
        Dim oDoc As WS_GP.OperacionEntradaSalida = New WS_GP.OperacionEntradaSalida

        Dim oLinea As WS_GP.Articulo = New WS_GP.Articulo
        Dim oLineas As List(Of WS_GP.Articulo) = New List(Of WS_GP.Articulo)

        oLinea = New WS_GP.Articulo
        oLinea.ArticuloMember = "M2Z010E070M68"
        oLinea.Cantidad = 100
        oLinea.Lote = "19122017004"
        oLinea.Ubicacion = "350,PL"
        oLinea.PrecioProducto = 20

        oDoc.Motivo = "60"
        oDoc.CC = "11300002"



        oLineas.Add(oLinea)


        oDoc.Lineas = oLineas.ToArray

        Dim str As New MemoryStream()
        Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oDoc.GetType)
        js.WriteObject(str, oDoc)
        str.Position = 0
        Dim sr As New StreamReader(str)
        Dim JSON As String = sr.ReadToEnd()

        respuestas = cliente.GenerarDocumentoEntradaManual(JSON, "PD_PPADILLA", "mperiz", "M@rt1nN1c0")
        MessageBox.Show(respuestas)
    End Sub

    'Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""
    '    Dim oDoc As WS_GP.OperacionEntradaSalida = New WS_GP.OperacionEntradaSalida

    '    Dim oLinea As WS_GP.Articulo = New WS_GP.Articulo
    '    Dim oLineas As List(Of WS_GP.Articulo) = New List(Of WS_GP.Articulo)

    '    oLinea = New WS_GP.Articulo
    '    oLinea.ArticuloMember = "111001089900001"
    '    oLinea.Cantidad = 10
    '    oLinea.Lote = ""
    '    oLinea.Ubicacion = "01_LANDE-A-A-2"
    '    oLineas.Add(oLinea)

    '    oLinea = New WS_GP.Articulo
    '    oLinea.ArticuloMember = "111001089900007"
    '    oLinea.Cantidad = 1
    '    oLinea.Lote = "11122017_1"
    '    oLinea.Ubicacion = "01_LANDE-A-A-2"
    '    oLineas.Add(oLinea)

    '    oLinea = New WS_GP.Articulo
    '    oLinea.ArticuloMember = "111001089900007"
    '    oLinea.Cantidad = 1
    '    oLinea.Lote = "11122017_2"
    '    oLinea.Ubicacion = "01_LANDE-A-A-2"
    '    oLineas.Add(oLinea)

    '    oDoc.Lineas = oLineas.ToArray

    '    Dim str As New MemoryStream()
    '    Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oDoc.GetType)
    '    js.WriteObject(str, oDoc)
    '    str.Position = 0
    '    Dim sr As New StreamReader(str)
    '    Dim JSON As String = sr.ReadToEnd()

    '    respuestas = cliente.GenerarDocumentoSalidaManual(JSON, "ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.DesgloseSolicitudesTraslado("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "1")
    '    MessageBox.Show(respuestas)

    'End Sub

    'Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.ListasSolicitudTraslado("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""
    '    Dim oDoc As WS_GP.OperacionTraslado = New WS_GP.OperacionTraslado

    '    Dim oLinea As WS_GP.LineasTraslado = New WS_GP.LineasTraslado
    '    Dim oLineas As List(Of WS_GP.LineasTraslado) = New List(Of WS_GP.LineasTraslado)

    '    oDoc.NumeroSolTraslado = 3


    '    oLinea = New WS_GP.LineasTraslado
    '    oLinea.Articulo = "104001020100001"
    '    oLinea.Cantidad = 35
    '    oLinea.Lote = ""
    '    oLinea.UbicacionOrigen = "01_LANDE-A-A-2"
    '    oLinea.UbicacionDestino = "01_LANDE-A-A-3"
    '    oLinea.NumeroLinea = 0
    '    oLineas.Add(oLinea)

    '    oLinea = New WS_GP.LineasTraslado
    '    oLinea.Articulo = "111001089900001"
    '    oLinea.Cantidad = 390
    '    oLinea.Lote = ""
    '    oLinea.UbicacionOrigen = "01_LANDE-A-A-5"
    '    oLinea.UbicacionDestino = "01_LANDE-A-A-3"
    '    oLinea.NumeroLinea = 1
    '    oLineas.Add(oLinea)

    '    oDoc.Lineas = oLineas.ToArray

    '    Dim str As New MemoryStream()
    '    Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oDoc.GetType)
    '    js.WriteObject(str, oDoc)
    '    str.Position = 0
    '    Dim sr As New StreamReader(str)
    '    Dim JSON As String = sr.ReadToEnd()

    '    respuestas = cliente.GenerarOperacionTraslado(JSON, "ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
    '    MessageBox.Show(respuestas)
    'End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click



        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        'respuestas = cliente.ComprobarExisteArticulo("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "2000004665241")
        respuestas = cliente.ComPruebaArticulo("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "MHB0RCZZZZZZZ", "2000004665241", "Y")
        MessageBox.Show(respuestas)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click

        Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
        Dim respuestas As String = ""

        'respuestas = cliente.ComprobarExisteArticulo("PD_PPADILLA", "mperiz", "M@rt1nN1c0", "2000004665241")
        respuestas = cliente.BasesDeDatos()
        MessageBox.Show(respuestas)

    End Sub





    'Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click

    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.GenerarDraftEntrega("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
    '    'respuestas = cliente.ComPruebaArticulo("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "110001180200001", "]C1011843504310096410123", "Y")
    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click

    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    respuestas = cliente.PedidoCompraGenerar("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0")
    'End Sub

    'Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    'respuestas = cliente.ConsultaStock("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "01LANDEC002A")
    '    respuestas = cliente.ConsultaStock("DEMO_SBO", "mperiz", "M@rt1nN1c0", "0125896314785241")

    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    'respuestas = cliente.ConsultaStock("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "01LANDEC002A")
    '    respuestas = cliente.ListasRecuentoInventario("DEMO_SBO", "mperiz", "M@rt1nN1c0")

    '    MessageBox.Show(respuestas)

    '    '
    'End Sub

    'Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    'respuestas = cliente.ConsultaStock("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "01LANDEC002A")
    '    respuestas = cliente.DesgloseRecuentoInventario("DEMO_SBO", "mperiz", "M@rt1nN1c0", "21")

    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click

    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""
    '    Dim oDoc As WS_GP.ListasRecuentoInventarioCabecera = New WS_GP.ListasRecuentoInventarioCabecera

    '    Dim oLinea As WS_GP.ListasRecuentoInventarioDetalle = New WS_GP.ListasRecuentoInventarioDetalle
    '    Dim oLineas As List(Of WS_GP.ListasRecuentoInventarioDetalle) = New List(Of WS_GP.ListasRecuentoInventarioDetalle)

    '    oLinea = New WS_GP.ListasRecuentoInventarioDetalle
    '    oLinea.CantidadContada = "80"
    '    oLinea.Articulo = "000003"
    '    oLinea.CodUbicacion = "1"
    '    oLineas.Add(oLinea)

    '    'oLinea = New WS_GP.ListasRecuentoInventarioDetalle
    '    'oLinea.ArticuloMember = "111001089900007"
    '    'oLinea.Cantidad = 1
    '    'oLinea.Lote = "11122017_1"
    '    'oLinea.Ubicacion = "01_LANDE-A-A-2"
    '    'oLineas.Add(oLinea)

    '    oDoc.NumeroInterno = "5"

    '    oDoc.Lineas = oLineas.ToArray

    '    Dim str As New MemoryStream()
    '    Dim js As New System.Runtime.Serialization.Json.DataContractJsonSerializer(oDoc.GetType)
    '    js.WriteObject(str, oDoc)
    '    str.Position = 0
    '    Dim sr As New StreamReader(str)
    '    Dim JSON As String = sr.ReadToEnd()

    '    respuestas = cliente.GenerarRecuentoInventario(JSON, "DEMO_SBO", "mperiz", "M@rt1nN1c0")
    '    MessageBox.Show(respuestas)
    'End Sub

    'Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
    '    Dim cliente As WS_GP.EXO_WS_GPClient = New WS_GP.EXO_WS_GPClient
    '    Dim respuestas As String = ""

    '    'respuestas = cliente.ConsultaStock("ALEX_BANKINTER", "mperiz", "M@rt1nN1c0", "01LANDEC002A")
    '    respuestas = cliente.RecuentoInventarioMarcarFinalizado("DEMO_SBO", "mperiz", "M@rt1nN1c0", "5")

    '    MessageBox.Show(respuestas)
    'End Sub
End Class
