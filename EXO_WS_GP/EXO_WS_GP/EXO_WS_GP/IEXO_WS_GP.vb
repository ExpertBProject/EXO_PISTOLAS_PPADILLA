' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IService1" en el código y en el archivo de configuración a la vez.
Imports EXO_WS_GP

<ServiceContract()>
Public Interface IEXO_WS_GP

#Region "Definiciones Interface"

    <OperationContract()>
    Function ping() As Boolean

    <OperationContract()>
    Function LoginUsuario(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function BasesDeDatos() As String

    <OperationContract()>
    Function UbicacionesDelAlmacen(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function RecepcionMaterialesBuscador(BaseDatos As String, Usuario As String, Password As String, NomProv As String, NumContenedor As String, CodEan As String, DescArt As String) As String

    <OperationContract()>
    Function PedidoCompraRegistrarLinea(BaseDatos As String, Usuario As String, Password As String, JSON As String) As String

    <OperationContract()>
    Function PedidoCompraGenerar(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function ListasPicking(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function UbicacionesDelAlmacenBahias(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function DesglosePicking(BaseDatos As String, Usuario As String, Password As String, NumeroPicking As String) As String

    <OperationContract()>
    Function OperacionesTraslado(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function GenerarPicking(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function ComprobarExisteArticulo(BaseDatos As String, Usuario As String, Password As String, CodEan As String) As String

    <OperationContract()>
    Function ComprobarArticuloSalida(BaseDatos As String, Usuario As String, Password As String, Articulo As String, Lote As String, Cantidad As Double, Ubicacion As String) As String

    <OperationContract()>
    Function GenerarDocumentoEntradaManual(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function GenerarDocumentoSalidaManual(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function ComPruebaArticulo(BaseDatos As String, Usuario As String, Password As String, CodArticulo As String, CodEan As String, EsLote As String) As String

    ' TODO: agregue aquí sus operaciones de servicio
    <OperationContract()>
    Function zIniClassPedidoCompraRegistrarLinea() As PedidoCompraRegistrarLinea

    <OperationContract()>
    Function zIniClassTraslados() As Traslado

    <OperationContract()>
    Function zIniGenerarPicking() As GenerarPicking

    <OperationContract()>
    Function zOperacionEntradaSalida() As OperacionEntradaSalida

    <OperationContract>
    Function PedioCompraResumenFinalizar(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function zOperacionTraslado() As OperacionTraslado

    <OperationContract()>
    Function GenerarPicking2(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function PedidoCompraRegistrarLinea2(BaseDatos As String, Usuario As String, Password As String, JSON As String) As String

    <OperationContract()>
    Function GenerarDraftEntrega(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function ConsultaStock(BaseDatos As String, Usuario As String, Password As String, Filtro As String) As String

    <OperationContract()>
    Function ListasPickingMultiple(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function DesglosePickingMultiple(BaseDatos As String, Usuario As String, Password As String, NumeroPicking As String) As String

    <OperationContract()>
    Function ListasRecuentoInventario(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function DesgloseRecuentoInventario(BaseDatos As String, Usuario As String, Password As String, NumRecuento As String) As String

    <OperationContract()>
    Function zListasRecuentoInventarioCabecera() As ListasRecuentoInventarioCabecera

    <OperationContract()>
    Function zListasRecuentoInventarioDetalle() As ListasRecuentoInventarioDetalle

    <OperationContract()>
    Function zListaPedidoCompraRegistrarLinea() As ListaPedidoCompraRegistrarLinea

    <OperationContract()>
    Function GenerarRecuentoInventario(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function RecuentoInventarioMarcarFinalizado(BaseDatos As String, Usuario As String, Password As String, NumRecuento As String) As String

    <OperationContract()>
    Function CompruebaUbicacion_busqueda(BaseDatos As String, Usuario As String, Password As String, Filtro As String, ItemCode As String) As String

    <OperationContract()>
    Function CompruebaArticulo_busqueda(BaseDatos As String, Usuario As String, Password As String, Filtro As String, Ubicacion As String) As String

    <OperationContract()>
    Function ListasPickingTraslado(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function DesglosePickingTraslado(BaseDatos As String, Usuario As String, Password As String, NumeroPicking As String) As String

    <OperationContract()>
    Function GenerarPickingTraslado(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function zListasGenerarTraslado() As GenerarTraslado

    <OperationContract()>
    Function LeerQR(BaseDatos As String, Usuario As String, Password As String, Lectura As String) As String

    <OperationContract()>
    Function ListadoImprimir(BaseDatos As String, Usuario As String, Password As String, DocEntry As String, ObjType As String) As String

    <OperationContract()>
    Function LanzoImprimir(JSON As String, BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function zListaLotesImprimir() As ListaLotesImprimir

    <OperationContract()>
    Function zLotesImprimir() As LotesImprimir

    <OperationContract()>
    Function RecepcionTrasladoListado(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function RecepcionTrasladosBuscador(BaseDatos As String, Usuario As String, Password As String, NumTraslado As String) As String

    <OperationContract()>
    Function RecepcionTrasladoRegistrarLinea2(BaseDatos As String, Usuario As String, Password As String, JSON As String) As String

    <OperationContract()>
    Function RecepcionTrasladoResumenFinalizar(BaseDatos As String, Usuario As String, Password As String) As String

    <OperationContract()>
    Function RecepcionTrasladoGenerar(BaseDatos As String, Usuario As String, Password As String) As String

#End Region

End Interface

<DataContract()>
<Serializable()>
Public Class Resultado
    <DataMember()>
    Public Resultado As String
End Class

<DataContract()>
<Serializable()>
Public Class BasesDatos
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public BD As String
    <DataMember()>
    Public Almacen As String
End Class

<DataContract()>
<Serializable()>
Public Class CompruebaArticulo
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Lote As String

End Class

<DataContract()>
<Serializable()>
Public Class CompruebaArticuloBus
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Codigo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public CodEan As String
End Class

<DataContract()>
<Serializable()>
Public Class Ubicaciones
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Codigo As String
    <DataMember()>
    Public PesoMaximo As String
    <DataMember()>
    Public PesoUbicacion As String
End Class

<DataContract()>
<Serializable()>
Public Class PedidosCompra
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumInterno As String
    <DataMember()>
    Public NumDocumento As String
    <DataMember()>
    Public NumLinea As String
    <DataMember()>
    Public Proveedor As String
    <DataMember()>
    Public ProveedorNombre As String
    <DataMember()>
    Public Codigo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public Cantidad As String
    <DataMember()>
    Public UnidadMedida As String
    <DataMember()>
    Public UdmLinea As String
    <DataMember()>
    Public EsLote As String
    <DataMember()>
    Public Largo As Double
    <DataMember()>
    Public Peso As Double
    <DataMember()>
    Public Alto As Double
    <DataMember()>
    Public Ancho As Double
    <DataMember()>
    Public UdmArticulo As List(Of UdmArticulo)
    <DataMember()>
    Public SCOF As String
    <DataMember()>
    Public EAN As String
End Class

<DataContract()>
<Serializable()>
Public Class UdmArticulo
    <DataMember()>
    Public Udm As String
End Class

<DataContract()>
<Serializable()>
Public Class ListaPedidoCompraRegistrarLinea
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Lineas As List(Of PedidoCompraRegistrarLinea)
End Class

<DataContract()>
<Serializable()>
Public Class PedidoCompraRegistrarLinea
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumInterno As String
    <DataMember()>
    Public NumLinea As String
    <DataMember()>
    Public Proveedor As String
    <DataMember()>
    Public Codigo As String
    <DataMember()>
    Public CantidadReal As String
    <DataMember()>
    Public CantidadSeleccionada As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public Largo As Double
    <DataMember()>
    Public Peso As Double
    <DataMember()>
    Public Alto As Double
    <DataMember()>
    Public Ancho As Double
    <DataMember()>
    Public QDES As Double
    <DataMember()>
    Public SCOF As String
    <DataMember()>
    Public UOMO As String
    <DataMember()>
    Public UOMD As String
    <DataMember()>
    Public ORIG As String
    <DataMember()>
    Public RATIO As Double
End Class

<DataContract()>
<Serializable()>
Public Class PedidoCompraResumenFinalizar

    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Proveedor As String
    <DataMember()>
    Public Codigo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public UnidadMedida As String

End Class

<DataContract()>
<Serializable()>
Public Class ListasPicking
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Numero As String
    <DataMember()>
    Public Fecha As String
    <DataMember()>
    Public Comentario As String
    <DataMember()>
    Public Transportista As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public NumeroInternoTraslado As String

End Class

<DataContract()>
<Serializable()>
Public Class Picking
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public PickingLinea As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public CantidadTotal As Double
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public UnidadMedida As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public UbicacionPropuesta As String
    <DataMember()>
    Public EsLote As String
    <DataMember()>
    Public Procesado As String
    <DataMember()>
    Public CantidadUDM As Double
    <DataMember()>
    Public CantidadPicking As String
End Class

<DataContract()>
<Serializable()>
Public Class Traslado
    <DataMember()>
    Public Almacen As String
    <DataMember()>
    Public CodigoArticulo As String
    <DataMember()>
    Public UbicacionOrigen As String
    <DataMember()>
    Public Cantidad As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public UbicacionDestino As String
    <DataMember()>
    Public NumeroPicking As String
    <DataMember()>
    Public PickingLinea As String
End Class

<DataContract()>
<Serializable()>
Public Class ArticuloPicking
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String

End Class

<DataContract()>
<Serializable()>
Public Class GenerarPicking
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumeroPicking As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public Lineas As List(Of LineasPicking)
    <DataMember()>
    Public Bultos As List(Of BultosPicking)
    <DataMember()>
    Public Palets As List(Of PaletsPicking)

End Class

<DataContract()>
<Serializable()>
Public Class LineasPicking
    <DataMember()>
    Public PickingLinea As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Lote As String

End Class

<DataContract()>
<Serializable()>
Public Class BultosPicking
    <DataMember()>
    Public Bulto As Integer
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public LineaPicking As String

End Class

<DataContract()>
<Serializable()>
Public Class PaletsPicking
    <DataMember()>
    Public Palet As Integer
    <DataMember()>
    Public Tipo As String
    <DataMember()>
    Public Peso As Double
    <DataMember()>
    Public Volumen As Double
    <DataMember()>
    Public Altura As Double
End Class

<DataContract()>
<Serializable()>
Public Class OperacionEntradaSalida
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Lineas As List(Of Articulo)
End Class

<DataContract()>
<Serializable()>
Public Class Articulo
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public UnidadMedida As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public EsLote As String
    <DataMember()>
    Public Largo As Double
    <DataMember()>
    Public Peso As Double
    <DataMember()>
    Public Alto As Double
    <DataMember()>
    Public Ancho As Double
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public CantidadUDM As Double

End Class

<DataContract()>
<Serializable()>
Public Class OperacionTraslado
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumeroSolTraslado As String
    <DataMember()>
    Public NumeroSolTrasladoInterno As String
    <DataMember()>
    Public Lineas As List(Of LineasTraslado)
End Class

<DataContract()>
<Serializable()>
Public Class LineasTraslado
    <DataMember()>
    Public NumeroLinea As String
    <DataMember()>
    Public UbicacionOrigen As String
    <DataMember()>
    Public UbicacionDestino As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public UnidadMedida As String
End Class


<DataContract()>
<Serializable()>
Public Class ListasPickingMultiple
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Numero As String
    <DataMember()>
    Public Fecha As String
    <DataMember()>
    Public Comentario As String
    <DataMember()>
    Public Transportista As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public NumeroInternoTraslado As String
    <DataMember()>
    Public CentroCoste As String
    <DataMember()>
    Public NumLineas As Integer
    <DataMember()>
    Public Lineas As List(Of ListasPickingDetalle)
End Class

<DataContract()>
<Serializable()>
Public Class ListasPickingDetalle
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public Cantidad As String
    <DataMember()>
    Public Proveedor As String
    <DataMember()>
    Public Stock As Double
End Class


<DataContract()>
<Serializable()>
Public Class PickingMultiple
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public PickingLinea As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public CantidadTotal As Double
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public UnidadMedida As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public UbicacionPropuesta As String
    <DataMember()>
    Public EsLote As String
    <DataMember()>
    Public Procesado As String
    <DataMember()>
    Public CantidadUDM As Double
    <DataMember()>
    Public UbicacionBahia As String
    <DataMember()>
    Public Orden As String
    <DataMember()>
    Public CantidadPicking As String
End Class

<DataContract()>
<Serializable()>
Public Class Stock
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Codigo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember>
    Public Lote As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public Estatus As String
    <DataMember()>
    Public InfoDetallada As String
    <DataMember()>
    Public UnidadMedida As String

End Class


''SIN MONTAR AUN EN ESTE WS

<DataContract()>
<Serializable()>
Public Class SolicitudTraslado
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public PickingLinea As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public CantidadTotal As Double
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public UnidadMedida As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public UbicacionPropuesta As String
    <DataMember()>
    Public EsLote As String
    <DataMember()>
    Public Procesado As String
    <DataMember()>
    Public CantidadUDM As Double
    <DataMember()>
    Public CantidadPicking As String
    <DataMember()>
    Public GestionaUbicacionDestino As String
    <DataMember()>
    Public UbicacionDestinoPropuesta As String

End Class


<DataContract()>
<Serializable()>
Public Class GenerarTraslado
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumeroTraslado As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public Lineas As List(Of LineasPicking)
    <DataMember()>
    Public Bultos As List(Of BultosPicking)
    <DataMember()>
    Public Palets As List(Of PaletsPicking)
End Class


<DataContract()>
<Serializable()>
Public Class ListasRecuentoInventario
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumeroInterno As String
    <DataMember()>
    Public Numero As String
    <DataMember()>
    Public Comentario As String
    <DataMember()>
    Public Fecha As String
End Class

<DataContract()>
<Serializable()>
Public Class ListasRecuentoInventarioCabecera
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumeroInterno As String

    <DataMember()>
    Public Lineas As List(Of ListasRecuentoInventarioDetalle)
End Class

<DataContract()>
<Serializable()>
Public Class ListasRecuentoInventarioDetalle
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public NumeroLinea As String
    <DataMember()>
    Public Articulo As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public CodUbicacion As String
    <DataMember()>
    Public CantidadTeorica As Double
    <DataMember()>
    Public CantidadContada As Double
    <DataMember()>
    Public Verificado As String
    '<DataMember()>
    'Public EAN As List(Of CodigoEAN)
End Class

<DataContract()>
<Serializable()>
Public Class CodigoEAN
    <DataMember()>
    Public EAN As String
End Class

<DataContract()>
<Serializable()>
Public Class CodigoQR
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public CodProveedor As String
    <DataMember()>
    Public NombreProveedor As String
    <DataMember()>
    Public Documento As String
    <DataMember()>
    Public CodArticulo As String
    <DataMember()>
    Public NombreArticulo As String
    <DataMember()>
    Public UbicacionQR As String
    <DataMember()>
    Public Unidades1 As String
    <DataMember()>
    Public Unidades2 As String
    <DataMember()>
    Public FechaEntrada As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public UDM1 As String
    <DataMember()>
    Public UDM2 As String
    <DataMember()>
    Public Actual As List(Of ArticulosQR)
End Class

<DataContract()>
<Serializable()>
Public Class ArticulosQR
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public Cantidad As Double
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public Cantidad2 As Double
End Class


<DataContract()>
<Serializable()>
Public Class LotesImprimir

    <DataMember()>
    Public SEL As String
    <DataMember()>
    Public TipoEtiqueta As String
    <DataMember()>
    Public Impresora As String
    <DataMember()>
    Public CodProveedor As String
    <DataMember()>
    Public NombreProveedor As String
    <DataMember()>
    Public CodArticulo As String
    <DataMember()>
    Public NombreArticulo As String
    <DataMember()>
    Public Ubicacion As String
    <DataMember()>
    Public Unidades1 As String
    <DataMember()>
    Public Unidades2 As String
    <DataMember()>
    Public FechaEntrada As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public UDM1 As String
    <DataMember()>
    Public UDM2 As String
    <DataMember()>
    Public SysNumber As String

End Class

<DataContract()>
<Serializable()>
Public Class ListaLotesImprimir
    <DataMember()>
    Public Resultado As String
    <DataMember()>
    Public LotesImprimir As List(Of LotesImprimir)

End Class
