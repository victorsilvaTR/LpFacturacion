Attribute VB_Name = "Actualizaciones"
Option Explicit
'° ¿

'29-03-2017 Se corrige error en emisión de Factura Exenta y Nota de Crédito Exenta
'30-03-2017 Se permite modificar el precio de un producto cuando este tiene valor cero en la lista de productos
'           Se agrega posibilidad de tener más de un correo Receptor de Doc. Electrónico. Estos deben ir separados por ";". Se agrega ToolTip en campo mail de recepción de DTE para indicar que se puede indicar más de un mail separándolo por ";"
'           Se agrega un campo de observaciones para un DTE al pie del documento. Las observaciones por omisión se ingresan en la ficha de la empresa emisora de documento. Estas observaciones fueron solicitadas para poner la cuenta corriente de pago del documento.
'31-03-2017 Se agrega equivalencias y conversión de monedas en la ventana principal y en la ventana de emisión de DTE. Guarda las monedas utilizadas en la última conversión realizada.
'20-04-2017 Se agrega un detalle extendido y más claro del estado de un documento a partir de la traza que entrega Acepta
'10-05-2017 Se cambia código de error en la función AcpAutenticar por cambio inesperado en web service provisto por Acepta
'11-05-2017 Se agrega EntRelacionada a Entidades, para facilitar la importación desde Contabilidad
'           Se genera actualizador 1.0.3
'11-05-2017 Se agrega validaciones al ingreso de referencias de un DTE
'22-06-2018 Se agrega posibilidad de ordenar lista de DTE Emitidos
'09/07/2018 Se agrega el ordenamiento de columnas a la lista de productos a través del click en el título de la columna
'           Se agrega búsqueda por nombre de producto a lista de productos
'12/07/2018 Se agrega reporte de ventas de productos con diversos filtros
'16/07/2018 Se corrige desplieque de la traza de un DTE (Botón Ver Detalle Completo) en la ventana Detalle Estado DTE
'07/08/2018 Se corrige el Save de las referencias de un DTE almacenando el ID del DTE de Referencia, si corresponde
'           Se corrige manejo de mensaje cuando se selecciona nota de crédito y, al seleccionar el documento de referencia, se da Cancel
'           Se agrega botón para eliminar una referencia, ubicado arriba a la derecha de la lista de referencias de un DTE
'           Se agrega la opción en la ventana de Nuevo DTE, para copiar un DTE con dos alternativas: último DTE emitido o DTE emotido previamente, en cuyo caso se debe seleccionar de la lista de DTE Emitidos
'21/08/2018 Se valida e-amil del receptor del documento, al momento de modificar una entidad y al momento de emitr un documento. Esto para evitar además, que no se imprima la observación den el DTE
'23/08/2018 Se mejora el manejo de las Guias de Despacho utilizando un tipo de documento distinto y Libro Otros
'24/08/2018 Se agrega botón a la ventana Adm. de DTE para copiar la URL del documento seleccionado al PortaPapeles
'           Se genera Ejecutable de Prueba
'03/09/2018 Se mejora reporte de ventas de productos
'27/09/2018 Se agrega Factura de exportación y sus datos asociados: Cláusulas de Venta, paiese, puestors, etc.
'30/01/2019 Se corrigen problemas con emisión  de Guía de Despacho
'14/03/2019 Se mejora la importación de productos
'22/03/2019 Se modifica datos de vehículos y choferes para guías de despacho
'05/04/2019 Se agrega notas de crédito de exportación y notas de débito de exportación
'06/05/2019 Se agrega opción para importar desde Acepta los DTE Recibidos y se habilita listado correspondiente
'21/06/2019 Se agrega función de previsualización de un DTE antes de emitirlo
'29/08/2019 Se actualizan los códigos de actividad económoca de acuerdo a los cambios del SII en nov 2018
'02/10/2020 Se corrige error al generar Xml de DTE de documentos de exportación (debe decir Exportaciones en vez de Documento)
'           Se agrega observaciones a Guia de Despacho
'16/02/2021 Se permite una FAV con todos los valoes exentos y sin IVA, a solicitud de Katherine
