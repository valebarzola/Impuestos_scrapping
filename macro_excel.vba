' ============================================================================
' MACRO VBA PARA EXCEL - Consumir API de Cotizaciones AFIP
' ============================================================================
' 
' INSTRUCCIONES DE INSTALACIÓN:
' 1. En Excel, presionar Alt+F11 para abrir el Editor de VBA
' 2. En el panel izquierdo, hacer clic derecho en "ThisWorkbook" -> 
'    "Insert Module"
' 3. Copiar TODO el código de abajo en el nuevo módulo
' 4. Cambiar la URL en la constante API_URL por la IP/DNS de tu VM
' 5. Guardar libro como .xlsm (Macro-enabled)
' 6. Habilitar macros cuando abra el archivo
'
' CONFIGURACIÓN DEL WORKSHEET:
' Ejemplo de estructura de datos:
' | Columna A:          | Columna B:          | Columna C:           |
' | Nro de Embarque     | Fecha Oficialización| Tipo Cambio Comprador|
' | PE-001             | 22/11/2024         | [Se llena automático]|
' | PE-002             | 23/11/2024         | [Se llena automático]|
'
' ============================================================================

' Importar referencias necesarias (en Editor VBA: Tools -> References)
' ✓ Microsoft XML, v3.0 (or higher)
' ✓ Microsoft Excel Object Library (ya está)

Option Explicit

' ============================================================================
' CONSTANTES DE CONFIGURACIÓN
' ============================================================================

' Cambiar esta URL a la IP/DNS de tu vm donde corre la API
' Ejemplos: "http://192.168.1.100:8000" o "http://api.local:8000"
' Si corre localmente en tu PC: "http://localhost:8000"
Const API_URL As String = "http://192.168.1.100:8000"

' Número de la columna donde están las fechas de oficialización
' A=1, B=2, C=3, etc.
Const COLUMNA_FECHAS As Long = 2

' Número de la columna donde se escribirá el tipo de cambio comprador
Const COLUMNA_RESULTADO As Long = 3

' Moneda por defecto (puede cambiar si vende otras monedas)
Const MONEDA_POR_DEFECTO As String = "DOL"

' Timeout en segundos para la llamada a la API
Const TIMEOUT_SEGUNDOS As Long = 10

' ============================================================================
' FUNCIONES PRINCIPALES
' ============================================================================

' Función para hacer GET HTTP a la API
Function ConsultarAPI(URLCompleta As String) As String
    On Error GoTo ErrorHandler
    
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Configurar timeout
    xmlHttp.SetTimeouts TIMEOUT_SEGUNDOS * 1000, TIMEOUT_SEGUNDOS * 1000, TIMEOUT_SEGUNDOS * 1000, TIMEOUT_SEGUNDOS * 1000
    
    ' Hacer la solicitud GET
    xmlHttp.Open "GET", URLCompleta, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send
    
    ' Verificar que la respuesta sea OK (status 200)
    If xmlHttp.Status = 200 Then
        ConsultarAPI = xmlHttp.ResponseText
    Else
        ConsultarAPI = "ERROR:" & xmlHttp.Status & ":" & xmlHttp.ResponseText
    End If
    
    Set xmlHttp = Nothing
    Exit Function
    
ErrorHandler:
    ConsultarAPI = "ERROR:CONEXION:" & Err.Description
End Function

' Función para extraer un valor de una respuesta JSON simple
' Usa búsqueda de texto (no es un parser JSON formal, pero funciona para esta API)
Function ExtraerDelJSON(jsonText As String, campo As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim buscar As String
    
    ' Buscar el patrón: "campo": valor
    buscar = """" & campo & """" & ":"
    startPos = InStr(1, jsonText, buscar)
    
    If startPos = 0 Then
        ExtraerDelJSON = ""
        Exit Function
    End If
    
    startPos = startPos + Len(buscar)
    
    ' Saltar espacios en blanco
    While Mid(jsonText, startPos, 1) = " "
        startPos = startPos + 1
    Wend
    
    ' Detectar si es string (") o número
    If Mid(jsonText, startPos, 1) = """" Then
        ' Es string: buscar cierre de comillas
        startPos = startPos + 1
        endPos = InStr(startPos, jsonText, """")
    Else
        ' Es número, booleano o null: buscar coma o cierre de llave
        endPos = InStr(startPos, jsonText, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, "}")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, "]")
        endPos = endPos - 1
    End If
    
    ExtraerDelJSON = Trim(Mid(jsonText, startPos, endPos - startPos))
End Function

' Función para obtener el tipo de cambio de la API
Function ObtenerTipoCambio(fechaOficializacion As String, moneda As String) As Variant
    Dim URLCompleta As String
    Dim respuesta As String
    Dim resultado As Variant
    
    ' Construir la URL
    URLCompleta = API_URL & "/cotizacion?fecha=" & fechaOficializacion & "&moneda=" & moneda
    
    ' Mostrar mensaje en la barra de estado (feedback al usuario)
    Application.StatusBar = "Consultando API para " & fechaOficializacion & "..."
    
    ' Hacer la consulta
    respuesta = ConsultarAPI(URLCompleta)
    
    ' Verificar si hay error
    If Left(respuesta, 6) = "ERROR:" Then
        ObtenerTipoCambio = Array("ERROR", respuesta)
        Exit Function
    End If
    
    ' Extraer el tipo de cambio comprador del JSON
    Dim tipoCambio As String
    tipoCambio = ExtraerDelJSON(respuesta, "tipo_cambio_comprador")
    
    If tipoCambio = "" Then
        ObtenerTipoCambio = Array("ERROR", "No se encontró tipo_cambio_comprador en respuesta")
        Exit Function
    End If
    
    ' Retornar éxito con el valor
    ObtenerTipoCambio = Array("OK", CDbl(tipoCambio))
End Function

' Función para validar que el formato de fecha sea DD/MM/YYYY
Function ValidarFecha(fechaStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim partes() As String
    Dim dia, mes, anio As Integer
    
    partes = Split(fechaStr, "/")
    
    If UBound(partes) <> 2 Then
        ValidarFecha = False
        Exit Function
    End If
    
    dia = CLng(partes(0))
    mes = CLng(partes(1))
    anio = CLng(partes(2))
    
    ' Validación básica
    If mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Or anio < 2000 Then
        ValidarFecha = False
        Exit Function
    End If
    
    ValidarFecha = True
    Exit Function
    
ErrorHandler:
    ValidarFecha = False
End Function

' ============================================================================
' MACROS QUE SE EJECUTAN EN EL WORKSHEET
' ============================================================================

' Esta macro se ejecuta automáticamente cuando hay cambios en el worksheet
' Se dispara para la columna de fechas (COLUMNA_FECHAS)
Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Me ' El worksheet que contiene esta macro
    
    ' Ignorar cambios fuera de la columna de fechas
    If Intersect(Target, ws.Columns(COLUMNA_FECHAS)) Is Nothing Then
        Exit Sub
    End If
    
    ' Desactivar recálculos temporalmente para mejorar rendimiento
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Dim celda As Range
    Dim fila As Long
    Dim fechaStr As String
    Dim resultado As Variant
    Dim celdaResultado As Range
    
    ' Procesar cada celda modificada
    For Each celda In Target
        fila = celda.Row
        fechaStr = Trim(celda.Value)
        
        ' Ignorar celdas vacías o cabeceras
        If fechaStr <> "" And Not celda.Row = 1 Then
            Set celdaResultado = ws.Cells(fila, COLUMNA_RESULTADO)
            
            ' Validar formato de fecha
            If Not ValidarFecha(fechaStr) Then
                celdaResultado.Value = "ERROR: Fecha inválida (Use DD/MM/YYYY)"
                celdaResultado.Interior.Color = RGB(255, 200, 200) ' Fondo rojo claro
                GoTo SiguienteCelda
            End If
            
            ' Obtener tipo de cambio
            resultado = ObtenerTipoCambio(fechaStr, MONEDA_POR_DEFECTO)
            
            If resultado(0) = "OK" Then
                ' Escribir el valor exitosamente
                celdaResultado.Value = resultado(1)
                celdaResultado.NumberFormat = "0.00"
                celdaResultado.Interior.Color = RGB(200, 255, 200) ' Fondo verde claro
                celdaResultado.Font.Color = RGB(0, 128, 0) ' Texto verde
            Else
                ' Mostrar error en la celda
                celdaResultado.Value = resultado(1)
                celdaResultado.Interior.Color = RGB(255, 200, 200) ' Fondo rojo claro
                celdaResultado.Font.Color = RGB(128, 0, 0) ' Texto rojo
            End If
        End If
        
SiguienteCelda:
    Next celda
    
    ' Limpiar mensaje de estado
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error en Worksheet_Change: " & Err.Description, vbCritical
End Sub

' Macro auxiliar para limpiar todos los resultados (botón de reset)
Sub LimpiarResultados()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, COLUMNA_FECHAS).End(xlUp).Row
    
    ws.Range(ws.Cells(2, COLUMNA_RESULTADO), ws.Cells(ultimaFila, COLUMNA_RESULTADO)).Clear
    ws.Range(ws.Cells(2, COLUMNA_RESULTADO), ws.Cells(ultimaFila, COLUMNA_RESULTADO)).Interior.ColorIndex = xlNone
    
    MsgBox "Resultados limpiados", vbInformation
End Sub

' Macro para probar conexión a la API
Sub VerificarConexionAPI()
    Dim respuesta As String
    Dim URLSalud As String
    
    URLSalud = API_URL & "/salud"
    respuesta = ConsultarAPI(URLSalud)
    
    If Left(respuesta, 6) = "ERROR:" Then
        MsgBox "❌ No se puede conectar a la API" & vbCrLf & vbCrLf & respuesta, vbCritical, "Error de Conexión"
    Else
        MsgBox "✓ Conexión exitosa a " & API_URL & vbCrLf & vbCrLf & respuesta, vbInformation, "Conexión OK"
    End If
End Sub

' Macro para rellenar cotizaciones en lote (de todo el rango a la vez)
Sub RellenarLote()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ultimaFila As Long
    Dim fila As Long
    Dim fechaStr As String
    Dim resultado As Variant
    Dim contador As Long
    Dim errores As Long
    
    Application.ScreenUpdating = False
    
    ultimaFila = ws.Cells(ws.Rows.Count, COLUMNA_FECHAS).End(xlUp).Row
    contador = 0
    errores = 0
    
    For fila = 2 To ultimaFila
        fechaStr = Trim(ws.Cells(fila, COLUMNA_FECHAS).Value)
        
        If fechaStr <> "" Then
            If ValidarFecha(fechaStr) Then
                resultado = ObtenerTipoCambio(fechaStr, MONEDA_POR_DEFECTO)
                
                If resultado(0) = "OK" Then
                    ws.Cells(fila, COLUMNA_RESULTADO).Value = resultado(1)
                    ws.Cells(fila, COLUMNA_RESULTADO).NumberFormat = "0.00"
                    ws.Cells(fila, COLUMNA_RESULTADO).Interior.Color = RGB(200, 255, 200)
                    contador = contador + 1
                Else
                    ws.Cells(fila, COLUMNA_RESULTADO).Value = resultado(1)
                    ws.Cells(fila, COLUMNA_RESULTADO).Interior.Color = RGB(255, 200, 200)
                    errores = errores + 1
                End If
            Else
                ws.Cells(fila, COLUMNA_RESULTADO).Value = "Fecha inválida"
                ws.Cells(fila, COLUMNA_RESULTADO).Interior.Color = RGB(255, 200, 200)
                errores = errores + 1
            End If
        End If
    Next fila
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Lote completado:" & vbCrLf & _
           "✓ Éxitos: " & contador & vbCrLf & _
           "✗ Errores: " & errores, vbInformation, "Resultado"
End Sub

' ============================================================================
' NOTAS IMPORTANTES:
' ============================================================================
' 
' 1. EVENTO Worksheet_Change:
'    - Se ejecuta automáticamente cuando modifica celdas en COLUMNA_FECHAS
'    - Consulta la API y llena COLUMNA_RESULTADO con el tipo de cambio
'    - Colorea la celda en verde si OK, rojo si hay error
'
' 2. SEGURIDAD:
'    - Para usar esta macro deben estar habilitadas las macros en Excel
'    - No ejecute este archivo si no confía en la fuente
'
' 3. CONFIGURACION:
'    - Cambiar API_URL al inicio del código con IP de tu VM
'    - Cambiar COLUMNA_FECHAS y COLUMNA_RESULTADO si la estructura es diferente
'
' 4. INSTALACION DE REFERENCIAS (si hay errores de XMLHTTP):
'    Alt+F11 -> Tools -> References -> Buscar "Microsoft XML, v3.0"
'    Marcar checkbox y OK
'
' 5. BOTONES ÚTILES PARA AGREGAR:
'    - Botón "Limpiar Resultados" -> Llama Sub LimpiarResultados()
'    - Botón "Verificar Conexión" -> Llama Sub VerificarConexionAPI()
'    - Botón "Rellenar Lote" -> Llama Sub RellenarLote()
'
