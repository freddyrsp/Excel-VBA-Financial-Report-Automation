Attribute VB_Name = "App"
Option Explicit
Option Base 1
Option Private Module

Sub Mcro_IniciaFrm_EdoFinanciero()

'Procedimiento para preparar el formulario
'Elaborado por FREDDY SANCHEZ

'Desactivamos propiedades de la aplicación
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Creamos control de errores
On Error GoTo finalizar

'Dimensionamos las variables
Dim H1 As Worksheet
Dim H2 As Worksheet
Dim H3 As Worksheet
Dim Lista(5, 2) As String
Dim Myfecha As Date
Dim Myfecha1 As Date
Dim Contador As Integer
Dim DigitoCuenta As Range

'Asignamos datos a las variables
Set H1 = Hoja4
Set H2 = Hoja2
Set H3 = Hoja3
Set DigitoCuenta = H1.Range("rngAux_DigitoCuenta")

'Asignamos datos al array
Lista(1, 1) = "AUXILIAR"
Lista(2, 1) = "CUENTA"
Lista(3, 1) = "GRUPO"
Lista(4, 1) = "CLASE"
Lista(5, 1) = "TIPO"
Lista(1, 2) = DigitoCuenta.Cells(5, 1) + DigitoCuenta.Cells(4, 1) + DigitoCuenta.Cells(3, 1) + DigitoCuenta.Cells(2, 1) + DigitoCuenta.Cells(1, 1)
Lista(2, 2) = DigitoCuenta.Cells(4, 1) + DigitoCuenta.Cells(3, 1) + DigitoCuenta.Cells(2, 1) + DigitoCuenta.Cells(1, 1)
Lista(3, 2) = DigitoCuenta.Cells(3, 1) + DigitoCuenta.Cells(2, 1) + DigitoCuenta.Cells(1, 1)
Lista(4, 2) = DigitoCuenta.Cells(2, 1) + DigitoCuenta.Cells(1, 1)
Lista(5, 2) = DigitoCuenta.Cells(1, 1)


'Configuramos propiedades del formulario
    Frm012_EdoFinanciero.ComboBox1.ColumnCount = 1
    Frm012_EdoFinanciero.ComboBox1.List = Lista
    Frm012_EdoFinanciero.ComboBox1.Value = "CUENTA"

    Frm012_EdoFinanciero.ComboBox2.ColumnCount = 1
    Frm012_EdoFinanciero.ComboBox2.Value = H1.Range("rngUsr_FinPeriodo").Value
    
    For Contador = 0 To 11
    Myfecha = Application.WorksheetFunction.EoMonth(H1.Range("rngUsr_InicioPeriodo"), Contador)
    Myfecha1 = Application.WorksheetFunction.EoMonth(H1.Range("rngUsr_InicioPeriodo"), (Contador - 1))
    
    With Frm012_EdoFinanciero.ComboBox2
        .AddItem
        .List(Contador, 0) = Myfecha
        .List(Contador, 1) = Myfecha1
    End With
    Next
    
    With Frm012_EdoFinanciero.ProgressBar1
        .Value = 0
        .Min = 0
        .Max = 100
        .Visible = False
    End With
    
finalizar:

'Control de errores
If Err.Number <> 0 Then
MsgBox Err.Number & Err.Description

End If
    
'Activamos propiedades de la Aplicación

Application.ScreenUpdating = True
Application.DisplayAlerts = True
       
End Sub

Sub Mcro_EdoFinanciero()

'Procedimiento para Actualizar estados financieros
'Elaborado por FREDDY SANCHEZ

'Desactivos propiedades de la aplicación
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'Control de errores
On Error GoTo finalizar
ThisWorkbook.Activate

'Dimensionamos las variables
Dim L1 As Workbook: Set L1 = ThisWorkbook
Dim H1 As Worksheet: Set H1 = Hoja2
Dim H2 As Worksheet: Set H2 = Hoja3
Dim H3 As Worksheet: Set H3 = L1.ActiveSheet
Dim H4 As Worksheet: Set H4 = Hoja4
Dim RangoDiario As Range
Dim RangoCuentas As Range
Dim RangoCuentasEdo As Range
Dim UfilaCuentas As Integer
Dim UfilaDiario As Integer
Dim FilaCuentas As Integer
Dim FilaDiario As Integer
Dim SaldoAcumulado As Double
Dim ContadorItems As Integer: ContadorItems = 1
Dim ContadorBarra As Integer
Dim SMDebe As Double
Dim SMHaber As Double
Dim SIDebe As Double
Dim SIHaber As Double
Dim Filtro() As Variant
Dim Diario() As Variant
Dim Cuentas() As Variant
Dim ContadorLista1 As Integer
Dim ContadorLista2 As Integer
Dim NivelEncabezado As String
Dim NivelDetalle As Integer
Dim FechaInicial As Date
Dim FechaFinal As Date

'Ordenamos tabla del plan unico de cuentas
H1.ListObjects("Tblusr_cuentas").Sort.SortFields.Clear
H1.ListObjects("Tblusr_cuentas").Sort.SortFields.Add Key:=Range("Tblusr_cuentas[[#All],[CODIGO]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers

With H1.ListObjects("Tblusr_cuentas").Sort
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Asignamos datos de la tabla del diario y plan unico de cuentas a las variables de rangos
Set RangoDiario = H2.Range("rngAux_diario")
Set RangoCuentas = H1.Range("rngAux_cuentas")
Set RangoCuentasEdo = H3.Range("A1:A200")

'Asignamos datos del diario y plan unico de cuentas a las matrices (mejora el rendimiento)
Diario = RangoDiario
Cuentas = RangoCuentas

'Protegemos la hoja contra escritura
H3.Protect userinterfaceonly:=True

'Limpiados los datos de la hoja informe
H3.Range("C1:D200").ClearContents

'Asignamos los datos del formulario a las variables
ContadorLista1 = Frm012_EdoFinanciero.ComboBox2.ListIndex
ContadorLista2 = Frm012_EdoFinanciero.ComboBox1.ListIndex
NivelEncabezado = Frm012_EdoFinanciero.ComboBox1.List(ContadorLista2, 0)
NivelDetalle = Frm012_EdoFinanciero.ComboBox1.List(ContadorLista2, 1)
FechaInicial = H4.Range("rngUsr_InicioPeriodo")
FechaFinal = Frm012_EdoFinanciero.ComboBox2.List(ContadorLista1, 0)
UfilaCuentas = Application.WorksheetFunction.CountA(RangoCuentas.Columns(3))
UfilaDiario = Application.WorksheetFunction.CountA(RangoDiario.Columns(3))

'Cambiamos propiedad a Visible
Frm012_EdoFinanciero.ProgressBar1.Visible = True

'Imprimimos la fecha de la actualizacion en la hoja de calculo
H3.Cells(3, 2) = FechaFinal

'Redimensionamos matrices que almacena los datos calculados
ReDim Filtro(5000, 2)

'Realizamos recorrido del plan unico de cuentas
For Each RangoCuentasEdo In RangoCuentasEdo.Rows
    
    'Realizamos recorrido del diario
    For FilaDiario = 1 To UBound(Diario)
        
        'Evaluamos nivel de detalle asignado por el usuario en el formulario
        If (Mid(Diario(FilaDiario, 3), 1, NivelDetalle) = _
        Mid(RangoCuentasEdo, 1, NivelDetalle)) _
        And Diario(FilaDiario, 1) <= FechaFinal Then
           
           'Coincide acumulamos el saldo de la cuenta
           SIDebe = Diario(FilaDiario, 6) + SIDebe
           SIHaber = Diario(FilaDiario, 7) + SIHaber
           
           If Diario(FilaDiario, 1) <= FechaInicial Then
                SMDebe = Diario(FilaDiario, 6) + SMDebe
                SMHaber = Diario(FilaDiario, 7) + SMHaber
            
           End If
        Else
        'No coincide no dirigimos a lineal
        GoTo linea1
                  
        End If
                        
                        
        Filtro(ContadorItems, 1) = SMDebe - SMHaber
        Filtro(ContadorItems, 2) = SIDebe - SIHaber
        
linea1:
       
'No coincide continuamos con la siguiente cuenta
       Next FilaDiario
       
         SIDebe = 0
         SIHaber = 0
         SMDebe = 0
         SMHaber = 0
            
         ContadorItems = ContadorItems + 1
         ContadorBarra = ContadorBarra + 1
         
         If ContadorBarra >= 200 Then
             Frm012_EdoFinanciero.ProgressBar1.Value = 100
         Else
             Frm012_EdoFinanciero.ProgressBar1.Value = (ContadorBarra / 200) * 100
         End If

Next RangoCuentasEdo


H3.Range("C1:D" & ContadorItems) = Filtro
 
MsgBox "¡¡¡Actualización Lista!!!"


'Control de errores
finalizar:
If Err.Number <> 0 Then
MsgBox Err.Number & Err.Description

End If

'Salimos del formulario
Unload Frm012_EdoFinanciero

'Activamos propiedades de la aplicación
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub Mcro_EmitirPdf()
'Procedimiento para emitir los informes en pdf
'Elaborado por FREDDY SANCHEZ

'Desactivamos propiedades de la aplicación
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'Diensionamos variables
Dim pathtosave As String
Dim H1 As Worksheet
Dim H2 As Worksheet

'Asignamos valores a las variables
Set H1 = Hoja1 'Hoja informe
Set H2 = Hoja4 'Hoja datos
pathtosave = Application.ThisWorkbook.Path & Application.PathSeparator & _
"Informe-Financiero-" & Format(H1.Cells(3, 2), "ddmmyyyy") & ".pdf"
        
'Escribimos la ruta en la hoja de datos del archivo
H2.Range("rngUsr_adjunto") = "Informe-Financiero-" & Format(H1.Cells(3, 2), "ddmmyyyy")

'Exportamos a pdf la hoja de informes
With Hoja1
    .ExportAsFixedFormat Type:=xlTypePDF, Filename:=pathtosave, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, OpenAfterPublish:=False
End With

MsgBox "Listo"

'Activamos propiedades a la aplicación
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub Mcro_EnvioMasivo()
'Procedimiento para enviar los informes fiancieros mediante Gmail
'Elaborado por FREDDY SANCHEZ

'Desactivamos propiedades de la aplicación
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

On Error GoTo finalizar

'Dimensionamos variables
Dim H1 As Worksheet
Dim remitente As String, clave As String, destinatario As String, destino As String, adjunto As String
Dim matriz() As Variant
Dim i As Long

'Asignamos valores a las variables
Set H1 = Hoja4
matriz = H1.Range("rngUsr_destinatario")
adjunto = H1.Range("rngUsr_adjunto")
remitente = H1.Range("rngUsr_email")
clave = H1.Range("rngUsr_clave")
i = 1

'Realizamos recorrido de las celdas con destinatarios
Do Until i >= UBound(matriz, 1)
    
    If CStr(matriz(i, 1)) <> "" And CStr(matriz(i, 2)) <> "" Then
    
        destinatario = matriz(i, 1)
        destino = matriz(i, 2)
       
        Call conexion_correo(remitente, clave, destinatario, destino, adjunto)
    
    End If
    
    i = i + 1

Loop

finalizar:

If Err.Number <> 0 Then
    
    MsgBox Err.Number & Err.Description

Else

    MsgBox "File sent without Errors"
    
End If


'Activamos propiedades de la aplicación
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub conexion_correo(remitente As String, clave As String, destinatario As String, destino As String, adjunto As String)
' Realizado por: Freddy Sánchez
' Modificación : 13/12/2022
    Dim email As CDO.Message
    Dim asunto As String, cuerpo As String
    
    asunto = "Informe Financiero"
    cuerpo = "Estimado," & " " & destinatario & Chr(13) & Chr(10) & _
    "El dia de hoy," & " " & Format(Date, "dd-mm-yyyy") & ", " & "le remitimos por este medio" & Chr(13) & Chr(10) & _
    "el Informe financiero actualizado al cierre de periodo anterior." & Chr(13) & Chr(10) & _
    "Para cualquier información puede dirigirse a las oficinas de administración del plantel" & Chr(13) & Chr(10) & _
    "Saludos Cordiales," & Chr(13) & Chr(10) & _
    "Lcdo. Freddy Sánchez"
   
    Set email = New CDO.Message
   
    With email.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(465)
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Abs(1)
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = remitente
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = clave
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        
    End With
    
    With email
        .To = destino
        .From = remitente
        .Subject = asunto
        .TextBody = cuerpo
        .AddAttachment (Application.ThisWorkbook.Path & Application.PathSeparator & adjunto & ".pdf")
        .Configuration.Fields.Update
         On Error Resume Next
        .Send
         
    End With
        
      
End Sub


Sub callback1(control As IRibbonControl)

    Dim H1 As Worksheet
    Set H1 = ActiveSheet

    If H1.Name = "Informe" Then
        Frm012_EdoFinanciero.Show
    End If
    
End Sub

Sub callback2(control As IRibbonControl)

    Dim H1 As Worksheet
    Set H1 = ActiveSheet
        
    If H1.Name = "Informe" Then
       Call Mcro_EmitirPdf
    End If

End Sub


Sub callback3(control As IRibbonControl)

   Dim H1 As Worksheet
    Set H1 = ActiveSheet
        
    If H1.Name = "Informe" Then
       Call Mcro_EnvioMasivo
    End If

End Sub




