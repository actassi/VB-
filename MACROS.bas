Attribute VB_Name = "MACROS"
Sub Macro2_creditos()
'
' Macro2_creditos Macro
'

'

Sheets("DATOS").Select
    

    ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array("CCERR SUCATA SA 30708270632 CIRC.CERRADO", _
"Liq.Resc 10 00060202", _
"TEF DATANET PR TERMINAL PUERTO RO 30708096063", _
"Transf. Factur EDECA SA 33515507219", _
"CCERR MICELI MADE 30529414923 CIRC.CERRADO", _
"TEF DATANET PR MAGEVA SRL 30711535639", _
"TEF DATANET PR FIDEICOMISO NAVES 30717386538", _
"ACREDITACION CHEQUE REMESAS Suc.:369", _
"TEF DATANET PR Crucijuegos insumo 30712064230", _
"TRANSF SINDICATO 30630359356 FAC FACTURAS 307", _
"ACREDITACION CHEQUE REMESAS Suc.:817", _
"CCERR JOCKEY CLUB 30500237305 CIRC.CERRADO", _
"CR DEPOSITO CANJE INTERNO Suc.:793", _
"TRANSF MASTRODIC 30716044102 FAC FACTURAS", _
"CCERR QUATRO SRL 30712074295 CIRC.CERRADO", _
"TRANSF CONS. PRO 30600872121 FAC FACTURAS 100", _
"TEF DATANET PR DISAL SA 30621181838", _
"CR PLAZO FIJO MONLINEÿÿ", _
"22 Liq.Resc 62199 38466", _
"22 Liq.Resc 61394 38466", "22 Liq.Resc 60793 60202", _
"22 Liq.Resc 51715 60202", _
"ACREDITACION CHEQUE REMESAS Suc.:761", _
"CCERR MUNICIPALID 30659729489 CIRC.CERRADO", "ACREDITACION CHEQUE REMESAS Suc.:817", _
"13 Liq.Resc 27310 60202", "TEF DATANET PR MINISTERIO DE LA P 30999184401", "CCERR JUNTAS ILLI 30521382984 CIRC.CERRADO", "TEF DATANET PR MERCATOR S.A. 30583210276", _
"TEF DATANET PR TERMINAL PUERTO RO 30708096063", "TRATM LLANEADOS OT 30711860300 FAC FACTURAS", "CR PLAZO FIJO MONLINE*"), Operator:=xlFilterValues


        
    
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("1.DEPOSITOS").Select
       
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("DATOS").Select
    
    ActiveSheet.Range("$C$1:$C$16100").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
    
End Sub

Sub Macro4_SIRCREB()
'
' Macro4_SIRCREB Macro
'
Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array( _
        "AJ IIBB STAFE SIRCREB Suc.:811", "IIBB SANTA FE SIRCREB Suc.:811", "IIBB SANTA FE SIRCREB Suc.:811"), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("4.SIRCREB").Select
       
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("DATOS").Select
    
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
'
End Sub
Sub Macro5_fondo()
'
' Macro5_fondo Macro
'
Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array( _
        "DB PAGO REMUNERACIONES", "DEB.FNDO.CESE LABORAL"), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("5.FONDO Y SUELDOS").Select
       
    Range("A1").Select
    ActiveSheet.Paste
   
    Sheets("DATOS").Select
    
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
'
End Sub
Sub Macro6_SUELDOS()
'
' Macro6_SUELDOS Macro
'
Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array( _
        ""), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("6.SUELDOS").Select
       
    Range("A1").Select
    ActiveSheet.Paste
   
    Sheets("DATOS").Select
   
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
'
End Sub
Sub Macro8_GASTOSIVA()
'
' Macro8_GASTOSIVA Macro
'
Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array( _
         "COMMOVCAJ - COMISIàN MOVIMIENTO POR CAJA", "COMISIàN MOVIMIENTO POR CAJA", "COMISIàN ADM. VALORES AL COBRO CC Suc.:817", "  DEBITO FISCAL IVA PERCEPCION", "Comision Trf. MacrOL E-set", "COMISION TRANSFERE", "COMISIàN ADMINISTRACIàN DE CHEQUERA", "INTER.ADEL.CC C/ACUERD", "COMISION CHQ PAG CLEARING", "AJ COMISION MANTENIMIENTO CUENTAÿÿÿ", "DEVOLUCION COMISIONES VARIAS", "MANTENIMIENTO MENSUAL PAQUETE", "COMISION CHEQUE CONSULTA", "DEBITO FISCAL IVA BASICO", "COMISION TRANSFERENCIAS", "COMISION DB PAGO REMUNERAC"), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("8.GASTOS").Select
       
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("DATOS").Select
    
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
'
End Sub
Sub Macro10_IMPDBCRED()
'
' Macro10_IMPDBCRED Macro
'
Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array("IMPDBCR 25413 S/CR TASA GRAL", _
        "DBCR 25413 S/DB TASA GRAL", "IMPDBCR 25413 S/DB TASA GRAL", "DBCR 25413 S/CR TASA GRAL"), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("10.IMP DEB CRED").Select
       
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("DATOS").Select
    
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
'
End Sub
Sub vaciar()
'
' vaciar Macro
'

'
    Sheets("DATOS").Select
    Range("B2:H1000").Select
    Cells.EntireRow.Hidden = False
       
    Cells.FormatConditions.Delete
    Selection.ClearContents
    
    Sheets("1.DEPOSITOS").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("2.DEBITOS Y CHEQUES").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("3.AFIP").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("4.SIRCREB").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("5.FONDO Y SUELDOS").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("8.GASTOS").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("10.IMP DEB CRED").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("11.LEASING").Select
    Cells.EntireRow.Hidden = False
    Range("A2:G1000").Select
    Selection.ClearContents
    
    Sheets("DATOS").Select
    Range("A1").Select
    
    
End Sub



Sub Macro3_afip()
'
' Macro3_afip Macro
'
Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array( _
        "IMP. AFIP", "AFIP", "TEF DATANET PAGOS AFIP"), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    Sheets("3.AFIP").Select
       
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("DATOS").Select
    
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
'
End Sub
Sub VOLVER()
'
' VOLVER Macro
'

'
    Sheets("BANCO").Select
    Range("A1").Select
End Sub

Sub Macro11_leasing()
'
' Macro11_leasing Macro
'

'
    Sheets("DATOS").Select
    
     ActiveSheet.Range("$E$1:$E$1610").AutoFilter Field:=1, Criteria1:=Array( _
        "Contrato 16802 - GASTOS DE ESCRIBANIA", "Contrato 16802 - Impuesto I.V.A", "Contrato 16802 - GASTOS DE GESTORIA", "Seguros Leasing OPCION"), Operator:=xlFilterValues
        
    Range("A1:G1000").Select
    Selection.Copy
    
    
    Sheets("11.LEASING").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("DATOS").Select
    ActiveSheet.Range("$C$1:$C$161").AutoFilter Field:=1
    Application.CutCopyMode = False
    
    
    Range("A1").Select
    
End Sub

Sub saldoini()

Dim saldoini As Currency
Dim saldouf As Currency
Dim debcreduf As Currency

Sheets("datos").Select
uf = Sheets("datos").Range("G" & Rows.Count).End(xlUp).Row

saldoini = Range("G" & uf) - Range("F" & uf)
Sheets("datos").Range("N2").Value = saldoini


End Sub


Sub AutoSuma()
Dim FilaSumas As Integer
FilaSumas = Range("A" & Rows.Count).End(xlUp).Row + 1
Range("F" & FilaSumas).FormulaLocal = "=SUMA(F2:F" & FilaSumas - 1 & ")"

End Sub

Sub ActualizarEgresos()
    'Establecer una referencia a Microsoft Office 16.0 Access Database Engine Object Library en Herramientas -> Referencias
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strConexion As String
    Dim strSQL As String
    Dim fila As Long
    Dim fechaefec As Date
    Dim nrocheque As String
    
    
    
    'Establecer la cadena de conexión a la base de datos de Access
    strConexion = "C:\Users\admin\Desktop\contable.accdb"
    
    'Abrir la base de datos
    Set db = DAO.OpenDatabase(strConexion)
    
    
    
    'Abrir el recordset para recorrer la tabla de egresos
    Set rs = db.OpenRecordset("SELECT * FROM egresos")
    
    'Recorrer la columna de datos en el archivo de Excel
    With ActiveWorkbook.Sheets("DATOS")
    
    
        For fila = 1 To .Cells(.Rows.Count, "D").End(xlUp).Row
        
            If Cells(fila, "D") = "85" Or Cells(fila, "D") = "2837" Then


           'Modificar los valores de la consulta SQL con los datos de la fila actual

            fechaefec = Format(Range("B" & fila).Value, "mm/dd/yyyy")
            nrocheque = Range("C" & fila).Value
            strSQL = "UPDATE EGRESOS SET EGRESOS.FECHA_EFECTIVA = #" & fechaefec & "# WHERE (((EGRESOS.OBSERVACIONES) Like '*" & nrocheque & "*'));"
           'Ejecutar la consulta SQL
            db.Execute strSQL
  
            End If
           
        Next fila
    End With
    
    'Cerrar el recordset y la base de datos
    rs.Close
    db.Close
    
    'Limpiar la memoria
    Set rs = Nothing
    Set db = Nothing
End Sub


Sub EscribirGastos()
    
    ' Declarar las variables para la conexión
    Dim cn As Object
    Dim rs As Object
   
    
   
    
    ' Crear una instancia de la conexión y abrir la base de datos de Access
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\Users\Admin\Desktop\contable.accdb;Persist Security Info=False;"
    

    ' Crear una instancia del conjunto de registros y especificar la tabla a escribir
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "EGRESOS", cn, adOpenDynamic, adLockOptimistic
    
     With rs
        .AddNew
            .Fields("BALANCE").Value = Sheets("DESGLOSE IVA").Range("L57").Value
            .Fields("EMPRESA").Value = Sheets("DESGLOSE IVA").Range("M57").Value
            .Fields("RUBRO").Value = Sheets("DESGLOSE IVA").Range("N57").Value
            .Fields("CC").Value = Sheets("DESGLOSE IVA").Range("O57").Value
            .Fields("ITEM").Value = Sheets("DESGLOSE IVA").Range("P57").Value
            .Fields("COMPROBANTE").Value = Sheets("DESGLOSE IVA").Range("Q57").Value
            .Fields("BANCO").Value = Sheets("DESGLOSE IVA").Range("R57").Value
            .Fields("FACTURA").Value = Sheets("DESGLOSE IVA").Range("S57").Value
            .Fields("NETO").Value = Sheets("DESGLOSE IVA").Range("T57").Value
            .Fields("NOTA").Value = Sheets("DESGLOSE IVA").Range("U57").Value
            .Fields("TIPO COMP").Value = Sheets("DESGLOSE IVA").Range("V57").Value
            .Fields("FECHA REAL").Value = Sheets("DESGLOSE IVA").Range("W57").Value
            .Fields("FECHA").Value = Sheets("DESGLOSE IVA").Range("X57").Value
            .Fields("FECHA_DB").Value = Sheets("DESGLOSE IVA").Range("Y57").Value
            .Fields("FECHA_EFECTIVA").Value = Sheets("DESGLOSE IVA").Range("Z57").Value
            .Fields("IVA21").Value = Sheets("DESGLOSE IVA").Range("AA57").Value
            .Fields("IVA27").Value = Sheets("DESGLOSE IVA").Range("AB57").Value
            .Fields("IVA10").Value = Sheets("DESGLOSE IVA").Range("AC57").Value
            .Fields("CAJA").Value = Sheets("DESGLOSE IVA").Range("AD57").Value
            .Fields("NC").Value = Sheets("DESGLOSE IVA").Range("AE57").Value
            .Fields("VALORES").Value = Sheets("DESGLOSE IVA").Range("AF57").Value
            .Fields("PARALELO").Value = Sheets("DESGLOSE IVA").Range("AG57").Value
            .Fields("FORMA DE PAGO").Value = Sheets("DESGLOSE IVA").Range("AI57").Value
            .Fields("CUENTA").Value = Sheets("DESGLOSE IVA").Range("AJ57").Value
            .Fields("OBSERVACIONES").Value = Sheets("DESGLOSE IVA").Range("AK57").Value
            .Fields("NO GRAVADO").Value = Sheets("DESGLOSE IVA").Range("AL57").Value
            .Fields("MONOTRIBUTO").Value = Sheets("DESGLOSE IVA").Range("AM57").Value
            .Fields("PERC IB").Value = Sheets("DESGLOSE IVA").Range("AN57").Value
            .Fields("RET GAN").Value = Sheets("DESGLOSE IVA").Range("AO57").Value
            .Fields("PERC IVA").Value = Sheets("DESGLOSE IVA").Range("AP57").Value
            .Fields("BANCO1").Value = Sheets("DESGLOSE IVA").Range("AQ57").Value
            .Fields("BANCO2").Value = Sheets("DESGLOSE IVA").Range("AR57").Value
            .Fields("REVISION").Value = Sheets("DESGLOSE IVA").Range("AS57").Value


        .Update
    End With
    
    ' Cerrar la conexión y el conjunto de registros
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
End Sub


Sub EscribirSircreb()
    
    ' Declarar las variables para la conexión
    Dim cn As Object
    Dim rs As Object
   
    
   
    
    ' Crear una instancia de la conexión y abrir la base de datos de Access
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\Users\admin\Desktop\contable.accdb;Persist Security Info=False;"
    

    ' Crear una instancia del conjunto de registros y especificar la tabla a escribir
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "EGRESOS", cn, adOpenDynamic, adLockOptimistic
    
     With rs
        .AddNew
                .Fields("BALANCE").Value = Sheets("4.SIRCREB").Range("Q2").Value
                .Fields("EMPRESA").Value = Sheets("4.SIRCREB").Range("R2").Value
                .Fields("RUBRO").Value = Sheets("4.SIRCREB").Range("S2").Value
                .Fields("CC").Value = Sheets("4.SIRCREB").Range("T2").Value
                .Fields("ITEM").Value = Sheets("4.SIRCREB").Range("U2").Value
                .Fields("COMPROBANTE").Value = Sheets("4.SIRCREB").Range("V2").Value
                .Fields("BANCO").Value = Sheets("4.SIRCREB").Range("W2").Value
                .Fields("FACTURA").Value = Sheets("4.SIRCREB").Range("X2").Value
                .Fields("NETO").Value = Sheets("4.SIRCREB").Range("Y2").Value
                .Fields("NOTA").Value = Sheets("4.SIRCREB").Range("Z2").Value
                .Fields("TIPO COMP").Value = Sheets("4.SIRCREB").Range("AA2").Value
                .Fields("FECHA REAL").Value = Sheets("4.SIRCREB").Range("AB2").Value
                .Fields("FECHA").Value = Sheets("4.SIRCREB").Range("AC2").Value
                .Fields("FECHA_DB").Value = Sheets("4.SIRCREB").Range("AD2").Value
                .Fields("FECHA_EFECTIVA").Value = Sheets("4.SIRCREB").Range("AE2").Value
                .Fields("IVA21").Value = Sheets("4.SIRCREB").Range("AF2").Value
                .Fields("IVA27").Value = Sheets("4.SIRCREB").Range("AG2").Value
                .Fields("IVA10").Value = Sheets("4.SIRCREB").Range("AH2").Value
                .Fields("CAJA").Value = Sheets("4.SIRCREB").Range("AI2").Value
                .Fields("NC").Value = Sheets("4.SIRCREB").Range("AJ2").Value
                .Fields("VALORES").Value = Sheets("4.SIRCREB").Range("AK2").Value
                .Fields("PARALELO").Value = Sheets("4.SIRCREB").Range("AL2").Value
                
                .Fields("FORMA DE PAGO").Value = Sheets("4.SIRCREB").Range("AN2").Value
                .Fields("CUENTA").Value = Sheets("4.SIRCREB").Range("AO2").Value
                .Fields("OBSERVACIONES").Value = Sheets("4.SIRCREB").Range("AP2").Value
                .Fields("NO GRAVADO").Value = Sheets("4.SIRCREB").Range("AQ2").Value
                .Fields("MONOTRIBUTO").Value = Sheets("4.SIRCREB").Range("AR2").Value
                .Fields("PERC IB").Value = Sheets("4.SIRCREB").Range("AS2").Value
                .Fields("RET GAN").Value = Sheets("4.SIRCREB").Range("AT2").Value
                .Fields("PERC IVA").Value = Sheets("4.SIRCREB").Range("AU2").Value
                .Fields("BANCO1").Value = Sheets("4.SIRCREB").Range("AV2").Value
                .Fields("BANCO2").Value = Sheets("4.SIRCREB").Range("AW2").Value
                .Fields("REVISION").Value = Sheets("4.SIRCREB").Range("AX2").Value




        .Update
    End With
    
    ' Cerrar la conexión y el conjunto de registros
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
End Sub



Sub EscribirImpdb()
    
    ' Declarar las variables para la conexión
    Dim cn As Object
    Dim rs As Object
   
    
   
    
    ' Crear una instancia de la conexión y abrir la base de datos de Access
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\Users\admin\Desktop\contable.accdb;Persist Security Info=False;"
    

    ' Crear una instancia del conjunto de registros y especificar la tabla a escribir
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "EGRESOS", cn, adOpenDynamic, adLockOptimistic
    
     With rs
        .AddNew
                .Fields("BALANCE").Value = Sheets("10.IMP DEB CRED").Range("Q2").Value
                .Fields("EMPRESA").Value = Sheets("10.IMP DEB CRED").Range("R2").Value
                .Fields("RUBRO").Value = Sheets("10.IMP DEB CRED").Range("S2").Value
                .Fields("CC").Value = Sheets("10.IMP DEB CRED").Range("T2").Value
                .Fields("ITEM").Value = Sheets("10.IMP DEB CRED").Range("U2").Value
                .Fields("COMPROBANTE").Value = Sheets("10.IMP DEB CRED").Range("V2").Value
                .Fields("BANCO").Value = Sheets("10.IMP DEB CRED").Range("W2").Value
                .Fields("FACTURA").Value = Sheets("10.IMP DEB CRED").Range("X2").Value
                .Fields("NETO").Value = Sheets("10.IMP DEB CRED").Range("Y2").Value
                .Fields("NOTA").Value = Sheets("10.IMP DEB CRED").Range("Z2").Value
                .Fields("TIPO COMP").Value = Sheets("10.IMP DEB CRED").Range("AA2").Value
                .Fields("FECHA REAL").Value = Sheets("10.IMP DEB CRED").Range("AB2").Value
                .Fields("FECHA").Value = Sheets("10.IMP DEB CRED").Range("AC2").Value
                .Fields("FECHA_DB").Value = Sheets("10.IMP DEB CRED").Range("AD2").Value
                .Fields("FECHA_EFECTIVA").Value = Sheets("10.IMP DEB CRED").Range("AE2").Value
                .Fields("IVA21").Value = Sheets("10.IMP DEB CRED").Range("AF2").Value
                .Fields("IVA27").Value = Sheets("10.IMP DEB CRED").Range("AG2").Value
                .Fields("IVA10").Value = Sheets("10.IMP DEB CRED").Range("AH2").Value
                .Fields("CAJA").Value = Sheets("10.IMP DEB CRED").Range("AI2").Value
                .Fields("NC").Value = Sheets("10.IMP DEB CRED").Range("AJ2").Value
                .Fields("VALORES").Value = Sheets("10.IMP DEB CRED").Range("AK2").Value
                .Fields("PARALELO").Value = Sheets("10.IMP DEB CRED").Range("AL2").Value
                
                .Fields("FORMA DE PAGO").Value = Sheets("10.IMP DEB CRED").Range("AN2").Value
                .Fields("CUENTA").Value = Sheets("10.IMP DEB CRED").Range("AO2").Value
                .Fields("OBSERVACIONES").Value = Sheets("10.IMP DEB CRED").Range("AP2").Value
                .Fields("NO GRAVADO").Value = Sheets("10.IMP DEB CRED").Range("AQ2").Value
                .Fields("MONOTRIBUTO").Value = Sheets("10.IMP DEB CRED").Range("AR2").Value
                .Fields("PERC IB").Value = Sheets("10.IMP DEB CRED").Range("AS2").Value
                .Fields("RET GAN").Value = Sheets("10.IMP DEB CRED").Range("AT2").Value
                .Fields("PERC IVA").Value = Sheets("10.IMP DEB CRED").Range("AU2").Value
                .Fields("BANCO1").Value = Sheets("10.IMP DEB CRED").Range("AV2").Value
                .Fields("BANCO2").Value = Sheets("10.IMP DEB CRED").Range("AW2").Value
                .Fields("REVISION").Value = Sheets("10.IMP DEB CRED").Range("AX2").Value





        .Update
    End With
    
    ' Cerrar la conexión y el conjunto de registros
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
End Sub

Sub CargarSueldos()

' Declarar las variables para la conexión
    Dim cn As Object
    Dim rs As Object
    Dim inicio As Range
    ' Crear una instancia de la conexión
    Set cn = CreateObject("ADODB.Connection")

    ' Definir la cadena de conexión DSN-less
    Dim connectionString As String
    connectionString = "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                     "Data Source=C:\Users\admin\Desktop\Contable.accdb;" & _
                     "Persist Security Info=False;"

    ' Abrir la conexión
    cn.Open connectionString

    ' Crear una instancia del conjunto de registros y especificar la tabla a escribir
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "EGRESOS", cn, 2, 3
    
    'Recorrer la columna de datos en el archivo de Excel
    With ThisWorkbook.Sheets("Depósitos")
    
        For fila = 164 To .Cells(.Rows.Count, "U").End(xlUp).Row
        
            If Cells(fila, "U") <> "" Then
     With rs
        .AddNew
                .Fields("BALANCE").Value = Sheets("Depósitos").Range("L" & fila).Value
            .Fields("EMPRESA").Value = Sheets("Depósitos").Range("M" & fila).Value
            .Fields("RUBRO").Value = Sheets("Depósitos").Range("N" & fila).Value
            .Fields("CC").Value = Sheets("Depósitos").Range("O" & fila).Value
            .Fields("ITEM").Value = Sheets("Depósitos").Range("P" & fila).Value
            .Fields("COMPROBANTE").Value = Sheets("Depósitos").Range("Q" & fila).Value
            .Fields("BANCO").Value = Sheets("Depósitos").Range("R" & fila).Value
            .Fields("FACTURA").Value = Sheets("Depósitos").Range("S" & fila).Value
            .Fields("NETO").Value = Sheets("Depósitos").Range("T" & fila).Value
            .Fields("NOTA").Value = Sheets("Depósitos").Range("U" & fila).Value
            .Fields("TIPO COMP").Value = Sheets("Depósitos").Range("V" & fila).Value
            .Fields("FECHA REAL").Value = Sheets("Depósitos").Range("W" & fila).Value
            .Fields("FECHA").Value = Sheets("Depósitos").Range("X" & fila).Value
            .Fields("FECHA_DB").Value = Sheets("Depósitos").Range("Y" & fila).Value
            .Fields("FECHA_EFECTIVA").Value = Sheets("Depósitos").Range("Z" & fila).Value
            .Fields("IVA21").Value = Sheets("Depósitos").Range("AA" & fila).Value
            .Fields("IVA27").Value = Sheets("Depósitos").Range("AB" & fila).Value
            .Fields("IVA10").Value = Sheets("Depósitos").Range("AC" & fila).Value
            .Fields("CAJA").Value = Sheets("Depósitos").Range("AD" & fila).Value
            .Fields("NC").Value = Sheets("Depósitos").Range("AE" & fila).Value
            .Fields("VALORES").Value = Sheets("Depósitos").Range("AF" & fila).Value
            .Fields("PARALELO").Value = Sheets("Depósitos").Range("AG" & fila).Value
            .Fields("FORMA DE PAGO").Value = Sheets("Depósitos").Range("AI" & fila).Value
            .Fields("CUENTA").Value = Sheets("Depósitos").Range("AJ" & fila).Value
            .Fields("OBSERVACIONES").Value = Sheets("Depósitos").Range("AK" & fila).Value
            .Fields("NO GRAVADO").Value = Sheets("Depósitos").Range("AL" & fila).Value
            .Fields("MONOTRIBUTO").Value = Sheets("Depósitos").Range("AM" & fila).Value
            .Fields("PERC IB").Value = Sheets("Depósitos").Range("AN" & fila).Value
            .Fields("RET GAN").Value = Sheets("Depósitos").Range("AO" & fila).Value
            .Fields("PERC IVA").Value = Sheets("Depósitos").Range("AP" & fila).Value
            .Fields("BANCO1").Value = Sheets("Depósitos").Range("AQ" & fila).Value
            .Fields("BANCO2").Value = Sheets("Depósitos").Range("AR" & fila).Value
            .Fields("REVISION").Value = Sheets("Depósitos").Range("AS" & fila).Value
            
            
                

        .Update
    End With

            End If
           
        Next fila
    End With
    
    'Cerrar el recordset y la base de datos
    rs.Close
    cn.Close
    
    'Limpiar la memoria
    Set rs = Nothing
    Set cn = Nothing

End Sub

Sub Macro1_debitos()

Dim wsDebitos As Worksheet


 Set wsDebitos = Sheets("2.DEBITOS Y CHEQUES")

 Call debitosVarios
 
 Call debitoMercadoPag
 
 Call debitoDebin
 
 Call liq
 
 Call transf
 
  
 Application.Wait Now + TimeValue("00:00:08")
 
' Descombinar celdas antes de ordenar
    wsDebitos.Columns("A:G").UnMerge

' Ordenar por columna A

    wsDebitos.Sort.SortFields.Clear
    wsDebitos.Sort.SortFields.Add Key:=wsDebitos.Range("A:A"), Order:=xlAscending
    wsDebitos.Sort.SetRange wsDebitos.UsedRange
    wsDebitos.Sort.Header = xlYes
    wsDebitos.Sort.Apply
    


End Sub


Sub debitosVarios()

    Dim wsDatos As Worksheet
    Dim wsDebitos As Worksheet
    Dim lastRowDatos As Long
    Dim criteriaArray As Variant
    Dim rngFiltrada As Range

    ' Establecer las hojas de trabajo
    Set wsDatos = Sheets("DATOS")
    Set wsDebitos = Sheets("2.DEBITOS Y CHEQUES")

    ' Definir el array de criterios
    criteriaArray = Array("GUMA FISHERTON", "ARROYITO MAQUINARIAS", "COMERCIAL CAB", "INDUFER S A", "MODB PLAZO FIJO MACRONLINEÿ", "MAUGERI REPUESTOS", "MADERAS AMIANO S.A.", "ZING TECH SRL", "DB TARJETA DE CRÉDITO VISA", "ROSARIO PACK SRL", "PINTURERIA EL COLIBRI", "DB TR..AUT.SDO.MISMO TIT.", "MERPAGO*SIDEAT", "TRANSF MINORISTA MISMO TITUL", "DISTRIBUIDORA DIQUE SR", "MEYER ELECTROMECANICA", "BULONERIA INTEGRAL", "SANIFER", "ING ROSEMBERG SA", "COMPANIA AMERICANA", "ESTACION YPF", "EL COLIBRI", "COM RED", "HIDROPISCINAS", "ARGELEC", "AGUAS ASSA", "CPIC SANTA FE DISTR II", "AIMARO ELECTRICIDAD", "TOYOTA (Compañía Financiera)", _
    "BLAZQUEZ ALEJANDRO A", "CCERR FUNDACION F 30617777386 CIRC.CERRADO", "ARCO IRIS SUP SUC FISH", "HELADERIA GRIDO", "SEC TRABAJO MULTAS", "TRANSF MARISOL/C 27315406841 VAR VARIOS", "PAPELERA ARROYO SECO", "ANSV CENAT", "MP *MPOTIGUARES", "JOSELY M DE SOUZA", "AL MARO ELECTRICIDAD", "DEB PERCEP GANAN PJ", "DO CHEF GOUMERT", "COMERCIAL RAM SA", "EL VAGUA", "ROSARIO ABRASIVOS SRL", "EMPRESA PROVINCIAL ENE", "TECNICA ALBERDI", "SAFIT CENAT", "API", "TRANSF:2307070000517608263771-23341571189", "RRSS Rosario 17 _21068", "PLUS D ENERGIE", "AGENCIA PROV DE SEGVIA", "Colegio de Medicos", "LA TENAZA", "AQUILES ELECTRICIDAD E", "RESINAS ROSARIO S R L", "DEBITO PRESTAMOS REC", "ENCOMIENDAS NUEVA CHEV", "ILUMINACION Y ALGO MAS", _
    "MUNICIPALIDAD DE ROSAR", "SAKURA MOTORS", "ESTACION MENDOZA", "COTO SUCURSAL 96", "REX", "ACCESANIGA", "BM SE9ALIZACIONES SA", "PRINTEMPS", "DB TRANSF MINORISTA DIST TIT", "HSG REPUESTOS CHEVROLE", "NIC.AR", "MURNE SA", "FERRETERIA REMO FRANCO", "DRICCO", "TRIAGO CESPED Y RIEGO", "GOTTIG FULL", "COLOR N - CANADA PINT", "MORO REVESTIMIENTOS", "DIST MOTOR METAL ROSAR", "ANCO SRL", "LA CASA DEL FILTRO", "PAGO SERV DBCTA C/CONS", "GRUPO AUTOPARTISTA SA", "LASER ELECTROTECNICA", "MARSILI NEUMATICOS", "API SANTA FE PATENTES", "YPF", "TRF MO CCDO DIST T - 30711871264", "BICICLETAS MO", "WWW.PLUSPAGOS.COM", "NEUMATICOS LA FE", "PU PINTURERIA UNIVERSO", "RUSMA SRL", "DECORAL", "CASA D'RICCO", "TEF DATANET PR UNION OBRERA DE LA 30503049097", "J Y M COMUNICACIONES", "TOYOTA CFA", "DGR SELLOS SANTA FE", "CHEQUE CANJE INTERNO", "CLARO (EX. CTI)", "CHEQUE P/CAMARA", "-30711871264", "TELEDIFUSORA SA", "MUNICIPALIDAD DE ROSARIO", "Y.P.F. S.A.", "TEF DATANET BTOB", "PAYPERTIC", "EPE SANTA FE", "TELECOM ARGENTINA", _
    "DB TARJETA DE CREDITO VISA", "LADER", "ISVA SRL", "DEBITO PERCEP GCIAS", "IMPUESTO PAIS S/OP. DEBITO", "BUSPLUS", "SUPERMERCADOS IRMAOS U", "VICRISTAL SRL", "HERRERO SRL", "WWW CLARO.COM", "ORLANDI IND. Y COM. S.", "PINOMAR FISHERTON", "TECNO FRENOS SA", "GINZA SA", "BULONERA FENIX", "LADRILLOS Y REVESTIMIE", "TERSUAVE   PINTURER A", "API SANTA FE", "EST DE SERV RAUL SRL", "API SANTA FE INMO URBANO", "CHIAPERO SRL", "DEBITO TRANSFERENCIAS MEP", "SAKURA MOTORS SA", "ATENAS", "PAG.SERV.DBCTA.C/CONS", "PAG.SERV DBCTA C/CONS", "CLARO (SUB1)", "LA CASA DEL ENGANCHE", "PAGO JUDICIALES", "MERCPAGO*MERCADOLIBRE", "INSARG", "PRINTEMPS SRL", "ROBERTO CALZAVARA SRL", "MONTARFE SRL", "CLARO (SUB1)", "CHEQUE P/CAMARA", "TECNICA FISHERTON", "PINTURERIA 7 COLORES", "DB TRANSF MINORISTA DIST TIT", "SAN CRISTOBAL SEGUROS Suc.:341", "TRANSFERENCIA  INTERBANCARIA", "LE MAS", "SCTBALSEG - SAN CRISTOBAL SEGUROS", "BELGRANO GOMA", "Transf. MacrOnline E-set D/T")

    ' Desactivar filtros si ya están activos
    wsDatos.AutoFilterMode = False

    ' Encontrar la última fila en la columna "E" en la hoja "DATOS"
    lastRowDatos = wsDatos.Cells(wsDatos.Rows.Count, "E").End(xlUp).Row
    
    ' Borrar datos en las columnas A a G en la hoja "2.DEBITOS Y CHEQUES"
    wsDebitos.Range("A:G").ClearContents

    ' Aplicar el filtro en la hoja "DATOS"
    wsDatos.Range("E1:E" & lastRowDatos).AutoFilter Field:=1, Criteria1:=criteriaArray, Operator:=xlFilterValues

    ' Verificar si hay filas filtradas
    On Error Resume Next
    Set rngFiltrada = wsDatos.AutoFilter.Range.Offset(1, 0).Resize(lastRowDatos - 1, 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Copiar datos filtrados a la hoja "2.DEBITOS Y CHEQUES"
    If Not rngFiltrada Is Nothing Then
        ' Copiar solo las columnas A a G de la hoja "DATOS"
        wsDatos.Range("A1:G" & lastRowDatos).SpecialCells(xlCellTypeVisible).Copy Destination:=wsDebitos.Cells(1, 1)
    End If

    ' Limpiar filtro en la hoja "DATOS"
    wsDatos.AutoFilterMode = False
End Sub

Sub debitoMercadoPag()

    Dim wsDatos As Worksheet
    Dim wsDebitos As Worksheet
    Dim lastRowDatos As Long
    Dim rngFiltrada As Range

    ' Establecer las hojas de trabajo
    Set wsDatos = Sheets("DATOS")
    Set wsDebitos = Sheets("2.DEBITOS Y CHEQUES")

    ' Desactivar filtros si ya están activos
    wsDatos.AutoFilterMode = False

    ' Encontrar la última fila en la columna "E" en la hoja "DATOS"
    lastRowDatos = wsDatos.Cells(wsDatos.Rows.Count, "E").End(xlUp).Row

    ' Aplicar el filtro en la hoja "DATOS"
    wsDatos.Range("E1:E" & lastRowDatos).AutoFilter Field:=1, Criteria1:="MERPAGO*"

    ' Verificar si hay filas filtradas
    On Error Resume Next
    Set rngFiltrada = wsDatos.AutoFilter.Range.Offset(1, 0).Resize(lastRowDatos - 1, 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Copiar datos filtrados a la hoja "2.DEBITOS Y CHEQUES"
    If Not rngFiltrada Is Nothing Then
        ' Copiar solo las columnas A a G de la hoja "DATOS"
        wsDatos.Range("A2:G" & lastRowDatos).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsDebitos.Cells(wsDebitos.Rows.Count, "A").End(xlUp).Offset(1, 0)
    End If

    ' Limpiar filtro en la hoja "DATOS"
    wsDatos.AutoFilterMode = False
End Sub


Sub debitoDebin()

    Dim wsDatos As Worksheet
    Dim wsDebitos As Worksheet
    Dim lastRowDatos As Long
    Dim rngFiltrada As Range

    ' Establecer las hojas de trabajo
    Set wsDatos = Sheets("DATOS")
    Set wsDebitos = Sheets("2.DEBITOS Y CHEQUES")

    ' Desactivar filtros si ya están activos
    wsDatos.AutoFilterMode = False

    ' Encontrar la última fila en la columna "E" en la hoja "DATOS"
    lastRowDatos = wsDatos.Cells(wsDatos.Rows.Count, "E").End(xlUp).Row

    ' Aplicar el filtro en la hoja "DATOS"
    wsDatos.Range("E1:E" & lastRowDatos).AutoFilter Field:=1, Criteria1:="DEBIN*"

    ' Verificar si hay filas filtradas
    On Error Resume Next
    Set rngFiltrada = wsDatos.AutoFilter.Range.Offset(1, 0).Resize(lastRowDatos - 1, 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Copiar datos filtrados a la hoja "2.DEBITOS Y CHEQUES"
    If Not rngFiltrada Is Nothing Then
        ' Copiar solo las columnas A a G de la hoja "DATOS"
        wsDatos.Range("A2:G" & lastRowDatos).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsDebitos.Cells(wsDebitos.Rows.Count, "A").End(xlUp).Offset(1, 0)
    End If

    ' Limpiar filtro en la hoja "DATOS"
    wsDatos.AutoFilterMode = False
End Sub

Sub liq()

Dim wsDatos As Worksheet
    Dim wsDebitos As Worksheet
    Dim lastRowDatos As Long
    Dim rngFiltrada As Range

    ' Establecer las hojas de trabajo
    Set wsDatos = Sheets("DATOS")
    Set wsDebitos = Sheets("2.DEBITOS Y CHEQUES")

    ' Desactivar filtros si ya están activos
    wsDatos.AutoFilterMode = False

    ' Encontrar la última fila en la columna "E" en la hoja "DATOS"
    lastRowDatos = wsDatos.Cells(wsDatos.Rows.Count, "E").End(xlUp).Row

    ' Aplicar el filtro en la hoja "DATOS"
    wsDatos.Range("E1:E" & lastRowDatos).AutoFilter Field:=1, Criteria1:="Liq.Susc*"

    ' Verificar si hay filas filtradas
    On Error Resume Next
    Set rngFiltrada = wsDatos.AutoFilter.Range.Offset(1, 0).Resize(lastRowDatos - 1, 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Copiar datos filtrados a la hoja "2.DEBITOS Y CHEQUES"
    If Not rngFiltrada Is Nothing Then
        ' Copiar solo las columnas A a G de la hoja "DATOS"
        wsDatos.Range("A2:G" & lastRowDatos).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsDebitos.Cells(wsDebitos.Rows.Count, "A").End(xlUp).Offset(1, 0)
    End If

    ' Limpiar filtro en la hoja "DATOS"
    wsDatos.AutoFilterMode = False
End Sub

Sub transf()

Dim wsDatos As Worksheet
    Dim wsDebitos As Worksheet
    Dim lastRowDatos As Long
    Dim rngFiltrada As Range

    ' Establecer las hojas de trabajo
    Set wsDatos = Sheets("DATOS")
    Set wsDebitos = Sheets("2.DEBITOS Y CHEQUES")

    ' Desactivar filtros si ya están activos
    wsDatos.AutoFilterMode = False

    ' Encontrar la última fila en la columna "E" en la hoja "DATOS"
    lastRowDatos = wsDatos.Cells(wsDatos.Rows.Count, "E").End(xlUp).Row

    ' Aplicar el filtro en la hoja "DATOS"
    wsDatos.Range("E1:E" & lastRowDatos).AutoFilter Field:=1, Criteria1:="TRANSF:*"

    ' Verificar si hay filas filtradas
    On Error Resume Next
    Set rngFiltrada = wsDatos.AutoFilter.Range.Offset(1, 0).Resize(lastRowDatos - 1, 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Copiar datos filtrados a la hoja "2.DEBITOS Y CHEQUES"
    If Not rngFiltrada Is Nothing Then
        ' Copiar solo las columnas A a G de la hoja "DATOS"
        wsDatos.Range("A2:G" & lastRowDatos).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsDebitos.Cells(wsDebitos.Rows.Count, "A").End(xlUp).Offset(1, 0)
    End If

    ' Limpiar filtro en la hoja "DATOS"
    wsDatos.AutoFilterMode = False
End Sub


