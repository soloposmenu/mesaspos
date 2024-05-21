VERSION 5.00
Object = "{CCB90150-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSPOSPrinter.ocx"
Begin VB.Form LoginMesas 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Meseros"
   ClientHeight    =   6555
   ClientLeft      =   3150
   ClientTop       =   1125
   ClientWidth     =   5100
   ClipControls    =   0   'False
   Icon            =   "Sys_pos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Borrar 
      Height          =   615
      Left            =   3840
      Picture         =   "Sys_pos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Borrar Codigo"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4455
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   3040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4455
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4455
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   1100
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4455
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4455
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   3040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1100
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3675
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   33000
      Left            =   4440
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3520
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lbHora 
      BackColor       =   &H00B39665&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin OposPOSPrinter_CCOCtl.OPOSPOSPrinter ImpresoraCuentas 
      Left            =   2160
      OleObjectBlob   =   "Sys_pos.frx":0884
      Top             =   5520
   End
   Begin VB.Label lbVersion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3555
      TabIndex        =   16
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   1650
      Picture         =   "Sys_pos.frx":08A8
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   4920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "Número de Mesero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1110
      TabIndex        =   3
      Top             =   2655
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "Deslice la tarjeta por el Lector y Presione ACEPTAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   690
      TabIndex        =   1
      Top             =   1965
      Width           =   3735
   End
End
Attribute VB_Name = "LoginMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private rsCheckTmpTrans As ADODB.Recordset
'Public rsCheckTmpTrans As New ADODB.Recordset
Private nlPase As Byte
Private rsUsr As Recordset
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'INFO: NOMBRE DE LA MAQUINA. MARZO 2010
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
'PARA MEJORAR LA SALIDA DEL SISTEMA
Private bErrorGrave As Boolean

Private Function NameOfTheComputer(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function

Private Function WindowsDirectory() As String
    ' Retrieve the Windows directory.
    Dim strBuffer As String
    Dim lngLen As Long
    strBuffer = Space(dhcMaxPath)
    lngLen = dhcMaxPath
    lngLen = GetWindowsDirectory(strBuffer, lngLen)
    ' If the path is longer than dhcMaxPath, then
    ' lngLen contains the correct length. Resize the
    ' buffer and try again.
    If lngLen > dhcMaxPath Then
        strBuffer = Space(lngLen)
        lngLen = GetWindowsDirectory(strBuffer, lngLen)
    End If
    WindowsDirectory = Left$(strBuffer, lngLen)
End Function
Private Function HandleSOCIOS() As Boolean
'INFO: VERIFICA QUE LA TABLA DE SOCIO EXISTA, SI NO EXISTE EL SISTEMA LA CREA
Dim rsTempSocios As ADODB.Recordset
Dim cSQL As String

Set rsTempSocios = New ADODB.Recordset

On Error Resume Next
rsTempSocios.Open "SELECT * FROM SOCIO", msConn, adOpenStatic, adLockOptimistic
If Err.Number = -2147217865 Then
    'INFO: LA TABLA DE SOCIO NO EXISTE
    'Y ESTE SISTEMA TIENE LA OPCION DE SOCIOS. SE VA A CREAR LA TABLA
    EscribeLog ("Function: HandleSOCIOS - CREACION DE TABLA DE SOCIOS.")
    cSQL = "CREATE TABLE SOCIO "
    cSQL = cSQL & "(MESA SHORT, SOCIO SHORT, "
    cSQL = cSQL & "SOCIO_NOMBRE TEXT(40), SOCIO_SALDO SINGLE) "
    msConn.Execute cSQL
End If
Set rsTempSocios = Nothing
HandleSOCIOS = True
On Error GoTo 0
End Function
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Sub VerificaCierre()
On Error GoTo ErrAdm:

'ELIMINANDO LOG FILES
'INFO: SEP2009

''''Open App.Path & "\INOUTLOG.TXT" For Input As #1
''''Do Until EOF(1)
''''   Line Input #1, a$
''''Loop
''''Close #1
''''
''''If a$ = "OK" Then
''''    'CERRO BIEN
''''Else
''''    Open App.Path & "\VENTALOG.TXT" For Append As #1
''''        Print #1, "-- SOLO POS NO CERRO BIEN --" & Date & " " & Time
''''    Close #1
''''End If
''''
''''Open App.Path & "\INOUTLOG.TXT" For Output As #1
''''    Print #1, "NUNCA BORRAR ESTE ARCHIVO"
''''Close #1
On Error GoTo 0
Exit Sub

ErrAdm:
MsgBox Err.Number & " - " & Err.Description
Resume Next
End Sub

Private Function AbrirFile() As Boolean
'VERIFICA SI ES NECESARIO BORRAR TRANS LOCAL
Dim FecHost As Variant
Dim FecLoc As Variant
Dim RSLOC01 As New ADODB.Recordset
Dim nUpdateFlag As Integer
Dim iInt As Integer

On Error GoTo ErrorAdm:

nUpdateFlag = 0 'CERO, NO HAY QUE ACTUALIZAR
iInt = 0

''Open App.Path & "\OrigenDB.txt" For Input As #1
''Do Until EOF(1)
''    Line Input #1, a$
''    If Left(a$, 1) = "*" Then
''        DATA_PATH = Mid(a$, 3, Len(a$) - 2)
''    Else
''        cDataPath = a$
''    End If
''Loop
''Close #1

''''DATA_PATH = GetENCRYPTEDINI("General", "DirectorioDatos", App.path & "\soloini.ini")
''''cDataPath = GetENCRYPTEDINI("General", "ProveedorDatos", App.path & "\soloini.ini")
''''NOM_PRN_COCINA = GetENCRYPTEDINI("SoloPosDisp", "Cocina", App.path & "\soloini.ini")
''''NOM_PRN_COCINA02 = GetENCRYPTEDINI("SoloPosDisp", "Cocina2", App.path & "\soloini.ini")
''''NOM_PRN_COCINA03 = GetENCRYPTEDINI("SoloPosDisp", "Cocina3", App.path & "\soloini.ini")

DATA_PATH = GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini")
cDataPath = GetFromINI("General", "ProveedorDatos", App.Path & "\soloini.ini")
NOM_PRN_COCINA = GetFromINI("SoloPosDisp", "Cocina", App.Path & "\soloini.ini")
NOM_PRN_COCINA02 = GetFromINI("SoloPosDisp", "Cocina2", App.Path & "\soloini.ini")
NOM_PRN_COCINA03 = GetFromINI("SoloPosDisp", "Cocina3", App.Path & "\soloini.ini")

cKitchenFile = DATA_PATH + KITCHEN_FILE
cBarFile = DATA_PATH + BAR_FILE
cFactFile = DATA_PATH + FACTURA_FILE

If Dir(DATA_PATH) = "" Then
    'INFO: VERIFICA LA CONEXION CON LA BASE DE DATOS
    ShowMsg "ERROR GRAVE. NO SE PUEDE CONECTAR A LA BASE DE DATOS" & vbCrLf & _
                    "EL PROGRAMA TERMINARA AHORA" & vbCrLf & vbCrLf & _
                    "BASE DE DATOS: " & DATA_PATH, vbYellow, vbRed
    AbrirFile = False
    Exit Function
Else
End If

cDataPath = "Provider=Microsoft.Jet.OLEDB.4.0;" & cDataPath & ";Jet OLEDB:Database Password=master24"
' ~~~~~ cDataPath = ";Provider=MSDASQL;DRIVER={Microsoft Access Driver (*.mdb)};"
' ~~~~~ cDataPath = cDataPath & ";DBQ=" & DATA_PATH & "SOLO.MDB"
' ~~~~~ cDataPath = cDataPath & ";UID=admin;PWD=master24;"

'INFO: 21ABR2016 CAMBIANDO AL APP.PATH YA QUE LAS NUEVAS VERSIONES DE WINDOWS
' NO PERMITEN ESCRITURA EN ESTE FOLDER.
'ADMIN_LOG = WindowsDirectory
ADMIN_LOG = App.Path
ADMIN_LOG = ADMIN_LOG & "\ADMLOG.SOL"

'If GetENCRYPTEDINI("Facturacion", "BuscaSocio", App.path & "\soloini.ini") = "Pereza" Then
If UCase(GetFromINI("Facturacion", "BuscaSocio", App.Path & "\soloini.ini")) = "PEREZA" Then
    bISThisSocios = True
    Call HandleSOCIOS
Else
    bISThisSocios = False
End If
On Error GoTo 0
ON_LINE = True
If ON_LINE = True Then
    
    '\\SOLO11\ACCESS\SOLO.mdb;"

    'INFO: NOV/2008 ABRIENDO DOMICILIO
    On Error Resume Next
    nDomicilio = GetFromINI("Facturacion", "Domicilio", App.Path & "\soloini.ini")
    If Err.Number = 13 Then
        nDomicilio = 0  'MARCANDO QUE NO SI ES UN SOLOINI VIEJO
    End If
    On Error GoTo 0
    On Error GoTo ErrDBMSOpen:
    
    If nDomicilio <> 0 Then
        LoginMesas.Caption = LoginMesas.Caption + ".DOMICILIO"
        HAS_Domicilio = True
        Call OpenDBDomicilio
    Else
        LoginMesas.Caption = LoginMesas.Caption + ".ON LINE"
    End If

    'INFO: 28MAY2011
    ' ~~~~~ msConn.CursorLocation = adUseClient
    msConn.Open cDataPath
    'msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=\\SOLO11\ACCESS\SOLO.mdb;" _
            + "Jet OLEDB:Database Password=master24"
    
    'INFO: 28/10/2006
    Set msPED = New ADODB.Connection
    'INFO: AGREGANDO EL MODO DE APERTURA
    msPED.Mode = adModeShareDenyNone
    
    msPED.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DATA_PATH & "MESASPED.MDB;"
    ' ~~~~~ msPED.Open "Provider=MSDASQL.1;Driver={Microsoft Access Driver (*.mdb)};Extended Properties=DBQ=" & DATA_PATH & "MESASPED.MDB;DriverId=25;FIL=MSAccess;MaxBufferSize=2048;PageTimeout=5;UID=admin;PWD=;"
    AbrirFile = True
    
    On Error GoTo 0
Else

    LoginMesas.Caption = LoginMesas.Caption + ".OFF LINE"
    On Error GoTo ErrDBMSOpen:

    ' ~~~~~ msConn.CursorLocation = adUseClient
    msConn.Open cDataPath
    
    'msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=C:\LoginMesas\LOCAL\SOLO.mdb;" _
            + "Persist Security Info=False"
    AbrirFile = True
    On Error GoTo 0
    'msConn.Open "Provider=Microsoft.Jet.OLEDB.3.51" _
        + ";Persist Security Info=False;Data Source=" _
        + app.path & "\LOCAL\SOLO.mdb"
    MsgBox "TRABAJANDO OFF-LINE (Fuera de Linea). Puede Continuar. Presione Enter", vbInformation, BoxTit
End If

RSLOC01.Open "SELECT * FROM OPCIONES", msConn, adOpenStatic, adLockOptimistic

If RSLOC01!CHECK_UP = "Null" Then
    TipoApplicacion = " ESTE PRODUCTO ES UNA DEMOSTRACION."
Else
    TipoApplicacion = ""
End If

Me.Caption = TipoApplicacion + Me.Caption
SLIP_OK = False

If Not RSLOC01.EOF Then
    SLIP_OK = RSLOC01!SLIP_PRINTER
    REPCAJAX_OK = RSLOC01!REPORTX_OK
    MAX_DESCUENTO = RSLOC01!MAX_DESC
    OPEN_PROPINA = RSLOC01!PANTA_PROP
    If IsNull(RSLOC01!PROP_DESCR) Then
        PROPINA_DESCRIP = ""
    Else
        PROPINA_DESCRIP = Trim(RSLOC01!PROP_DESCR)
    End If
    HABITACION_OK = RSLOC01!HABITACION
    RSLOC01.Close
End If
Exit Function

ErrorAdm:
ON_LINE = False
Resume Next

'''ErrorCopiaON:
'''    ' La BD no se pudo copiar alguien lo esta usando en la oficina
'''    MsgBox "ON LINE ¡ ERROR AL COPIAR BASES DE DATOS ! POSIBLEMENTE " & _
'''           "LA ESTEN USANDO EN LA OFICINA. EL PROGRAMA TERMINARA AHORA.", vbCritical, BoxTit
'''    Unload Me
'''    End

ErrDBMSOpen:
'Error grave NO SE ABRE DBMS
''Dim OBJERR As Error
''MsgBox "ERROR GRAVE. - " & Err.Description, vbCritical, "Error abriendo archivo"
''EscribeLog Err.Description
''For Each OBJERR In msConn.Errors
''     MsgBox OBJERR.Number & " <-> " & OBJERR.Description & vbCrLf & _
''            DATA_PATH & "MESASPED.MDB;", vbCritical, "Error Grave. ANOTE EL NUMERO"
''Next

Dim OBJERR As Error
ShowMsg "ERROR GRAVE. - (" & Err.Number & ") " & Err.Description & vbCrLf & "Error al Iniciar programa." & vbCrLf & "Funcion.LoginMesas.AbrirFile ", vbYellow, vbRed
'Resume
For Each OBJERR In msConn.Errors
    ShowMsg OBJERR.NativeError & " (" & OBJERR.Number & ") <-> " & OBJERR.Description & vbCrLf & vbCrLf & "Error Grave. SOLO.MDB", vbYellow, vbRed
Next

For Each OBJERR In msPED.Errors
    ShowMsg OBJERR.NativeError & " (" & OBJERR.Number & ") <-> " & OBJERR.Description & vbCrLf & vbCrLf & "Error Grave. MESASPED.MDB", vbYellow, vbRed
Next

'''Unload Me
'''End
End Function

Private Sub Borrar_Click()
Dim i As Integer

Text1.Text = ""
Text1.Refresh
Text1.SetFocus
'For i = 0 To Command8.Count - 1
'    Command8(i).BackColor = &H8000000F
'Next
End Sub

Private Sub Command1_Click()
Dim cSuperMesero As String
Dim cLoginMesero As String

If Len(LoginMesas.Text1) < 3 Then Exit Sub
'If Not IsNumeric(Text1) Then Exit Sub

On Error GoTo ErrorSuperMesero:

LoginMesas.Text1 = Trim(LoginMesas.Text1.Text)
If Len(LoginMesas.Text1) < 3 Then Exit Sub
'////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
    'INFO: MODIFICACION PARA SUPERMESEROS
    'SEP2009
    'UPDATE 17NOV2009
    '1- LEER REGISTRO PARA VER SI ES SUPER MESERO
     'cSuperMesero = RegRead("HKLM\Software\SoloSoftware\Meseros\SuperMeseros")
     cSuperMesero = RegRead("HKCU\Software\SoloSoftware\Meseros\SuperMeseros")
     If cSuperMesero = "" Then
        'DO NOTHING, CONTINUE
     Else
        If Len(LoginMesas.Text1.Text) < 4 Then
            'DO NOTHING, IS NOT SUPERMESEROS
            cSuperMesero = ""
        Else
            If InStr(1, cSuperMesero, LoginMesas.Text1.Text) > 0 Then
                'OK EL LOGIN ES DE UN SUPERMESERO. LLAMAR AL PROGRAMA
                
                
                'INFO: DESACTIVADO 22MAY2015. Timer1.Enabled = False
                'INFO: ACTIBANDO PARA RELOJ. 7ENE2019
                Timer1.Enabled = False

                ImpresoraCuentas.ReleaseDevice
                ImpresoraCuentas.Close

                Me.MousePointer = vbHourglass
                cLoginMesero = LoginMesas.Text1.Text
                LoginMesas.Text1.Text = ""
                
                'PASA COMO PARAMETRO EL DATO LEIDO EN LoginMesas.Text1.Text
                Call Shell(App.Path & "\SuperMesero.exe" & Space(1) & cLoginMesero, vbNormalFocus)
    ''
    ''            LoginMesas.Text1.SetFocus
                Me.MousePointer = vbDefault
                Exit Sub            '/////  HAY QUE SALIR DEL CLICK PARA QUE NO CONTINUE LA OPERACION
                'Y DEJAR ESTE ABIERTO
                'DESACTIVAR IMPRESORA DE PRECUENTAS
            End If
        End If
    End If
'////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////

On Error GoTo 0

On Error GoTo ErrorADO:
Set rs = New Recordset
'Set rsUsr = New Recordset

rs.Open "SELECT numero,nombre,apellido FROM meseros WHERE CLAVE = '1' AND track_reader = '" & LoginMesas.Text1.Text & "'", msConn, adOpenForwardOnly, adLockReadOnly

'rsUsr.Open "SELECT numero,nombre FROM USUARIOS WHERE numero = " & _
    Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

Set rs00 = New Recordset
rs00.Open "SELECT * FROM ORGANIZACION ", msConn, adOpenDynamic, adLockReadOnly

DoEvents
'Label2(2) = App.CompanyName
Label2(2) = rs00!descrip: Label2(2).Refresh

If rs.EOF Then
    ShowMsg "Informacion es INCORRECTA o EL MESERO NO ESTA ACTIVO, Intente de Nuevo", vbRed, vbYellow
    LoginMesas.Text1.Text = ""
    LoginMesas.Text1.SetFocus
    Exit Sub
End If

npNumCaj = rs!numero
cNomCaj = rs!nombre
cNomMesero = rs!nombre

'INFO: ACTUALIZACION DE VARIABLES (JUNIO 2010)
nMesero = rs!numero

nDesc01 = rs00!desc_01
nDesc02 = rs00!desc_02
nMesaBarra = rs00!MESA_BARRA
If bAutoLogin Then
Else
   LoginMesas.Text1.Text = ""
End If
On Error GoTo 0


'INICIAR EL TIMER
'INFO: DESACTIVADO 22MAY2015. Timer1.Enabled = True

On Error Resume Next
LoginMesas.Text1.SetFocus
On Error GoTo 0

'' --- > LoginMesas.Visible = False


'INFO: 18JUN2017
'Call GetFechaHoraFromServer

'INFO: ACTIBANDO PARA RELOJ. 7ENE2019
Timer1.Enabled = False

If bAutoLogin Then
    If IsLogged Then
        PLU.Show
    End If
Else
    If VerificaAreas Then
        SelectArea.Show 1
    End If
    PLU.Show
End If
Exit Sub

ErrorSuperMesero:
    If Err.Number = 53 Then
        ShowMsg "PROGRAMA DE SUPER.MESEROS NO HA SIDO ENCONTRADO"
    Else
        ShowMsg "SUPER.MESEROS" & vbCrLf & Err.Number & " - " & Err.Description
    End If
    Me.MousePointer = vbDefault
    Exit Sub

ErrorADO:
  Dim ADOError As Error
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
  Next ADOError
  MsgBox sError, vbCritical, "Error Grave. ANOTE EL NUMERO"
  Resume Next

End Sub
Private Function VerificaAreas() As Boolean
Dim cSQL As String, cSQL2 As String
Dim rsAreas As ADODB.Recordset

   On Error GoTo VerificaAreas_Error

cSQL = "SELECT * FROM AREAS ORDER BY DESCRIPCION"

Set rsAreas = New ADODB.Recordset
rsAreas.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If rsAreas.EOF Then
    isArea = False
    VerificaAreas = False
Else
    isArea = True
    VerificaAreas = True
End If
rsAreas.Close
Set rsAreas = Nothing

   On Error GoTo 0
   Exit Function

VerificaAreas_Error:

    ShowMsg "DEBE ACTUALIZAR ADMINISTRACION A LA VERSION" & vbCrLf & "(12.3.1+)" & vbCrLf & "ANTES DE USAR ESTA APLICACION" & vbCrLf & _
            "Error " & Err.Number, vbYellow, vbRed
End Function
Private Sub Command2_Click()
Dim hwnd As Integer
Dim Mifrm As Form
Dim ccError As String
Dim ccBase As String

On Error GoTo 0

Label2(2).ForeColor = &HFF&
'Label2(2) = "EL SISTEMA SE ESTA CERRANDO... ESPERE UNOS SEGUNDOS"
Label2(2) = "EL SISTEMA SE ESTA CERRANDO... ESPERE": Label2(2).Refresh


'INFO: DESACTIVADO 22MAY2015. Me.Timer1.Enabled = False

On Error GoTo ErrAm:

nVeriSalida = 2
'INFO: ACTIVANDO TIMER. 7ENE2019
Timer1.Enabled = False

'INFO: ELIMINANDO LOGFILES
'SEP2009

''Open App.Path & "\INOUTLOG.TXT" For Output As #1
''    Print #1, "OK"
''Close #1

'****** CIERRE DE OBJETOS OLE POS ******
'***************************************
If NOM_PRN_FACTURA = "" Then
    'INFO: NO HACER NADA
    'Set ImpresoraCuentas = Nothing
Else
    'Debug.Print ImpresoraCuentas.State
    rc = ImpresoraCuentas.State
    If rc = OposSBusy Then
        ShowMsg "LA IMPRESORA ESTA OCUPADA" & vbCrLf & _
                "INTENTE OTRA VEZ CUANDO TERMINE DE IMPRIMIR o INTENTE SALIR OTRA VEZ." & vbCrLf & " Error # " & rc, vbRed, vbYellow
        'INFO: DESACTIVADO 22MAY2015. Me.Timer1.Enabled = True
        On Error GoTo 0
        Exit Sub
    End If
    'INFO: 8ABR2018. AGREGA .DeviceEnabled = False AL PROCESO.
    ImpresoraCuentas.DeviceEnabled = False
    Label2(2) = "ImpresoraCuentas.ReleaseLocalPrinter": Label2(2).Refresh
    'INFO: 28ABRIL2014
    ImpresoraCuentas.ReleaseDevice
    
    Label2(2) = "ImpresoraCuentas.Close": Label2(2).Refresh
    ImpresoraCuentas.Close

    
    ''-- Gaveta de Dinero Cocash1.Release
    ''-- Gaveta de Dinero Cocash1.Close
End If
'***************************************
'***************************************
'

EscribeLog ("Salida de MESEROS." & cMachineName)

For Each Mifrm In Forms
       Mifrm.Hide          ' hide the form
       Unload Mifrm        ' deactivate the form
       Set Mifrm = Nothing   ' remove from memory
Next

On Error Resume Next
'StatMesa nMesa, 0
StatMesa nMesa, vbLibre, "LoginMesas"
On Error GoTo 0

On Error GoTo ErrAm:

Set rs = Nothing
ccBase = " (DB MESASPED) ": msPED.Close
ccBase = "( DB SOLO) ": msConn.Close
On Error GoTo 0

On Error Resume Next
If HAS_Domicilio Then
    Call CloseDBDomicilio
End If
On Error GoTo 0

Unload Me
End
On Error GoTo 0
Exit Sub

ErrAm:
If Not bErrorGrave Then
    ccError = Err.Number & ccBase & " - " & Err.Description
    EscribeLog "Meseros.Salida.Error: " & ccError
    ShowMsg "Meseros.Salida.Error: " & ccError, vbBlue, vbYellow
    On Error Resume Next
    Set rs = Nothing
    If msConn.State = adStateOpen Then
        msConn.Close
    End If
    '''If msPED.State = adStateOpen Then
    '''    msPED.Close
    '''End If
Else
    On Error Resume Next
End If
Unload Me
End
End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String
Dim i As Integer

On Error Resume Next
If nlPase = 0 Then
    Text1 = Command8(Index).Index
Else
    cCant = Str(Text1)
    cCant = cCant & Command8(Index).Index
    Text1 = cCant
End If
nlPase = nlPase + 1
'Command8(Index).BackColor = &HFFFF&
''Command8(Index).Refresh
'For i = 0 To Command8.Count
'    If i = Index Then
'    Else
'        Command8(i).BackColor = &H8000000F
'    End If
'Next
On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim lReturnn As Long

On Error GoTo ErrAdm:
'Verifica si App esta abierta, para solamente cargarla una vez
If App.PrevInstance Then ActivatePrevInstance

'Call VerificaFecha
'OBTIENE LOS DATOS DE LOS NOMBRE DE LOS DISPOST. POS y OTRAS ESPECIFICACIONES
'''NOM_PRN_FACTURA = GetENCRYPTEDINI("SoloPosDisp", "Facturacion", App.path & "\soloini.ini")
'''NOM_PRN_COCINA = GetENCRYPTEDINI("SoloPosDisp", "Cocina", App.path & "\soloini.ini")
'''NOM_GAV_DINERO = GetENCRYPTEDINI("SoloPosDisp", "Gaveta", App.path & "\soloini.ini")
'''MESEROS_TIENEN_BANCO = GetENCRYPTEDINI("Meseros", "Banco", App.path & "\soloini.ini")
'''cHAYClientes = UCase(GetENCRYPTEDINI("Meseros", "HayClientes", App.path & "\soloini.ini"))

'INFO: MARZO 2010
lReturnn = NameOfTheComputer(cMachineName)

cMachineName = LTrim(RTrim(RemoveNULL(cMachineName)))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 5ENE2022
FontWindowsPrinter = GetFromINI("SoloPosDisp", "FontWindowsPrinter", App.Path & "\soloini.ini")
If FontWindowsPrinter = "" Then
    nFontWindowsPrinter = 10
Else
    nFontWindowsPrinter = Int(FontWindowsPrinter)
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NOM_PRN_FACTURA = GetFromINI("SoloPosDisp", "Facturacion", App.Path & "\soloini.ini")
'NOM_PRN_COCINA = GetFromINI("SoloPosDisp", "Cocina", App.path & "\soloini.ini")
NOM_GAV_DINERO = GetFromINI("SoloPosDisp", "Gaveta", App.Path & "\soloini.ini")
MESEROS_TIENEN_BANCO = GetFromINI("Meseros", "Banco", App.Path & "\soloini.ini")
cHAYClientes = UCase(GetFromINI("Meseros", "HayClientes", App.Path & "\soloini.ini"))
PedidoBarEnLocalPrinter = UCase(GetFromINI("Meseros", "PedidoBarEnLocalPrinter", App.Path & "\soloini.ini"))

'abrir coneccion
BoxTit = "MENSAJE DEL SISTEMA DE VENTAS"
'Show

DoEvents
Label2(2) = "Verificando Impresora/Gaveta ...": Label2(2).Refresh

DoEvents
Set msConn = New Connection

VerificaCierre
ON_LINE = True
nVeriSalida = 1

msConn.Mode = adModeShareDenyNone

DoEvents
lbVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision

On Error Resume Next
If Not AbrirFile Then  'Verifica Conección con la base de datos
    bErrorGrave = True
    Command2_Click
End If
On Error GoTo 0

DoEvents
Label2(2) = App.CompanyName

' MEJORAR VELOCIDAD. 14ENE2022
Me.Show

On Error GoTo ErrAdm:

Call GetISC 'ITBMS

'GUARDAR LA LLAVE DE SUPER MESEROS
'INFO: SEPT2009
'INFO: 19NOV2009, NO SE USARA, VER DOCUMENTACION PARA USO CORRECTO
'RegWrite "HKLM\Software\SoloSoftware\Meseros\SuperMeseros", " 021967240919380905199404"

If UCase(GetFromINI("Facturacion", "AutoLogin", App.Path & "\soloini.ini")) = "PEREZA" Then
    bAutoLogin = True
    Call AutoLogin
End If

If bAutoLogin Then
    Text1.Enabled = False
End If

'''''==================================
'''''========OLE POS FIN===============
'''''==================================
'''''==================================
'''''======== OLE POS INICIO =============
'''''==================================
'POR DEFAULT. NO HAY TABLETA (28ABRIL2014.)
cHayTableta = "NO"
If OpenLocalPrinter Then
    OPOS_DevName = LoginMesas.ImpresoraCuentas.DeviceName
    'INFO: 28ABRIL2014.
    ' SI HAY IMPRESORA y TIENE LA OPCION DE TABLETAS ENTONCES LA VA A COMPARTIR
    ' LO QUE SIGNIFICA QUE LA IMPRESORA SE ABRE DE MODO COMPARTIDO.
    ' DE NO HABER TABLETA, LA ABRE DE MODO NORMAL. AL INICIO DE LA APLICACION
    If UCase(GetFromINI("Facturacion", "Tablet", App.Path & "\soloini.ini")) = "PEREZA" Then
        cHayTableta = "SI"
        'Call Claim_Enable_LocalPrinter
    Else
        cHayTableta = "NO"
        Call Claim_Enable_LocalPrinter
    End If
    
    If LoginMesas.ImpresoraCuentas.RecEmpty = True Or LoginMesas.ImpresoraCuentas.RecNearEnd = True Then
        ShowMsg "Advertencia de Papel" & vbCrLf & "POR FAVOR REVISE EL PAPEL EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO", vbBlue, vbYellow
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: 23AGO2019. EVALUA LA IMPRESORA QUE SE ESTA USANDO
    Select Case OPOS_DevName
        '14SEP2019. MUNBY
        Case "TM-U950P", "TM-U950", "POSPrinter80", "TM-U220B", "SEMOPOS.SO.SERIAL.POSPrinter", _
                "TM-U200D", "SRP270", "SRP270P", "MP4200TH"
            nl_Descrip = 15
            nl_Line = 30
        Case "SRP-E300", "LR2000", "TM-T20-42CU", "TM-T20-42CE", "TM-T20II-42CE", "TM-T20II-42C", _
                "TM-T20E", "TM-T20U", "TM-T20III-42C"
            'INFO: 10OCT2019. VALIDA QUE FUNCIONE CON LA BIXOLON E300
            'INFO: 22NOV2019. VALIDA QUE FUNCIONE CON LA BEMATECH
            'INFO:   3MAR2020. EPSON TERMICA
            nl_Descrip = 25
            nl_Line = 40
        Case Else
            'INFO EL DEFAULT SI ES TERMICA PASA A SER 25 Y 40, YA NO 30,45
            'INFO: 14MAR2021
            nl_Descrip = 25
            nl_Line = 40
            'nl_Descrip = 30
            'nl_Line = 45
    End Select
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If
'''''==================================
'''''========OLE POS FIN===============
'''''==================================
'''''==================================
'''''========OLE POS FIN===============
'''''==================================

If UCase(GetFromINI("General", "Depura", App.Path & "\soloini.ini")) = "PEREZA" Then
    Depura.Show
End If

'INFO: 21ABRIL2014
'RESTRINGE LA PANTALLA DE ANULACION PARA QUE UNICAMENTE SE PUEDA ANULAR EN LA CAJA
If UCase(GetFromINI("Facturacion", "PermiteAnular", App.Path & "\soloini.ini")) = "NO" Then
    bPermiteAnular = False
Else
    bPermiteAnular = True
End If

'INFO: 2DIC2015. PARA LOS CLIENTES QUE SOLICITAN LA IMPRESION DEL MESERO EN LA PRECUENTA.
If UCase(GetFromINI("Facturacion", "MeseroEnPrecuenta", App.Path & "\soloini.ini")) = "PEREZA" Then
    bMeseroEnPrecuenta = True
Else
    bMeseroEnPrecuenta = False
End If
'INFO: ACTIBANDO PARA RELOJ. 7ENE2019
lbHora.Caption = Format(Time(), "hh:mm AM/PM")

' MEJORAR VELOCIDAD. 14ENE2022
Timer1.Enabled = True
Text1.SetFocus

Call LoadFullMenu

'''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 5OCT2023
cAllowSeparar = GetFromINI("Facturacion", "AllowSeparar", App.Path & "\soloini.ini")
'''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

EscribeLog ("Inicio de MESEROS." & cMachineName & " (" & App.Major & "." & App.Minor & "." & App.Revision & ")")

On Error GoTo 0
Exit Sub

ErrAdm:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "ERROR EN INICIO DEL SISTEMA"
    Resume Next
End Sub
'-------------------------------------------------------------------------------------------------
' Procedure : LoadFullMenu
' Author    : hsequeira
' Date      : 10/06/2023
' Purpose   : PREPARA TODOS LOS ITEMS DEL MENU PARA SU ACCESO POSTERIOR
'-------------------------------------------------------------------------------------------------
'
Private Sub LoadFullMenu()
Dim cSQL As String

Set rsFullMenu = New ADODB.Recordset

'        cSQL = "SELECT A.CODIGO AS DEPTO, A.DESCRIP AS DEPTO_DESCRIP,"
'        cSQL = cSQL & " B.CODIGO, B.DESCRIP AS PRODUCTO, B.PRECIO1, "
'        cSQL = cSQL & " C.CONTENEDOR, D.DESCRIP, C.PRECIO "
'        cSQL = cSQL & " FROM ((DEPTO AS A LEFT JOIN PLU AS B ON A.CODIGO = B.DEPTO)"
'        cSQL = cSQL & " LEFT JOIN CONTEND_02 AS C ON B.CODIGO = C.CODIGO)"
'        cSQL = cSQL & " LEFT JOIN CONTENED AS D ON C.CONTENEDOR = D.CONTENEDOR"

        cSQL = "SELECT A.CODIGO AS DEPTO, A.DESCRIP AS DEPTO_DESCRIP, "
        cSQL = cSQL & " B.CODIGO, B.DESCRIP AS PRODUCTO, B.PRECIO1, C.CONTENEDOR, "
        cSQL = cSQL & " D.DESCRIP, C.PRECIO, B.IMPRESORA, B.CON_TAX, B.DISPONIBLE "
        cSQL = cSQL & " FROM ((DEPTO AS A LEFT JOIN PLU AS B ON A.CODIGO = B.DEPTO)"
        cSQL = cSQL & " LEFT JOIN CONTEND_02 AS C ON B.CODIGO = C.CODIGO)"
        cSQL = cSQL & " LEFT JOIN CONTENED AS D ON C.CONTENEDOR = D.CONTENEDOR "

rsFullMenu.Open cSQL, msConn, adOpenStatic, adLockOptimistic


End Sub




Private Function AutoLogin() As Boolean
LoginMesas.Text1.Text = "500"
Command1_Click
If Not rs.EOF Then
    PLU.Show
    IsLogged = True
Else
    Command2_Click
End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If nVeriSalida = 1 Then
    MsgBox "¡ Favor utilize el boton Salir !", vbExclamation, BoxTit
    Cancel = True
End If
End Sub

Private Sub Image1_DblClick()
'MsgBox "Empresa : " & App.CompanyName & Chr(13) & _
       "Derechos Reservados : " & App.LegalCopyright & Chr(13) & _
       "Nombre  : " & App.EXEName & Chr(13) & _
       "Versión : " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Informacion de la Aplicación"
       
ShowMsg App.LegalCopyright & Chr(13) & _
       "Nombre: " & App.EXEName & ". Versión: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & _
       "Edición Clásica y para Tablets", vbGreen, vbBlue
End Sub

Private Sub OPOSPOSPrinter_DirectIOEvent(ByVal EventNumber As Long, pData As Long, pString As String)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        Command1_Click
    Case Else
End Select
End Sub

Private Sub Timer1_Timer()
lbHora.Caption = Format(Time(), "hh:mm AM/PM")
'INFO: DESACTIVADO 22MAY2015. Dim retVal
'INFO: DESACTIVADO 22MAY2015. Dim iFile As Integer

'INFO: DESACTIVADO 22MAY2015. On Error GoTo ErrAdm:
'INFO: ES PARA CAPTURAR EL PROBLEMA DE MESERO-CAJERO,
'PERO EL PROBLEMA APARENTEMENTE ES EN LA CAJA, NO AQUI
'''''Set rsCheckTmpTrans = New ADODB.Recordset
'''''rsCheckTmpTrans.Open "SELECT CAJERO,MESA,MESERO,LIN,DESCRIP " & _
'''''            " FROM TMP_TRANS " & _
'''''            " WHERE CAJERO <> MESERO ", msConn, adOpenStatic, adLockOptimistic
'''''If rsCheckTmpTrans.RecordCount > 0 Then
'''''    MsgBox rsCheckTmpTrans!CAJERO & " - " & rsCheckTmpTrans!mesero & vbCrLf & _
'''''            "Mesa : " & rsCheckTmpTrans!mesa & vbCrLf & _
'''''            "Login de mesero : " & nMesero & vbCrLf & _
'''''            "Producto : " & rsCheckTmpTrans!LIN & ")" & rsCheckTmpTrans!descrip & vbCrLf & _
'''''            "Llamar a SOLO SOFTWARE y Decirle TODO lo que aparece Aqui", vbCritical, "Llamar a SOLO SOFTWARE (CAJERO-MESERO)"
'''''    MsgBox rsCheckTmpTrans!CAJERO & " - " & rsCheckTmpTrans!mesero & vbCrLf & _
'''''            "Mesa : " & rsCheckTmpTrans!mesa & vbCrLf & _
'''''            "Login de mesero : " & nMesero & vbCrLf & _
'''''            "Producto : " & rsCheckTmpTrans!LIN & ")" & rsCheckTmpTrans!descrip & vbCrLf & _
'''''            "Llamar a SOLO SOFTWARE y Decirle TODO lo que aparece Aqui", vbCritical, "Llamar a SOLO SOFTWARE (CAJERO-MESERO)"
'''''End If
'''''rsCheckTmpTrans.Close
'''''Set rsCheckTmpTrans = Nothing

'INFO: MESASPED MANTENIMIENTO
'REMOVER DE AQUI (14SEP2009)
'''''''If Dir(DATA_PATH & "MANTENIMIENTO.TXT") = "" Then
'''''''    'DO NOTHING
'''''''Else
'''''''    Mesas.Timer1.Enabled = False
'''''''    Me.Timer1.Enabled = False
'''''''    msPED.Close
'''''''    ShowMsg "LA BASE DE DATOS ESTA EN MANTENIMIENTO, POR FAVOR ESPERE" & vbCrLf & _
'''''''        "EL SISTEMA CONTINUARA EN UNOS MOMENTOS." & vbCrLf & _
'''''''        "PRESIONE EL BOTON DE ACEPTAR"
'''''''    Me.MousePointer = vbHourglass
'''''''    Sleep 10000
'''''''    msPED.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DATA_PATH & "MESASPED.MDB;"
'''''''    Me.MousePointer = vbDefault
'''''''    Timer1.Enabled = True
'''''''    Mesas.Timer1.Enabled = True
'''''''End If

'INFO: DESACTIVADO 22MAY2015. If IsFormLoaded(PLU) = True Then
'INFO: DESACTIVADO 22MAY2015.     If PLU.Hora.Caption <> CStr(Time) Then
'INFO: DESACTIVADO 22MAY2015.         ' It's now a different second than the one displayed.
'INFO: DESACTIVADO 22MAY2015.         PLU.Hora.Caption = Time
'INFO: DESACTIVADO 22MAY2015.     End If
'INFO: DESACTIVADO 22MAY2015. End If
'INFO: DESACTIVADO 22MAY2015. On Error GoTo 0
'INFO: DESACTIVADO 22MAY2015. Exit Sub

'INFO: DESACTIVADO 22MAY2015. ErrAdm:
'INFO: DESACTIVADO 22MAY2015. MsgBox Err.Number & "-" & Err.Description, vbCritical, "Error en timer"
'INFO: DESACTIVADO 22MAY2015. Resume Next
'If ON_LINE = True Then
'    iFile = FreeFile
'    Open DATA_PATH + "ACCESS\SOLOLINE.TXT" For Input As iFile
'    Do Until EOF(1)
'        Line Input #iFile, a$
'    Loop
'    Close #iFile

'    If a$ = "OFF_LINE" Then
        ''''''''''''ES NECESARIO SALIR DEL PROGRAMA UN MOMENTO
        ''''''''''''YA QUE HASTA AHORA HABIAMOS TRABAJADO ON_LINE
        ''''''''''''EL PROGRAMA MSGUSER BORRA DB-LOCAL
        ''''''''''''BorraLocal
'        nVeriSalida = 2
'        RetVal = Shell(App.Path & "\MsgUser.exe", vbNormalFocus)
'        Unload Me
'        End
'    End If
'Else
    'NO SE PUEDE VERIFICAR DE ESTA MANERA, PONE AL SISTEMA MUY LENTO
'End If

End Sub

Private Sub VerificaFecha()
Dim cMaxFecha As Date
Dim cMaxDia As String
Dim cMaxMes As String
Dim cMaxYear As String
Dim cLocalFecha As String

cMaxFecha = Date
cMaxMes = Mid(Format(cMaxFecha, "short date"), 4, 2)
cMaxDia = Mid(Format(cMaxFecha, "short date"), 1, 2)
cMaxYear = Mid(Format(cMaxFecha, "short date"), 7, 4)

cLocalFecha = cMaxYear & cMaxMes & cMaxDia
If Val(cLocalFecha) > Val("20010430") Then
    MsgBox "***** SU PERIODO DE EVALUACION A TERMINADO *****" & vbCrLf & _
            "- GRACIAS POR PROBAR PRODUCTOS DE SOLO SOFTWARE DEVELOPMENT" & vbCrLf & _
            "- CONTACTE A SU PROVEEDOR" & vbCrLf & _
            "" & vbCrLf & "El programa terminara AHORA", vbCritical, "CONTACTE A SU PROVEEDOR"
    Unload Me
    End
End If
End Sub


Private Function RemoveNULL(cTexto As String) As String
'INFO: 1FEB2016. COPIA DE FUNCTION QUE ESTA EN MODULO00.
Dim i As Integer

For i = 1 To Len(cTexto)
    Select Case Mid(cTexto, i, 1)
        Case Chr(0)
           Mid(cTexto, i, 1) = "-"
        Case Chr(10) + Chr(13)
           Mid(cTexto, i, 1) = "-"
        Case Chr(13) + Chr(10)
           Mid(cTexto, i, 1) = "-"
        Case Chr(10)
            Mid(cTexto, i, 1) = "-"
        Case Chr(13)
            Mid(cTexto, i, 1) = "-"
    End Select
Next
RemoveNULL = cTexto
End Function

