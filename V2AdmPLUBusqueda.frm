VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form AdmPLUBusqueda 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSQUEDA DE PRODUCTOS"
   ClientHeight    =   7875
   ClientLeft      =   17865
   ClientTop       =   3315
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "AGREGAR AL PEDIDO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Borrar 
      Height          =   495
      Left            =   11040
      Picture         =   "V2AdmPLUBusqueda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   10960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   9440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   10960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   9440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   10960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "ACOMPAÑANTES"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2620
      Left            =   9360
      TabIndex        =   6
      Top             =   840
      Width           =   2520
      Begin VB.CommandButton cmdAcomp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdRestoAco 
      DisabledPicture =   "V2AdmPLUBusqueda.frx":0442
      Enabled         =   0   'False
      Height          =   495
      Left            =   9360
      Picture         =   "V2AdmPLUBusqueda.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1170
   End
   Begin VB.CommandButton cmdUpAcompa 
      DisabledPicture =   "V2AdmPLUBusqueda.frx":0CC6
      Enabled         =   0   'False
      Height          =   495
      Left            =   10680
      Picture         =   "V2AdmPLUBusqueda.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1170
   End
   Begin VB.CommandButton cmdRegresar 
      BackColor       =   &H000000FF&
      Caption         =   "REGRESAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      MaxLength       =   12
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9015
      _cx             =   15901
      _cy             =   11033
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   1
      HeadingColCount =   0
      TextAlignment   =   5
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   12632256
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   1
      ColorEven       =   -2147483628
      ColorOdd        =   14737632
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   2
      FocusRectColor  =   255
      FocusRectLineWidth=   7
      TabKeyBehavior  =   0
      EnterKeyBehavior=   1
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1219
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   0   'False
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   -1  'True
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   241
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   0
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Arrastre el Titulo de la columna aqui para agrupar por esa columna"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"V2AdmPLUBusqueda.frx":154A
      ColumnsCollection=   $"V2AdmPLUBusqueda.frx":3379
      ValueItems      =   $"V2AdmPLUBusqueda.frx":3CEF
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "SOLO UN ITEM A LA VEZ"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   21
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lbCant 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   19
      Top             =   6660
      Width           =   1695
   End
   Begin VB.Shape ShapeCantidad 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   9360
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "NOMBRE PRODUCTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "AdmPLUBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsBusqueda As New ADODB.Recordset
Dim nCantidadMarcada As Integer
Dim nDepto As Long
Dim nPlu As Long
Dim cItemMarcado As String
Dim MiImpresora As Integer
Dim MiPrecio As Single
Dim MiTax As Single

Dim cDescripAcompa As String
Dim nIDAcompa As Integer
Dim rslocalAcompa As New ADODB.Recordset
Dim nAcoBookMark As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: MANEJO DE TAMAÑOS (9JUN2023)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim rsLocal As ADODB.Recordset
Dim nContenedor As Long
Dim cContenedor As String
Dim MiEnvase As Integer

Private Sub Borrar_Click()
lbCant.Caption = ""
lbCant.BackColor = vbYellow
nCantidadMarcada = 1
End Sub

Private Sub cmdAcomp_Click(Index As Integer)
cDescripAcompa = cmdAcomp(Index).Caption
nIDAcompa = cmdAcomp(Index).Tag
End Sub

Private Sub cmdRegresar_Click()
ShowMsg "SE REGRESA SIN AGREGAR NADA", vbBlue, vbYellow
Unload Me
End Sub

Private Sub cmdRestoAco_Click()
'INFO: MOSTRAR EL RESTO DE LOS ACOMPAÑANTES
Dim iLoc As Integer
Dim iAcom As Integer
Dim nAcoTop As Integer

iLoc = 0: iAcom = 0: nAcoTop = 240
On Error GoTo ErrAdm:

For iLoc = 0 To cmdAcomp.Count - 1
    'LIMPIA LOS 4 ACOMPAÑANTES
    cmdAcomp(iLoc).Caption = ""
    cmdAcomp(iLoc).Tag = 0
    If iLoc > 0 Then
        cmdAcomp(iLoc).Visible = False
    End If
Next
cmdUpAcompa.Enabled = True
'SE PARA EN EL ULTIMO ACOMPAÑANTE
rslocalAcompa.Bookmark = nAcoBookMark
'SE MUEVE AL PROXIMO
rslocalAcompa.MoveNext
Do Until rslocalAcompa.EOF
    If iAcom = 0 Then
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Caption = rslocalAcompa!DESCRIP
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        'NUEVO CODIGO PARA SALIR DESPUES DE 4 o CUANDO rslocalAcompa=EOF
        nAcoBookMark = rslocalAcompa.Bookmark
        rslocalAcompa.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    Else
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Top = nAcoTop + 600
        cmdAcomp(iAcom).Caption = rslocalAcompa!DESCRIP
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        nAcoTop = nAcoTop + 600
        'NUEVO CODIGO PARA SALIR DESPUES DE 4 o CUANDO rslocalAcompa=EOF
        nAcoBookMark = rslocalAcompa.Bookmark
        rslocalAcompa.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
        '---------------------------------------------------------------------------------
    End If
Loop
'cmdRestoAco.Enabled = False
On Error GoTo 0
Exit Sub

ErrAdm:
If Err.Number = 340 Then
    ShowMsg "Este plato tiene demasiados acompañantes" & vbCrLf & "El maximo de acompañantes es 48"
Else
    Resume Next
End If

End Sub

Private Sub cmdUpAcompa_Click()
'INFO: MOSTRAR LOS ACOMPAÑANTES desde el inicio
Dim iLoc As Integer
Dim iAcom As Integer
Dim nAcoTop As Integer

iLoc = 0: iAcom = 0: nAcoTop = 240

On Error Resume Next
rslocalAcompa.MoveFirst
On Error GoTo 0

On Error GoTo ErrAdm:
If rslocalAcompa.EOF Then
    If rslocalAcompa.State = adStateOpen Then rslocalAcompa.Close
    cmdRestoAco.Enabled = False
    cmdAcomp(0).Caption = ""
    cmdAcomp(0).Tag = 0
    For iLocal = 1 To cmdAcomp.Count - 1
        cmdAcomp(iLocal).Visible = False
    Next
    'Frame2(2).Enabled = False
Else
    'ACTIVA FRAME DE ACOMPAÑANTES
    'Frame2(2).Enabled = True
    iAcom = 0: iLocal = 0
    'COMIENZA A CREAR LOS OTROS 3 BOTONES
    For iLocal = 1 To 3
        On Error Resume Next
            Load cmdAcomp(iLocal)
        On Error GoTo 0
    Next
    On Error Resume Next
    Do Until rslocalAcompa.EOF
        'MUESTRA LOS PRIMEROS 4 ACOMPAÑANTES
        'LES ASIGNA EL DEPARTAMENTO DEL PRODUCTO MARCADO
        If iAcom = 0 Then
            cmdAcomp(iAcom).Visible = True
            cmdAcomp(iAcom).Caption = rslocalAcompa!DESCRIP
            cmdAcomp(iAcom).Tag = rs03!DEPTO
            iAcom = iAcom + 1
            rslocalAcompa.MoveNext
            If rslocalAcompa.EOF Then Exit Do
        End If
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Top = nAcoTop + 600
        cmdAcomp(iAcom).Caption = rslocalAcompa!DESCRIP
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        nAcoTop = nAcoTop + 600
        nAcoBookMark = rslocalAcompa.Bookmark
        rslocalAcompa.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    Loop
    On Error GoTo 0
End If
'---------------------------------------------------------------------------------------------------------------
'------------------FIN TRATAMIENTO DE ACOMPAÑANTES--------------------------------
'---------------------------------------------------------------------------------
On Error GoTo 0
Exit Sub

ErrAdm:
If Err.Number = 340 Then
    ShowMsg "Este plato tiene demasiados acompañantes" & vbCrLf & "El maximo de acompañantes es 48"
Else
    Resume Next
End If

End Sub

Private Sub Command1_Click()
Dim vResp As Variant

If lbCant = "" Then nCantidadMarcada = 1
If cItemMarcado = "" Then Exit Sub

If cDescripAcompa = "" Then
    vResp = ShowMsg("DESEA AGREGAR (" & nCantidadMarcada & ")" & vbCrLf & vbCrLf & _
             cItemMarcado & " ? ", vbYellow, vbBlue, vbYesNo)
Else
    vResp = ShowMsg("DESEA AGREGAR (" & nCantidadMarcada & ")" & vbCrLf & vbCrLf & _
        cItemMarcado & vbCrLf & _
        " ... " & cDescripAcompa & " ? ", vbYellow, vbBlue, vbYesNo)
End If

If vResp = vbYes Then
    AddEnBusqueda = True
    Set rsItemAdd = New ADODB.Recordset
    rsItemAdd.Fields.Append "CODIGO", adInteger
    rsItemAdd.Fields.Append "DEPTO", adInteger
    rsItemAdd.Fields.Append "CANT", adInteger
    rsItemAdd.Fields.Append "ENVASE", adInteger
    rsItemAdd.Fields.Append "DESCRIP", adChar, 40
    'rsItemAdd.Fields.Append "CORTO", adChar, 18
    rsItemAdd.Fields.Append "PRECIO", adCurrency
    rsItemAdd.Fields.Append "IMPRESORA", adInteger
    rsItemAdd.Fields.Append "CON_TAX", adInteger
    rsItemAdd.Open
    
    rsItemAdd.AddNew Array("CODIGO", "DEPTO", "CANT", "ENVASE", "DESCRIP", "PRECIO", "IMPRESORA", "CON_TAX"), _
                                Array(nPlu, nDepto, nCantidadMarcada, MiEnvase, cItemMarcado, MiPrecio, MiImpresora, MiTax)
    
    rsItemAdd.Update
    
    If cDescripAcompa = "" Then
    Else
        rsItemAdd.AddNew Array("CODIGO", "DEPTO", "CANT", "ENVASE", "DESCRIP", "PRECIO", "IMPRESORA", "CON_TAX"), _
                                Array(0, nDepto, nCantidadMarcada, MiEnvase, " @@ " & cDescripAcompa, 0#, MiImpresora, MiTax)
        rsItemAdd.Update
    End If
    
    Unload Me
Else
End If
End Sub

Private Sub Command8_Click(Index As Integer)
nCantidadMarcada = nCantidadMarcada + Index
lbCant.Caption = lbCant.Caption + RTrim(LTrim(Str(Index)))
lbCant.BackColor = vbWhite
nCantidadMarcada = Val(lbCant.Caption)
If lbCant.Caption > 99 Then
    ShowMsg "NO SE PUEDE MARCAR ESA CANTIDAD", vbRed, vbYellow
    lbCant.Caption = ""
    lbCant.BackColor = vbYellow
End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DD_PEDDETALLE_Click
' Author    : hsequeira
' Date      : 10/06/2023
' Purpose   : CAMBIOS A LA PROGRAMACION PARA INCLUIR TAMAÑOS
'---------------------------------------------------------------------------------------
'
Private Sub DD_PEDDETALLE_Click()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo DD_PEDDETALLE_Click_Error

cDescripAcompa = "" ' se limpia acompanante
nIDAcompa = 0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
nDepto = DD_PEDDETALLE.Rows.Current.Cells(4).value
nPlu = DD_PEDDETALLE.Rows.Current.Cells(5).value

If DD_PEDDETALLE.Rows.Current.Cells(3) = "" Then
    cItemMarcado = DD_PEDDETALLE.Rows.Current.Cells(1).value
    MiEnvase = 0
Else
    cItemMarcado = DD_PEDDETALLE.Rows.Current.Cells(1).value & "-" & DD_PEDDETALLE.Rows.Current.Cells(3).value
    MiEnvase = DD_PEDDETALLE.Rows.Current.Cells(6).value    'ENVASE
End If
'cItemMarcado = DD_PEDDETALLE.Rows.Current.Cells(1).value
MiPrecio = DD_PEDDETALLE.Rows.Current.Cells(2).value
MiImpresora = DD_PEDDETALLE.Rows.Current.Cells(7).value
MiTax = DD_PEDDETALLE.Rows.Current.Cells(8).value
'Array(nPlu, nDepto, ncant, citem, Left(citem, 18), MiPrecio, MiImpresora, MiTax)


Call BuscaAcompa(nDepto, nPlu)


'Clipboard.SetText DD_PEDDETALLE.Rows.Current.Cells(0).Text
''rsBusqueda.Close
'Set rsBusqueda = Nothing
'Unload Me
 On Error GoTo 0
   Exit Sub

DD_PEDDETALLE_Click_Error:
    If Err.Number = -2147024809 Then
        ShowMsg "PRIMERO DEBE BUSCAR UN PRODUCTO", vbYellow, vbRed
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.Description & ") in AdmPLUBusqueda"
    End If
    txtNombre.SetFocus
End Sub
Private Sub BuscaAcompa(n_Depto As Long, n_Plu As Long)
Dim CadenaSql As String
Dim iLocal As Integer
Dim iAcom As Integer
Dim nAcoTop As Integer

    iLocal = 0: iAcom = 0: nAcoTop = 240

On Error Resume Next
rslocalAcompa.Close
On Error GoTo 0

    CadenaSql = "SELECT A.PLU_ID,A.ACOMP_ID,B.DESCRIP "
    CadenaSql = CadenaSql & " FROM PLU_ACOMP AS A, ACOMPA AS B "
    CadenaSql = CadenaSql & " WHERE A.PLU_ID = " & n_Plu
    CadenaSql = CadenaSql & " AND A.ACOMP_ID = B.CODIGO "
    CadenaSql = CadenaSql & " ORDER BY B.DESCRIP "
    rslocalAcompa.Open CadenaSql, msConn, adOpenStatic, adLockReadOnly

Frame2.Enabled = True
cmdRestoAco.Enabled = True
cmdUpAcompa.Enabled = True

'BORRA TODOS LOS ACOMPAÑANTES
'INFO: ABRIL2010
For iLocal = 1 To cmdAcomp.Count - 1
    cmdAcomp(iLocal).Caption = ""
    cmdAcomp(iLocal).Tag = 0
    cmdAcomp(iLocal).Visible = False
Next

If rslocalAcompa.EOF Then
    If rslocalAcompa.State = adStateOpen Then rslocalAcompa.Close
    cmdRestoAco.Enabled = False
    cmdAcomp(0).Caption = ""
    cmdAcomp(0).Tag = 0
    For iLocal = 1 To cmdAcomp.Count - 1
        cmdAcomp(iLocal).Visible = False
    Next
    Frame2.Enabled = False
Else
    'ACTIVA FRAME DE ACOMPAÑANTES
    Frame2.Enabled = True
    iAcom = 0: iLocal = 0
    'COMIENZA A CREAR LOS OTROS 3 BOTONES
    For iLocal = 1 To 3
        On Error Resume Next
            Load cmdAcomp(iLocal)
        On Error GoTo 0
    Next
    On Error Resume Next
    Do Until rslocalAcompa.EOF
        'MUESTRA LOS PRIMEROS 4 ACOMPAÑANTES
        'LES ASIGNA EL DEPARTAMENTO DEL PRODUCTO MARCADO
        If iAcom = 0 Then
            cmdAcomp(iAcom).Visible = True
            cmdAcomp(iAcom).Caption = rslocalAcompa!DESCRIP
            cmdAcomp(iAcom).Tag = n_Depto
            iAcom = iAcom + 1
            rslocalAcompa.MoveNext
            If rslocalAcompa.EOF Then Exit Do
        End If
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Top = nAcoTop + 600
        cmdAcomp(iAcom).Caption = rslocalAcompa!DESCRIP
        cmdAcomp(iAcom).Tag = n_Depto
        iAcom = iAcom + 1
        nAcoTop = nAcoTop + 600
        nAcoBookMark = rslocalAcompa.Bookmark
        rslocalAcompa.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    Loop
    On Error GoTo 0
End If
'---------------------------------------------------------------------------------
'------------------FIN TRATAMIENTO DE ACOMPAÑANTES--------------------------------
'---------------------------------------------------------------------------------
    
    
    
End Sub

Private Sub DD_PEDDETALLE_OnInit()
DD_PEDDETALLE.VScrollWidth = DD_PEDDETALLE.ClientHeight / 10
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : hsequeira
' Date      : 11/05/2023
' Purpose   : CIERRA EL ITEM DE BUSQUEDA
'---------------------------------------------------------------------------------------

Private Sub Form_Load()
On Error Resume Next
rsItemAdd.Close
On Error GoTo 0

Set rsLocal = New ADODB.Recordset
Set rsLocal = rsFullMenu.Clone(adLockOptimistic)

cItemMarcado = ""
AddEnBusqueda = False
End Sub

Private Sub txtNombre_Change()

   On Error GoTo txtNombre_Change_Error

If txtNombre = "" Then
    'NO HACE NADA
    cItemMarcado = ""
    Set DD_PEDDETALLE.DataSource = Nothing
   
    DD_PEDDETALLE.ReBind
    
Else
    If App.EXEName = "SOLORETAIL" Then
        cSQL = "SELECT A.NOMBRE, FORMAT(A.PRECIO,'STANDARD') AS PRECIO_BASE,"
        cSQL = cSQL & " A.NOMLIN , A.NOMFAM, A.NOMMAR "
        cSQL = cSQL & " FROM PLU_RETAIL AS A "
        cSQL = cSQL & " WHERE A.NOMBRE LIKE '%" & txtNombre & "%' "
        cSQL = cSQL & " ORDER BY 1,3,4,5"
    ElseIf App.EXEName = "FastRETAIL" Then
        cSQL = "SELECT B.BARCODE, B.DESCRIP AS PRODUCTO,  "
        cSQL = cSQL & " B.PRECIO1 AS PRECIO_BASE, iif(B.ENVASES,'OTROS Tamaños','Tamaño Unico') AS INFO "
        cSQL = cSQL & " FROM DEPTO AS A, PLU AS B "
        cSQL = cSQL & " WHERE B.DESCRIP LIKE '%" & txtNombre & "%' "
        cSQL = cSQL & " AND B.DISPONIBLE "
        cSQL = cSQL & " AND B.DEPTO = A.CODIGO "
        cSQL = cSQL & " ORDER BY 1,2"
    Else
        '        cSQL = "SELECT A.DESCRIP AS DEPTO, B.DESCRIP AS PRODUCTO,  "
        '        cSQL = cSQL & " B.PRECIO1 AS PRECIO_BASE, "
        '        cSQL = cSQL & " iif(B.ENVASES,'OTROS Tamaños','Tamaño Unico') AS INFO, "
        '        cSQL = cSQL & " A.CODIGO AS C_DEPTO, B.CODIGO AS C_PLU, "
        '        cSQL = cSQL & " B.IMPRESORA, B.CON_TAX "
        '        cSQL = cSQL & " FROM DEPTO AS A, PLU AS B "
        '        cSQL = cSQL & " WHERE B.DESCRIP LIKE '%" & txtNombre & "%' "
        '        cSQL = cSQL & " AND B.DISPONIBLE "
        '        cSQL = cSQL & " AND B.DEPTO = A.CODIGO "
        '        cSQL = cSQL & " ORDER BY 1,2"
    End If
    
'   rsBusqueda.Open cSQL, msConn, adOpenStatic, adLockReadOnly
'
    
'    DD_PEDDETALLE.DataMode = sgBound
'    Set DD_PEDDETALLE.DataSource = rsBusqueda
'
'
'    DD_PEDDETALLE.ReBind
'    cItemMarcado = ""
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: MANEJO DE TAMAÑOS (30MAY2023)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    rsLocal.Filter = "PRODUCTO LIKE '%" & txtNombre & "%' AND DISPONIBLE = TRUE "
    
    DD_PEDDETALLE.DataMode = sgUnbound
    DD_PEDDETALLE.DataRowCount = 0
    DD_PEDDETALLE.DataColCount = 11
    rsLocal.MoveFirst
    Do While Not rsLocal.EOF
        If IsNull(rsLocal!DESCRIP) Then nPrecio = rsLocal!PRECIO1 Else nPrecio = rsLocal!precio
        If IsNull(rsLocal!CONTENEDOR) Then
            nContenedor = 0
            cContenedor = ""
        Else
            nContenedor = rsLocal!CONTENEDOR
            cContenedor = rsLocal!DESCRIP
        End If
        cData = rsLocal!DEPTO_DESCRIP & "|" & rsLocal!PRODUCTO & "|" & Format(nPrecio, "STANDARD") & "|" & cContenedor & "|" & _
                    rsLocal!DEPTO & "|" & rsLocal!CODIGO & "|" & nContenedor & "|" & rsLocal!IMPRESORA & "|" & _
                    rsLocal!CON_TAX & "|" & String(4, Chr(126))
        DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, "|"
        rsLocal.MoveNext
    Loop
    
    cItemMarcado = ""
    
    
    DD_PEDDETALLE.Columns(1).Caption = "DEPARTAMENTO"
    DD_PEDDETALLE.Columns(2).Caption = "PRODUCTO"
    DD_PEDDETALLE.Columns(3).Caption = "PRECIO"
    DD_PEDDETALLE.Columns(4).Caption = "ENVASE"
    
    
     On Error Resume Next
     With DD_PEDDETALLE
        .ColumnClickSort = False
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
        .Columns(1).Width = 1800:        'DEPTO_DESCRIP
        .Columns(1).Style.WordWrap = True
        .Columns(2).Width = 3400:        'PRODUCTO
        .Columns(2).Style.WordWrap = True
        .Columns(3).Width = 1000:     'PRECIO
        .Columns(4).Width = 2500:   'ENVASE DESCRIP
        .Columns(5).Width = 0:   'DEPTO
        .Columns(6).Width = 0:   'PLU
        .Columns(7).Width = 0:   'ENVASE CODIGO
        .Columns(8).Width = 0:   'IMPRESORA
        .Columns(9).Width = 0:   'CON_TAX
        .Columns(10).Width = 0:   'FILLER
        .Columns(11).Width = 0:   'FILLER
    End With
    On Error GoTo 0
    
    
    DD_PEDDETALLE.RowHeightMin = 650
    DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
    'rsBusqueda.Close
End If

   On Error GoTo 0
   Exit Sub

txtNombre_Change_Error:

    If txtNombre.Text = "" Then
          DD_PEDDETALLE.DataRowCount = 0
    Else
        If rsLocal.EOF Then
              DD_PEDDETALLE.DataRowCount = 0
        Else
            ShowMsg "Error " & Err.Number & " (" & Err.Description & ") in Meseros.AdmPLUBusqueda", vbYellow, vbRed
        End If
    End If

End Sub
