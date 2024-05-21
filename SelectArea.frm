VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form SelectArea 
   BackColor       =   &H00B39665&
   BorderStyle     =   0  'None
   Caption         =   "SELECCIONE EL AREA DE TRABAJO"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddArea 
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2190
      TabIndex        =   1
      Top             =   7560
      Width           =   3375
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_AREAS 
      Height          =   6975
      Left            =   1110
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      _cx             =   9763
      _cy             =   12303
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
      ColorEven       =   16744576
      ColorOdd        =   16744576
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
      CellBackColor   =   16744576
      CellForeColor   =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
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
      StylesCollection=   $"SelectArea.frx":0000
      ColumnsCollection=   $"SelectArea.frx":1E25
      ValueItems      =   $"SelectArea.frx":279B
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00000040&
      BorderWidth     =   15
      Height          =   8535
      Left            =   100
      Top             =   100
      Width           =   7575
   End
End
Attribute VB_Name = "SelectArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData(nValorInicial As Integer)
Dim cSQL As String, cSQL2 As String
Dim rsAreas As ADODB.Recordset
Dim rsMesas As ADODB.Recordset
Dim i As Integer
Dim rowData As String

cSQL = "SELECT * FROM AREAS ORDER BY DESCRIPCION"

Set rsAreas = New ADODB.Recordset
rsAreas.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Set DD_AREAS.DataSource = Nothing
   
DD_AREAS.ReBind
DD_AREAS.DataMode = sgBound
Set DD_AREAS.DataSource = rsAreas
DD_AREAS.ReBind
DD_AREAS.RowHeightMin = 685
DD_AREAS.TextAlignment = sgAlignCenterCenter

 On Error Resume Next
 With DD_AREAS
    .ColumnClickSort = False
    .Columns(1).Width = 0:        'DEPTO
    '.Columns(1).Style.WordWrap = True
    .Columns(2).Width = 5500:        'AREA
    .Columns(2).Style.WordWrap = True
End With
On Error GoTo 0


DD_AREAS.RowHeightMin = 685
DD_AREAS.TextAlignment = sgAlignCenterCenter
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
rsAreas.Close

Set rsAreas = Nothing
End Sub


Private Sub cmdAddArea_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

nArea = DD_AREAS.Rows.Current.Cells(0).value
cArea = DD_AREAS.Rows.Current.Cells(1).value

If ShowMsg("¿ ENTRAR EN ESTA AREA ?" & vbCrLf & vbCrLf & cArea, vbYellow, vbBlue, vbYesNo) = vbYes Then
    BoxResp = vbYes
Else
    BoxResp = vbNo
End If
If BoxResp = vbYes Then
    Unload Me
Else
End If

End Sub

Private Sub Form_Load()
Call LoadData(0)
End Sub
