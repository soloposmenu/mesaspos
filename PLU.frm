VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PLU 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE VENTAS"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   14295
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "PLU.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCtas 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13150
      Picture         =   "PLU.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Correccion 
      Caption         =   "Correción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2760
      Picture         =   "PLU.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8235
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestoAco 
      DisabledPicture =   "PLU.frx":0890
      Enabled         =   0   'False
      Height          =   615
      Left            =   13080
      Picture         =   "PLU.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Proximos Acompañantes"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpAcompa 
      DisabledPicture =   "PLU.frx":1114
      Enabled         =   0   'False
      Height          =   615
      Left            =   11640
      Picture         =   "PLU.frx":1556
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Inicio de Acompañantes"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "ACOMPAÑANTES"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2930
      Index           =   2
      Left            =   11640
      TabIndex        =   7
      Top             =   2430
      Width           =   2520
      Begin VB.CommandButton cmdAcomp 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComctlLib.StatusBar StatBar 
      Height          =   290
      Left            =   3310
      TabIndex        =   52
      Top             =   2400
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7232
            MinWidth        =   7232
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1323
            MinWidth        =   1323
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdNota 
      BackColor       =   &H00C0C0FF&
      Caption         =   "NOTAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   11100
      TabIndex        =   50
      Top             =   210
      Width           =   735
   End
   Begin VB.CommandButton GridDOWN 
      Height          =   615
      Left            =   10180
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Bajar Lista de Productos"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton GridUP 
      Height          =   615
      Left            =   10180
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Subir Lista de Productos"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   13350
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7365
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7365
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7365
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   13350
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6765
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6765
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6765
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   13350
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6165
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6165
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6165
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7980
      Width           =   855
   End
   Begin VB.CommandButton cmdFacturacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Clear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11100
      TabIndex        =   34
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H000000FF&
      Caption         =   "CORTESIA"
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
      Index           =   6
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5475
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   2640
      Width           =   9675
      Begin VB.CommandButton cmdPlus 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Anulación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   1
      Left            =   1470
      Picture         =   "PLU.frx":1998
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8235
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Pre-CUENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   3
      Left            =   120
      Picture         =   "PLU.frx":1DDA
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8235
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   6
      Left            =   14280
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   23
         Top             =   195
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   18
         Top             =   520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cuenta Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   27
         Top             =   1995
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbCuenta 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   26
         Top             =   1900
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Mesa #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cajer@"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   22
         Top             =   900
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Hora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   -5760
         TabIndex        =   20
         Top             =   -1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Meser@"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Top             =   550
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid PlatosMesa 
      Height          =   2175
      Left            =   3310
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   6800
      _ExtentX        =   11986
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   65280
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      GridColor       =   16777215
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TAMAÑO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2600
      Index           =   4
      Left            =   1920
      TabIndex        =   8
      Top             =   0
      Width           =   1360
      Begin VB.CommandButton cmdEnvases 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnvases 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnvases 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnvases 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSelMesa 
      Caption         =   "MESAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   10
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5250
      Picture         =   "PLU.frx":20E4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4170
      Picture         =   "PLU.frx":33E6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      Caption         =   "Departamentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7935
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1900
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   50
         Picture         =   "PLU.frx":46E8
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   7080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   960
         Picture         =   "PLU.frx":59EA
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   7080
         Width           =   855
      End
      Begin VB.CommandButton cmdDepto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   635
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1750
      End
   End
   Begin VB.CommandButton cmdSlip 
      Caption         =   "CHEF"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   29
      Top             =   960
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   6240
      TabIndex        =   53
      Top             =   8760
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image ImageLUPA 
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   13440
      Picture         =   "PLU.frx":6CEC
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda de Productos"
      Top             =   1695
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Lin          Platos en la Mesa                                                Cant        P.Unit           Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3315
      TabIndex        =   16
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "CANT."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   11160
      TabIndex        =   51
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   915
      Left            =   0
      Top             =   8160
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   6720
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label SubTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   9075
      TabIndex        =   28
      Top             =   8160
      Width           =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "Sub-Tot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11160
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private num As Integer
Private numplu As Integer
Private nNLinSel As Integer
Private Arreg_Deptos(10) As Long
'Private Arreg_Plu(17) As Integer
'INFO: 1024 ---> Private Arreg_Plu(23) As Integer
Private Arreg_Plu(23) As Integer
Private nPase As Integer 'Cantidad de Clicks a Cantidad
Private ElDepto As Long 'Es el Departamento Seleccionado
Private nGlobEnv As Long    'El envase seleccionado
Private TextEnv As String
Private rsTmpAco As New ADODB.Recordset
Private nAcoBookMark As Variant
Private lGo As Boolean
Private nCortesia As Integer
Private Function InsertIntoPED_PRINT(cConsec As String, cFilename As String, cNameCocina As String, cText As String, nPrinter As Byte, oIsPrinted As Boolean) As Boolean
'InsertIntoPED_PRINT cConsecutivo, KITCHEN_FILE, NOM_PRN_COCINA, Space(2), nSelectedPrinter, False
Dim cFecha As String
Dim cHora As String
Dim cSQL As String

cFecha = Format(Date, "YYYYMMDD")
cHora = Format(Time, "HH:MM:SS")

cSQL = "INSERT INTO PED_PRINT (NUM_PEDIDO,FILE_NAME,"
cSQL = cSQL & "DESCRIP_PRINTER , TRANS_DATE, TRANS_TIME, "
cSQL = cSQL & "FULL_TEXT, PRINTER, PRINTED) VALUES ('"
cSQL = cSQL & cConsec & "','" & cFilename & "','" & cNameCocina & "','"
cSQL = cSQL & cFecha & "','" & cHora & "','"
cSQL = cSQL & cText & "'," & nPrinter & "," & False & ")"

Call SOLOTrans("BEGIN")
msConn.Execute cSQL
Call SOLOTrans("COMMIT")

End Function
Private Function PutClientesOnMesa(ByRef NumClientes As Integer) As Boolean
'INFO: GUEST_COUNTER = CONTADOR DIARIO DE CLIENTES
'GUEST_TOTAL = CONTADOR ACUMULADO DE CLIENTES
'UPDATE : 03/ABR/2005
Dim cSQL As String

On Error Resume Next
'''cSQL = "UPDATE MESAS SET GUEST_COUNTER = GUEST_COUNTER + " & NumClientes
'''cSQL = cSQL & ", GUEST_TOTAL = GUEST_TOTAL + " & NumClientes
'''cSQL = cSQL & " WHERE NUMERO = " & nMesa
'''cSQL = cSQL & " OR NUMERO = -99 "

cSQL = "UPDATE MESAS SET "
cSQL = cSQL & " GUEST_COUNTER = IIF(ISNULL(GUEST_COUNTER)," & NumClientes & ", GUEST_COUNTER + " & NumClientes & ")"
'cSQL = cSQL & " GUEST_COUNTER = " & NumClientes
cSQL = cSQL & ",GUEST_TOTAL = IIF(ISNULL(GUEST_TOTAL)," & NumClientes & ", GUEST_TOTAL + " & NumClientes & ")"
cSQL = cSQL & " WHERE NUMERO = " & nMesa
cSQL = cSQL & " OR NUMERO = -99 "

'msConn.Execute cSQL
Call SQL_Update(False, "PutClientesOnMesa", cSQL)

NumClientes = 1
On Error GoTo 0

End Function

Private Sub AddPrecuenta2Grid()
'marca una linea mas en el Grid
'14OCT2002.9:45 PM
'INFO: AGREGANDO CON_TAX, VALOR EN 0% (MAY/2006)
Dim cSQL As String
    CajLin = CajLin + 1

    SOLO_FECHA = Format(Date, "YYYYMMDD")
    
    cSQL = "INSERT INTO TMP_TRANS "
    cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX) VALUES ("
    cSQL = cSQL & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & ","
    cSQL = cSQL & "'  PRE-CUENTA',0,0,0,0,0,0,'" & SOLO_FECHA & "','" & Time & "'"
    'cSQL = cSQL & ",'  '," & 0# & "," & nCta & ",FALSE,0,0)"
    'INFO: CAMBIANDO IMPRESO A TRUE, PARA QUE NO LE PUEDAN HACER ERROR CORRECT
    '08/ABR/2007
    cSQL = cSQL & ",'  '," & 0# & "," & nCta & ",TRUE,0,0)"

Call SOLOTrans("BEGIN")
msConn.Execute cSQL
Call SOLOTrans("COMMIT")
End Sub
'---------------------------------------------------------------------------------------
' Procedure : DGI_ImprPreCta
' Author    : hsequeira
' Date      : 02/09/2012
' Purpose   : 'INFO: 2SEP2012
' Cambiando la Precuenta segun ley de la DGI
' Los cambios se marcan con: '=========
'---------------------------------------------------------------------------------------
'
Private Sub DGI_ImprPreCta()

Dim sqltext As String
Dim LinTx As String
Dim rsPreCta As ADODB.Recordset
Dim MiMatriz(0, 3) As String
Dim MiLen1, Milen2 As Integer
Dim MiPropina As Single, nProp As Single
Dim rsParciales As ADODB.Recordset
Dim lParc As Integer
Dim nLinCta As Integer
Dim nTotCta As Single
Dim nTotProp As Single
Dim nErrCount As Integer
Dim i As Integer
'MAY/2006
Dim rsTAX As ADODB.Recordset
Dim nSubtotalMesa As Single
Dim nLDESCUENTO As Single
'27/08/2005
Dim HayPrinterLocal As Boolean
Dim nFreefile As Integer
'02/JUN/2006
Dim nTempISC As Single
Dim cErrorText As String
'INFO: DOMICILIO
Dim aDomiInfo() As Variant
'INFO: Muestra INFO de Propina
Dim aPropina As Variant                       'CARGA LOS % DE LA PROPINA
Dim bPropinaAparte As Boolean           'DETERMINA SI IMPRIME LA PROPINA APARTE
Dim nSUMA As Single
'INFO: 23ENE2011
Dim bDetalleTax As Boolean  'DEFINE SI SE IMPRIME C/U DE LOS IMPUESTOS

nSubtotalMesa = GetSubTotalNOTAXFromMesa()

If NOM_PRN_FACTURA = "" Or NOM_PRN_FACTURA = " " Then
    'NO HAY IMPRESORA DE PRE-CUENTAS, IMPRIMIR EN
    'LA DE FACTURACION
    HayPrinterLocal = False
Else
    HayPrinterLocal = True
End If

If HayPrinterLocal Then
    If LoginMesas.ImpresoraCuentas.RecNearEnd = True Or LoginMesas.ImpresoraCuentas.RecEmpty = True Then
        ShowMsg "Advertencia de Papel" & vbCrLf & "POR FAVOR REVISE EL PAPEL (RECIBO) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO"
    End If
End If

'============================================
'INFO: MAYO 2010
'============================================
aPropina = Split(GetFromINI("Facturacion", "PropinaAparte", App.Path & "\soloini.ini"), ",")

If UBound(aPropina) >= 0 Then
    bPropinaAparte = True
Else
    bPropinaAparte = False
End If
'============================================

'============================================
'INFO: ENERO 2011
'============================================
bDetalleTax = True
If UCase(GetFromINI("Facturacion", "DetalleTax", App.Path & "\soloini.ini")) = "NO" Then
    bDetalleTax = False
Else
    bDetalleTax = True
End If
'============================================

Set rsPreCta = New ADODB.Recordset
Set rsParciales = New ADODB.Recordset

OKCancelar = 0
'INFO: PANTALLA DE PRECUENTA NO ESTA DISPONIBLE EN LAS ESTACIONES DE MESERO
'YA QUE SE IMPRIME DIRECTAMENTE SIN PREGUNTAR NADA
If OPEN_PROPINA = True Then
    ''Opc01.Show 1
End If

nTotProp = 0#

On Error GoTo AdmErr:
'PROPINA_DESCRIP
If OKCancelar = 1 Then OKCancelar = 0: Exit Sub

sqltext = "SELECT MESA,SUM(MONTO) AS VALOR "
sqltext = sqltext & " FROM TMP_PAR_PAGO "
sqltext = sqltext & " WHERE MESA = " & nMesa
sqltext = sqltext & " GROUP BY MESA"

rsParciales.Open sqltext, msConn, adOpenDynamic, adLockOptimistic

'VERIFICA SI HAY PAGOS PARCIALES
If rsParciales.EOF Then lParc = 0 Else lParc = 1

sqltext = "SELECT MESA,CUENTA,LIN,DESCRIP,CANT,PRECIO "
sqltext = sqltext & " FROM TMP_TRANS "
sqltext = sqltext & " WHERE MESA = " & nMesa
sqltext = sqltext & " AND CUENTA = " & nCta
sqltext = sqltext & " ORDER BY CUENTA,LIN "

rsPreCta.Open sqltext, msConn, adOpenStatic, adLockReadOnly
If rsPreCta.EOF Then
    ShowMsg "NO HAY PLATOS EN LA MESA. PRE-CUENTA NO SE IMPRIMIRA"
    'INFO: 25/4/2006
    rsParciales.Close
    Set rsParciales = Nothing
    Exit Sub
End If
rsPreCta.MoveFirst
nLinCta = rsPreCta!CUENTA

'INFO: ELIMINANDO LA IMPRESION DE LA PRE-CUENTA
'JUNIO 11 2010
'Call AddPrecuenta2Grid

If HayPrinterLocal Then
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, rs00!descrip & Chr(&HD) & Chr(&HA)
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, rs00!RAZ_SOC & Chr(&HD) & Chr(&HA)
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "RUC:" & rs00!RUC & Chr(&HD) & Chr(&HA)
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "SERIAL:" & rs00!SERIAL & Chr(&HD) & Chr(&HA)
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Chr(27) & "|3C" & "         PRE-CUENTA" & Chr(&HD) & Chr(&HA)
    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Format(Date, "SHORT DATE") & "     " & Time & Chr(&HD) & Chr(&HA)
Else
    nFreefile = FreeFile
    'Open DATA_PATH + FACTURA_FILE For Output Shared As #nFreeFile
    'INFO: SI NO HAY IMPRESORA LOCAL, LA PRE-CUENTA HAY QUE IMPRIMIRLA EN LA CAJA
    'SEP2009
    'Open App.Path & "\" & FACTURA_FILE For Output Shared As #nFreefile
    'INFO: 29MAR2017
    Open App.Path & "\" & npNumCaj & "_" & FACTURA_FILE For Output Shared As #nFreefile
    Print #nFreefile, Space(2)
    
    '========= Print #nFreefile, rs00!descrip
    '========= Print #nFreefile, rs00!RAZ_SOC
    '========= Print #nFreefile, "RUC:" & rs00!RUC
    '========= Print #nFreefile, "SERIAL:" & rs00!SERIAL
    '========= Print #nFreefile, Space(2)
    '========= Print #nFreefile, Chr(27) & "|3C" & "         PRE-CUENTA"
    Print #nFreefile, Format(Date, "SHORT DATE") & "     " & Time
End If

If bISThisSocios Then
    'INFO: ES RESTAURANTE CON SOCIOS
    Dim rsSocio As ADODB.Recordset
    Set rsSocio = New ADODB.Recordset
    rsSocio.Open "SELECT * FROM SOCIO WHERE MESA = " & nMesa, msConn, adOpenStatic, adLockOptimistic
    If Not rsSocio.EOF Then
        If HayPrinterLocal Then
            LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
            LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "SOCIO : " & rsSocio!SOCIO_NOMBRE & Chr(&HD) & Chr(&HA)
            LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
        Else
            Print #nFreefile, Space(2)
            Print #nFreefile, "SOCIO : " & rsSocio!SOCIO_NOMBRE
            Print #nFreefile, Space(2)
        End If
    End If
    rsSocio.Close
    Set rsSocio = Nothing
End If

If HayPrinterLocal Then
    
    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Mesero : " & cNomMesero & Chr(&HD) & Chr(&HA)
    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
    
    '~~~~~~~~~~~~~~
    'INFO DOMICILIO
    '~~~~~~~~~~~~~~
    If HAS_Domicilio Then
        If nMesa >= nDomicilio Then
            
            If HAS_Domicilio Then
                If MesaAssigned() = "" Then
                    'INFO: SI NO HAY CLIENTE ASIGNADO, ENTONCES NO INTENTA IMPRIMIR LA INFO DEL CLIENTE.
                Else
                    aDomiInfo = GetDomicilioInfo()
        
                    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "PEDIDO : " & nMesa & "     " & aDomiInfo(14, 0) & Chr(&HD) & Chr(&HA)
                    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Nombre  : " & aDomiInfo(3, 0) & Space(1) & aDomiInfo(4, 0) & Chr(&HD) & Chr(&HA)
                    If aDomiInfo(6, 0) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Empresa : " & aDomiInfo(6, 0) & Chr(&HD) & Chr(&HA)
                    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Telefono: " & FormatPhone(aDomiInfo(1, 0)) & IIf(aDomiInfo(2, 0) <> "", " (" & aDomiInfo(2, 0) & ")", Space(1)) & Chr(&HD) & Chr(&HA)
                    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "ZONA: " & aDomiInfo(12, 0) & Chr(&HD) & Chr(&HA)
                    
                    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Dir: " & Left(aDomiInfo(8, 0), 25) & Chr(&HD) & Chr(&HA)
                    If Mid(aDomiInfo(8, 0), 26, 25) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "- " & Mid(aDomiInfo(8, 0), 26, 25) & Chr(&HD) & Chr(&HA)
                    
                    If aDomiInfo(9, 0) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "- " & Left(aDomiInfo(9, 0), 25) & Chr(&HD) & Chr(&HA)
                    If Mid(aDomiInfo(9, 0), 26, 25) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "- " & Mid(aDomiInfo(9, 0), 26, 25) & Chr(&HD) & Chr(&HA)
                    
                    If aDomiInfo(10, 0) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "- " & Left(aDomiInfo(10, 0), 25) & Chr(&HD) & Chr(&HA)
                    If Mid(aDomiInfo(10, 0), 26, 25) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "- " & Mid(aDomiInfo(10, 0), 26, 25) & Chr(&HD) & Chr(&HA)
                    
                    If Left(aDomiInfo(11, 0), 25) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "INFO:" & Left(aDomiInfo(11, 0), 25) & Chr(&HD) & Chr(&HA)
                    If Mid(aDomiInfo(11, 0), 26, 25) <> "" Then LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "- " & Mid(aDomiInfo(11, 0), 26, 25) & Chr(&HD) & Chr(&HA)
                    'INFO: 23ENE2011
                    If IsNull(aDomiInfo(15, 0)) Then
                        'DO NOTHING, NO SE HA ASIGNADO MOTORIZADO
                    Else
                        LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Motorizado: " & GetMotorizado(CLng(aDomiInfo(15, 0))) & Chr(&HD) & Chr(&HA)
                    End If
                End If
            End If
        Else
            LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Mesa : " & nMesa & Chr(&HD) & Chr(&HA)
            LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
        End If
    Else
        LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Mesa : " & nMesa & Chr(&HD) & Chr(&HA)
        LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
    End If
Else
    Print #nFreefile, "Mesero : " & cNomMesero
    Print #nFreefile, Space(2)

    '~~~~~~~~~~~~~~
    'INFO DOMICILIO
    '~~~~~~~~~~~~~~
    If HAS_Domicilio Then
        If nMesa >= nDomicilio Then
            
            aDomiInfo = GetDomicilioInfo()
            
            Print #nFreefile, "PEDIDO : " & nMesa & "     " & aDomiInfo(14, 0)
            Print #nFreefile, "Nombre  : " & aDomiInfo(3, 0) & Space(1) & aDomiInfo(4, 0)
            If aDomiInfo(6, 0) <> "" Then Print #nFreefile, "Empresa : " & aDomiInfo(6, 0)
            Print #nFreefile, "Telefono: " & FormatPhone(aDomiInfo(1, 0)) & IIf(aDomiInfo(2, 0) <> "", " (" & aDomiInfo(2, 0) & ")", Space(1))
            Print #nFreefile, "ZONA: " & aDomiInfo(12, 0)
            
            Print #nFreefile, "Dir: " & Left(aDomiInfo(8, 0), 25)
            If Mid(aDomiInfo(8, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(8, 0), 26, 25)
            
            If aDomiInfo(9, 0) <> "" Then Print #nFreefile, "- " & Left(aDomiInfo(9, 0), 25)
            If Mid(aDomiInfo(9, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(9, 0), 26, 25)
            
            If aDomiInfo(10, 0) <> "" Then Print #nFreefile, "- " & Left(aDomiInfo(10, 0), 25)
            If Mid(aDomiInfo(10, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(10, 0), 26, 25)
            
            If Left(aDomiInfo(11, 0), 25) <> "" Then Print #nFreefile, "INFO:" & Left(aDomiInfo(11, 0), 25)
            If Mid(aDomiInfo(11, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(11, 0), 26, 25)
            'INFO: 23ENE2011
            If IsNull(aDomiInfo(15, 0)) Then
                'DO NOTHING, NO SE HA ASIGNADO MOTORIZADO
            Else
                Print #nFreefile, "Motorizado: " & GetMotorizado(CLng(aDomiInfo(15, 0)))
            End If
        Else
            Print #nFreefile, "Mesa : " & nMesa
            Print #nFreefile, Space(2)
        End If
    Else
        'INFO: 05/MAY/2009 - 23ENE2011
        Print #nFreefile, "Mesa : " & nMesa
        Print #nFreefile, Space(2)
    End If
End If

If nLinCta <> 0 Then
    'INFO: HAY CUENTAS SEPARADAS
    If HayPrinterLocal Then
        LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Cuenta # : " & nLinCta & Chr(&HD) & Chr(&HA)
    Else
        Print #nFreefile, "Cuenta # : " & nLinCta
    End If
End If

If HayPrinterLocal Then
    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(nEspacio) & "------------------------------" & Chr(&HD) & Chr(&HA)
Else
    Print #nFreefile, Space(nEspacio) & "------------------------------"
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: DETALLE DE LA CUENTA
'~~~~~~~~~~~~~~~~~~~~~~~~~~
Do Until rsPreCta.EOF
    Do Until rsPreCta.EOF
        'MiMatriz(0, 0) = FormatTexto(rsPreCta!descrip, 22)      'DE 15 22
        '========= MiMatriz(0, 1) = Format(rsPreCta!cant, "general number")
        'MiMatriz(0, 1) = "(" & Format(rsPreCta!cant, "GENERAL NUMBER") & ")"
        '========= MiMatriz(0, 2) = Format(rsPreCta!precio, "#,###.00")
        '========= MiLen1 = Len(MiMatriz(0, 1))
        '========= Milen2 = Len(MiMatriz(0, 2))
        'LinTx = MiMatriz(0, 0) & Space(5 - MiLen1) & MiMatriz(0, 1) & Space(10 - Milen2) & MiMatriz(0, 2)
        
        LinTx = Format(FormatTexto(rsPreCta!descrip, 22), "@@@@@@@@@@@@@@@@@@@@@@")
        LinTx = LinTx & Space(3) & Format("(" & Format(rsPreCta!CANT, "GENERAL NUMBER") & ")", "@@@@@@")
        
        If HayPrinterLocal Then
            LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, LinTx & Chr(&HD) & Chr(&HA)
        Else
            Print #nFreefile, LinTx
        End If
        'INFO: ACUMULA SUBTOTAL SIN IMPUESTO
        nTotCta = nTotCta + rsPreCta!precio
        rsPreCta.MoveNext
        If rsPreCta.EOF Then Exit Do
        If nLinCta <> rsPreCta!CUENTA Then
            nLinCta = rsPreCta!CUENTA
            Exit Do
        End If
    Loop
    
    If nLinCta <> 0 Then
        'INFO: HAY CUENTAS SEPARADAS
        MiLen1 = Len(Format(nTotCta, "STANDARD"))
        nProp = 0#
        
        nTotCta = nTotCta + nProp
        
        If Not rsPreCta.EOF Then
            For i = 1 To 10
                If HayPrinterLocal Then
                    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
                Else
                    Print #nFreefile, Space(2)
                End If
            Next
            If HayPrinterLocal Then
                LoginMesas.ImpresoraCuentas.CutPaper 100
                LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Format(Date, "short date") & Chr(&HD) & Chr(&HA)
                LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Mesero : " & cNomMesero & Chr(&HD) & Chr(&HA)
                LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(1) & Chr(&HD) & Chr(&HA)
                LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Mesa : " & nMesa & Chr(&HD) & Chr(&HA)
                LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(1) & Chr(&HD) & Chr(&HA)
                LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "Cuenta # : " & nLinCta & Chr(&HD) & Chr(&HA)
            Else
                Print #nFreefile, Format(Date, "short date")
                Print #nFreefile, "Mesero : " & cNomMesero
                Print #nFreefile, Space(2)
                Print #nFreefile, "Mesa : " & nMesa
                Print #nFreefile, Space(2)
                Print #nFreefile, "Cuenta # : " & nLinCta
            End If
        End If

    End If
    nTotCta = 0#
Loop
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If HayPrinterLocal Then
    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(nEspacio) & "------------------------------" & Chr(&HD) & Chr(&HA)
Else
    Print #nFreefile, Space(nEspacio) & "------------------------------"
End If

If lParc = 1 Then
    'INFO: HAY PAGO PARCIAL
    MiLen1 = -1
    Milen2 = Len(Format(rsParciales!Valor * (-1), "STANDARD"))
    If HayPrinterLocal Then
        LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & _
            Space(10 - Milen2) & Format(rsParciales!Valor * (-1), "STANDARD") & Chr(&HD) & Chr(&HA)
    Else
        Print #nFreefile, "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & Space(10 - Milen2) & Format(rsParciales!Valor * (-1), "STANDARD")
    End If
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error Resume Next
If HayPrinterLocal Then
    LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(1) & Chr(&HD) & Chr(&HA)
Else
    Print #nFreefile, Space(1)
End If

'========= Milen2 = Len(Format(SubTot, "STANDARD"))

'INFO: 23ENE2011. QUITANDO nTotProp de aqui, ya que no debe estar incluida en este Sub total
'========= If HayPrinterLocal Then
    'LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, _
        "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa + nTotProp), "STANDARD") & Chr(&HD) & Chr(&HA)
    '========= LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa), "STANDARD") & Chr(&HD) & Chr(&HA)
'========= Else
    'Print #nFreefile, "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa + nTotProp), "STANDARD")
    '========= Print #nFreefile, "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa), "STANDARD")
'========= End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MARZO2011
'INFO: TIENE LA OPCION DE DESCUENTO Y HAY DESCUENTO MARCADO
'ENTRA AQUI SI HAY DESCUENTO, PERO COMO EN LA PANTALLA DE MESEROS
'NO SE PUEDE DAR DESCUENTO, ASI QUE LA RUTINA QUE ESTABA AQUI, SE ELIMINA.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'INFO: NO HAY DESCUENTO MARCADO
'========= Call GetGROUPTAX(rsTAX, DescPreCta)
'INFO: ENE 2011. DEJAR DE IMPRIMIR EL DETALLE DE LOS IMPUESTOS
'NADA MAS INCLUIR EL TOTAL DEL IMPUESTO CALCULADO
'========= Do While Not rsTAX.EOF
'=========     Milen2 = Len(Format(rsTAX!TAX, "STANDARD"))
'=========     txtString = "   *ITBMS (" & Format(rsTAX!CON_TAX, "@@") & "%):" & Space(14 - Milen2)
'=========     txtString = txtString & Format(rsTAX!TAX, "STANDARD")
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'SI ESTA ACTIVO IMPRIME EL DETALLE DEL IMPUESTO
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'=========     If bDetalleTax Then
'=========         If HayPrinterLocal Then
'=========             LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, txtString & Chr(&HD) & Chr(&HA)
'=========         Else
'=========             Print #nFreefile, txtString
'=========         End If
'=========     End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
'=========     nTempISC = nTempISC + FormatCurrency(rsTAX!TAX, 2)
'=========     rsTAX.MoveNext
'========= Loop
'========= rsTAX.Close
'========= Set rsTAX = Nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'========= Milen2 = Len(Format(nTempISC, "STANDARD"))
'========= txtString = "   *ITBMS TOTAL:" & Space(14 - Milen2)
'========= txtString = txtString & Format(nTempISC, "STANDARD")
'txtString = txtString & Format(iISCTransaccion, "STANDARD")
'========= If HayPrinterLocal Then
'=========     LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, txtString & Chr(&HD) & Chr(&HA)
'========= Else
'=========     Print #nFreefile, txtString
'========= End If

'========= Milen2 = Len(Format(nSubtotalMesa, "STANDARD"))
'========= If HayPrinterLocal Then
'=========     LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "     Sub-Total :" & Space(14 - Milen2) & _
        Format((nSubtotalMesa - nLDESCUENTO + nTempISC), "STANDARD") & Chr(&HD) & Chr(&HA)
'=========     LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(1) & Chr(&HD) & Chr(&HA)
'========= Else
'=========     Print #nFreefile, "     Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa - nLDESCUENTO + nTempISC), "STANDARD")
'=========     Print #nFreefile, Space(1)
'========= End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'If (OKProp = 1 And nLinCta = 0) Or (OPEN_PROPINA = False And mlincta = 0) Then
'INFO: MAR2011. ELIMINANDO RESTRICCION DE # CUENTA EN LA MESA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'========= If (OKProp = 1) Or (OPEN_PROPINA = False And mlincta = 0) Then
    'SI HAY PROPINA
    'MiPropina = Format(Round(SubTot * 0.1, 1) * 1#, "STANDARD")
'=========     If nMesa <> rs00!MESA_BARRA Then
'=========         MiPropina = RoundToNearest((SBTot - nTempISC) * 0.1, 0.05, 1)
        'MiPropina = RoundToNearest((SBTot - iISCTransaccion) * 0.1, 0.05, 1)
'=========         nProp = MiPropina * 100
'=========         nProp = nProp / 100
'=========         nTotProp = nProp
'=========         NLEN = Len(PROPINA_DESCRIP)
        
'=========         If HayPrinterLocal Then
'=========             LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, PROPINA_DESCRIP & " : " & Space(23 - NLEN) & Format(nProp, "##0.00") & Chr(&HD) & Chr(&HA)
'=========             LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(1) & Chr(&HD) & Chr(&HA)
'=========         Else
'=========             Print #nFreefile, PROPINA_DESCRIP & " : " & Space(23 - NLEN) & Format(nProp, "##0.00")
'=========             Print #nFreefile, Space(2)
'=========         End If
'=========     End If
    'EL SOLOINI DECIDE CUANTO ES LA PROPINA (31/10/2007)
    'OKProp = 0
'========= End If

'========= Milen2 = Len(Format((SubTot + nTotProp - DescPreCta), "CURRENCY"))

'INFO: MAYO 2010
'========= nSUMA = SubTot + nTotProp - DescPreCta

'========= If HayPrinterLocal Then
'=========     LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, "     SUMA      :" & Space(14 - Milen2) & Format((SubTot + nTotProp - DescPreCta), "CURRENCY") & Chr(&HD) & Chr(&HA)
'========= Else
'=========     Print #nFreefile, "     SUMA      :" & Space(14 - Milen2) & Format((SubTot + nTotProp - DescPreCta), "CURRENCY")
'========= End If
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 2 JUNIO 2010
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'========= If HayPrinterLocal Then
'=========     If bPropinaAparte Then
        'IMPRIME LA PROPINA FUERA DEL TOTAL DE LA FACTURA
        'ENTONCES NO AFECTA EL TOTAL DE LA FACTURA
'=========         LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(nEspacio) & Chr(&HD) & Chr(&HA)
        
'=========         NLEN = Len(Space(1) & PROPINA_DESCRIP & Space(1))
'=========         LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(nEspacio) & _
                String((28 - NLEN) / 2, "=") & Space(1) & PROPINA_DESCRIP & Space(1) & String((28 - NLEN) / 2, "=") & Chr(&HD) & Chr(&HA)
                
'=========         For i = 0 To UBound(aPropina)
            'NLEN = Len(Format(aPropina(i) & " % " & " : ", "@@@@@@@@"))
        
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '25 JULIO 2010
            'CORRECCION: CALCULANDO LA PROPINA SOBRE EL SUBTOTAL - DESCUENTO (ANTES DEL IMPUESTO y EXONERACIONES)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'=========             NLEN = Len("( " & Format(aPropina(i), "@@") & " % )")
'=========             LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, _
                "( " & Format(aPropina(i), "@@") & " % )" & Space(10) & _
                Format(Format((RoundToNearest((aPropina(i) / 100) * (nSubtotalMesa - nLDESCUENTO), 0.05, 1)), "##0.00"), "@@@@@@@") & Chr(&HD) & Chr(&HA)
'=========         Next
'=========         LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(nEspacio) & String(28, "=") & Chr(&HD) & Chr(&HA)
'=========     End If
    
'========= End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For i = 1 To 10
    If HayPrinterLocal Then
        LoginMesas.ImpresoraCuentas.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
    Else
        Print #nFreefile, Space(1)
    End If
Next

If HayPrinterLocal Then
    LoginMesas.ImpresoraCuentas.CutPaper 100
Else
    Close #nFreefile
End If

On Error GoTo 0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If nErrCount >= 3 Then
    ShowMsg "POR FAVOR REVISE LA PRE-CUENTA", vbRed, vbYellow
    Resume Next
End If
On Error GoTo 0
Exit Sub

AdmErr:
nErrCount = nErrCount + 1
EscribeLog "Meseros.ImprPreCuenta: (" & Err.Number & ") - " & Err.Description
Milen2 = 10
If nErrCount < 3 Then
    Resume
Else
    Resume Next
End If
'CALL sOLOTRANS("BEGIN")
'msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'CALL sOLOTRANS("COMMIT")
End Sub
Private Sub ImprPreCta()
'INFO: PENDIENTE
Dim sqltext As String
Dim LinTx As String
Dim rsPreCta As ADODB.Recordset
Dim MiMatriz(0, 3) As String
Dim MiLen1, Milen2 As Integer
Dim MiPropina As Single, nProp As Single
Dim rsParciales As ADODB.Recordset
Dim lParc As Integer
Dim nLinCta As Integer
Dim nTotCta As Single
Dim nTotProp As Single
Dim nErrCount As Integer
Dim i As Integer
'MAY/2006
Dim rsTAX As ADODB.Recordset
Dim nSubtotalMesa As Single
Dim nLDESCUENTO As Single
'27/08/2005
Dim HayPrinterLocal As Boolean
Dim nFreefile As Integer
'02/JUN/2006
Dim nTempISC As Single
Dim cErrorText As String
'INFO: DOMICILIO
Dim aDomiInfo() As Variant
'INFO: Muestra INFO de Propina
Dim aPropina As Variant                       'CARGA LOS % DE LA PROPINA
Dim bPropinaAparte As Boolean           'DETERMINA SI IMPRIME LA PROPINA APARTE
Dim nSUMA As Single
'INFO: 23ENE2011
Dim bDetalleTax As Boolean  'DEFINE SI SE IMPRIME C/U DE LOS IMPUESTOS
'INFO: 2DIC2012
Dim txtString As String
Dim arrayMensaje() As String

nSubtotalMesa = GetSubTotalNOTAXFromMesa()

If NOM_PRN_FACTURA = "" Or NOM_PRN_FACTURA = " " Then
    'NO HAY IMPRESORA DE PRE-CUENTAS, IMPRIMIR EN
    'LA DE FACTURACION
    HayPrinterLocal = False
Else
    'INFO: 22SEP2013. REVISION PARA QUE MAS DE UNA APLICACION PUEDA USAR LA IMPRESORA
    'INFO: 28ABRIL2014. REVISION DE SI ES UN SISTEMA CON MULTIPLE TABLETA
    If cHayTableta = "SI" Then
        If Claim_Enable_LocalPrinter Then
            HayPrinterLocal = True
        Else
            HayPrinterLocal = False
        End If
    Else
        HayPrinterLocal = True
    End If
End If

If HayPrinterLocal Then
    If LoginMesas.ImpresoraCuentas.RecNearEnd = True Or LoginMesas.ImpresoraCuentas.RecEmpty = True Then
        ShowMsg "Advertencia de Papel" & vbCrLf & "POR FAVOR REVISE EL PAPEL EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO", vbBlue, vbYellow
    End If
End If

'============================================
'INFO: MAYO 2010
'============================================
aPropina = Split(GetFromINI("Facturacion", "PropinaAparte", App.Path & "\soloini.ini"), ",")

If UBound(aPropina) >= 0 Then
    bPropinaAparte = True
Else
    bPropinaAparte = False
End If
'============================================

'============================================
'INFO: ENERO 2011
'============================================
bDetalleTax = True
If UCase(GetFromINI("Facturacion", "DetalleTax", App.Path & "\soloini.ini")) = "NO" Then
    bDetalleTax = False
Else
    bDetalleTax = True
End If
'============================================

Set rsPreCta = New ADODB.Recordset
Set rsParciales = New ADODB.Recordset

OKCancelar = 0
'INFO: PANTALLA DE PRECUENTA NO ESTA DISPONIBLE EN LAS ESTACIONES DE MESERO
'YA QUE SE IMPRIME DIRECTAMENTE SIN PREGUNTAR NADA
If OPEN_PROPINA = True Then
    ''Opc01.Show 1
End If

nTotProp = 0#

On Error GoTo AdmErr:
'PROPINA_DESCRIP
If OKCancelar = 1 Then OKCancelar = 0: Exit Sub

sqltext = "SELECT MESA,SUM(MONTO) AS VALOR "
sqltext = sqltext & " FROM TMP_PAR_PAGO "
sqltext = sqltext & " WHERE MESA = " & nMesa
sqltext = sqltext & " GROUP BY MESA"

rsParciales.Open sqltext, msConn, adOpenDynamic, adLockOptimistic

'VERIFICA SI HAY PAGOS PARCIALES
If rsParciales.EOF Then lParc = 0 Else lParc = 1

'sqltext = "SELECT MESA,CUENTA,LIN,DESCRIP,CANT,PRECIO "
sqltext = "SELECT MESA,CUENTA,LIN,DESCRIP,CANT,PRECIO, HORA "
sqltext = sqltext & " FROM TMP_TRANS "
sqltext = sqltext & " WHERE MESA = " & nMesa
sqltext = sqltext & " AND CUENTA = " & nCta
sqltext = sqltext & " ORDER BY CUENTA,LIN "

rsPreCta.Open sqltext, msConn, adOpenStatic, adLockReadOnly
If rsPreCta.EOF Then
    ShowMsg "NO HAY PLATOS EN LA MESA. PRE-CUENTA NO SE IMPRIMIRA"
    'INFO: 25/4/2006
    rsParciales.Close
    Set rsParciales = Nothing
    Exit Sub
End If
rsPreCta.MoveFirst
nLinCta = rsPreCta!CUENTA

'INFO: ELIMINANDO LA IMPRESION DE LA PRE-CUENTA
'JUNIO 11 2010
'Call AddPrecuenta2Grid

If HayPrinterLocal Then
    'INFO: Batch processing mode (7AGO2013) / 22SEP2013
    'LoginMesas.ImpresoraCuentas.TransactionPrint PtrSReceipt, PtrTpTransaction
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(LoginMesas.ImpresoraCuentas.Name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    'INFO: CAMBIO SEGUN DGI (NO INCLUIR EMPRESA NI RUC)
    'Print2_OPOS_Dev rs00!descrip
    'Print2_OPOS_Dev rs00!Mensaje
    'Print2_OPOS_Dev rs00!RAZ_SOC
    'INFO: CAMBIO SEGUN DGI (SOLAMENTE LA HORA)
    'Print2_OPOS_Dev "RUC:" & rs00!RUC
    'Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
    Print2_OPOS_Dev Space(2)
    Print2_OPOS_Dev Chr(27) & "|3C" & "NO FISCAL / NO FISCAL"
    Print2_OPOS_Dev Chr(27) & "|3C" & "         PRE-CUENTA"
    'INFO: CAMBIO SEGUN DGI (SOLAMENTE LA HORA)
    'Print2_OPOS_Dev Format(Date, "SHORT DATE") & "     " & Time
    Print2_OPOS_Dev Time
Else
    nFreefile = FreeFile
    'Open DATA_PATH + FACTURA_FILE For Output Shared As #nFreeFile
    'INFO: SI NO HAY IMPRESORA LOCAL, LA PRE-CUENTA HAY QUE IMPRIMIRLA EN LA CAJA
    'INFO: SEP2009
    'Open App.Path & "\" & FACTURA_FILE For Output Shared As #nFreefile
    'INFO: 29MAR2017
    Open App.Path & "\" & npNumCaj & "_" & FACTURA_FILE For Output Shared As #nFreefile
    
    Print #nFreefile, Space(2)
    
    'INFO: CAMBIO SEGUN DGI (NO INCLUIR EMPRESA NI RUC)
    'Print #nFreefile, rs00!descrip
    'Print #nFreefile, rs00!RAZ_SOC
    'INFO: CAMBIO SEGUN DGI (SOLAMENTE LA HORA)
    'Print #nFreefile, "RUC:" & rs00!RUC
    'Print #nFreefile, "SERIAL:" & rs00!SERIAL
    Print #nFreefile, Space(2)
    Print #nFreefile, Chr(27) & "|3C" & "         PRE-CUENTA"
    
    'INFO: CAMBIO SEGUN DGI (SOLAMENTE LA HORA)
    'Print #nFreefile, Format(Date, "SHORT DATE") & "     " & Time
    Print #nFreefile, Time
End If

If bISThisSocios Then
    'INFO: ES RESTAURANTE CON SOCIOS
    Dim rsSocio As ADODB.Recordset
    Set rsSocio = New ADODB.Recordset
    rsSocio.Open "SELECT * FROM SOCIO WHERE MESA = " & nMesa, msConn, adOpenStatic, adLockOptimistic
    If Not rsSocio.EOF Then
        If HayPrinterLocal Then
            Print2_OPOS_Dev Space(2)
            Print2_OPOS_Dev "SOCIO : " & rsSocio!SOCIO_NOMBRE
            Print2_OPOS_Dev Space(2)
        Else
            Print #nFreefile, Space(2)
            Print #nFreefile, "SOCIO : " & rsSocio!SOCIO_NOMBRE
            Print #nFreefile, Space(2)
        End If
    End If
    rsSocio.Close
    Set rsSocio = Nothing
End If

If HayPrinterLocal Then
    
    'INFO: 2DIC2015
    'INFO: CAMBIO SEGUN DGI (ELIMINAR MESERO)
    If bMeseroEnPrecuenta Then Print2_OPOS_Dev "Mesero : " & cNomMesero
    Print2_OPOS_Dev Space(2)
    
    '~~~~~~~~~~~~~~
    'INFO DOMICILIO
    '~~~~~~~~~~~~~~
    If HAS_Domicilio Then
        If nMesa >= nDomicilio Then
            
            If HAS_Domicilio Then
                If MesaAssigned() = "" Then
                    'INFO: SI NO HAY CLIENTE ASIGNADO, ENTONCES NO INTENTA IMPRIMIR LA INFO DEL CLIENTE.
                Else
                    aDomiInfo = GetDomicilioInfo()
        
                    Print2_OPOS_Dev "PEDIDO : " & nMesa & "     " & aDomiInfo(14, 0)
                    Print2_OPOS_Dev "Nombre  : " & aDomiInfo(3, 0) & Space(1) & aDomiInfo(4, 0)
                    If aDomiInfo(6, 0) <> "" Then Print2_OPOS_Dev "Empresa : " & aDomiInfo(6, 0)
                    Print2_OPOS_Dev "Telefono: " & FormatPhone(aDomiInfo(1, 0)) & IIf(aDomiInfo(2, 0) <> "", " (" & aDomiInfo(2, 0) & ")", Space(1))
                    Print2_OPOS_Dev "ZONA: " & aDomiInfo(12, 0)
                    
                    Print2_OPOS_Dev "Dir: " & Left(aDomiInfo(8, 0), 25)
                    If Mid(aDomiInfo(8, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(8, 0), 26, 25)
                    
                    If aDomiInfo(9, 0) <> "" Then Print2_OPOS_Dev "- " & Left(aDomiInfo(9, 0), 25)
                    If Mid(aDomiInfo(9, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(9, 0), 26, 25)
                    
                    If aDomiInfo(10, 0) <> "" Then Print2_OPOS_Dev "- " & Left(aDomiInfo(10, 0), 25)
                    If Mid(aDomiInfo(10, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(10, 0), 26, 25)
                    
                    If Left(aDomiInfo(11, 0), 25) <> "" Then Print2_OPOS_Dev "INFO:" & Left(aDomiInfo(11, 0), 25)
                    If Mid(aDomiInfo(11, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(11, 0), 26, 25)
                    'INFO: 23ENE2011
                    If IsNull(aDomiInfo(15, 0)) Then
                        'DO NOTHING, NO SE HA ASIGNADO MOTORIZADO
                    Else
                        Print2_OPOS_Dev "Motorizado: " & GetMotorizado(CLng(aDomiInfo(15, 0)))
                    End If
                End If
            End If
        Else
            Print2_OPOS_Dev "Mesa : " & nMesa
            Print2_OPOS_Dev Space(2)
        End If
    Else
        Print2_OPOS_Dev "Mesa : " & nMesa
        Print2_OPOS_Dev Space(2)
    End If
Else
    'INFO: CAMBIO SEGUN DGI (ELIMINAR MESERO)
    'Print #nFreefile, "Mesero : " & cNomMesero
    If bMeseroEnPrecuenta Then Print #nFreefile, "Mesero : " & cNomMesero
    Print #nFreefile, Space(2)

    '~~~~~~~~~~~~~~
    'INFO DOMICILIO
    '~~~~~~~~~~~~~~
    If HAS_Domicilio Then
        If nMesa >= nDomicilio Then
            
            aDomiInfo = GetDomicilioInfo()
            
            Print #nFreefile, "PEDIDO : " & nMesa & "     " & aDomiInfo(14, 0)
            Print #nFreefile, "Nombre  : " & aDomiInfo(3, 0) & Space(1) & aDomiInfo(4, 0)
            If aDomiInfo(6, 0) <> "" Then Print #nFreefile, "Empresa : " & aDomiInfo(6, 0)
            Print #nFreefile, "Telefono: " & FormatPhone(aDomiInfo(1, 0)) & IIf(aDomiInfo(2, 0) <> "", " (" & aDomiInfo(2, 0) & ")", Space(1))
            Print #nFreefile, "ZONA: " & aDomiInfo(12, 0)
            
            Print #nFreefile, "Dir: " & Left(aDomiInfo(8, 0), 25)
            If Mid(aDomiInfo(8, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(8, 0), 26, 25)
            
            If aDomiInfo(9, 0) <> "" Then Print #nFreefile, "- " & Left(aDomiInfo(9, 0), 25)
            If Mid(aDomiInfo(9, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(9, 0), 26, 25)
            
            If aDomiInfo(10, 0) <> "" Then Print #nFreefile, "- " & Left(aDomiInfo(10, 0), 25)
            If Mid(aDomiInfo(10, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(10, 0), 26, 25)
            
            If Left(aDomiInfo(11, 0), 25) <> "" Then Print #nFreefile, "INFO:" & Left(aDomiInfo(11, 0), 25)
            If Mid(aDomiInfo(11, 0), 26, 25) <> "" Then Print #nFreefile, "- " & Mid(aDomiInfo(11, 0), 26, 25)
            'INFO: 23ENE2011
            If IsNull(aDomiInfo(15, 0)) Then
                'DO NOTHING, NO SE HA ASIGNADO MOTORIZADO
            Else
                Print #nFreefile, "Motorizado: " & GetMotorizado(CLng(aDomiInfo(15, 0)))
            End If
        Else
            Print #nFreefile, "Mesa : " & nMesa
            Print #nFreefile, Space(2)
        End If
    Else
        'INFO: 05/MAY/2009 - 23ENE2011
        Print #nFreefile, "Mesa : " & nMesa
        Print #nFreefile, Space(2)
    End If
End If

If nLinCta <> 0 Then
    'INFO: HAY CUENTAS SEPARADAS
    If HayPrinterLocal Then
        Print2_OPOS_Dev "Cuenta # : " & nLinCta
    Else
        Print #nFreefile, "Cuenta # : " & nLinCta
    End If
End If

If HayPrinterLocal Then
    'Print2_OPOS_Dev Space(nEspacio) & "------------------------------"
    '23AGO2019
    Print2_OPOS_Dev Chr(27) & "|3C" & "NO FISCAL / NO FISCAL"
    Print2_OPOS_Dev Space(nEspacio) & String(nl_Line, "-")
Else
    Print #nFreefile, Space(nEspacio) & "------------------------------"
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: DETALLE DE LA CUENTA
'~~~~~~~~~~~~~~~~~~~~~~~~~~
Do Until rsPreCta.EOF
    Do Until rsPreCta.EOF
        If HayPrinterLocal Then
            'MiMatriz(0, 0) = FormatTexto(rsPreCta!descrip, GetLargoDescrip())
            '23AGO2019
            'MiMatriz(0, 0) = FormatTexto(rsPreCta!descrip, nl_Descrip)
            'MiMatriz(0, 0) = Format(FormatTexto(rsPreCta!descrip, nl_Descrip), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
            'MiMatriz(0, 0) = FormatTexto(FormatTexto(rsPreCta!descrip, nl_Descrip), 30)
            'MiMatriz(0, 0) = FormatTexto(rsPreCta!descrip, nl_Descrip)
            MiMatriz(0, 0) = FormatTexto(rsPreCta!descrip, nl_Descrip)
        Else
            MiMatriz(0, 0) = FormatTexto(rsPreCta!descrip, 15)
        End If
        '@@@@@@@
        'MiMatriz(0, 1) = Format(rsPreCta!cant, "GENERAL NUMBER")
        MiMatriz(0, 1) = FormatTexto(rsPreCta!CANT, 3)
        'MiMatriz(0, 2) = Format(rsPreCta!precio, "#,###.00")
        'MiMatriz(0, 2) = FormatTexto(Format(rsPreCta!precio, "#,###.00"), 8)
        MiMatriz(0, 2) = Format(rsPreCta!precio, "#,###.00")
        MiLen1 = Len(MiMatriz(0, 1))
        Milen2 = Len(MiMatriz(0, 2))
        LinTx = MiMatriz(0, 0) & Space(5 - MiLen1) & MiMatriz(0, 1) & Space(10 - Milen2) & MiMatriz(0, 2)
        'Debug.Print LinTx
        'Debug.Print Len(MiMatriz(0, 0))
        If HayPrinterLocal Then
            Print2_OPOS_Dev FormatTexto(LinTx, 41)
        Else
            Print #nFreefile, LinTx
            'Print #nFreefile, "                   < " & Format(rsPreCta!Hora, "MEDIUM TIME") & ">"
        End If
        'INFO: ACUMULA SUBTOTAL SIN IMPUESTO
        nTotCta = nTotCta + rsPreCta!precio
        rsPreCta.MoveNext
        If rsPreCta.EOF Then Exit Do
        If nLinCta <> rsPreCta!CUENTA Then
            nLinCta = rsPreCta!CUENTA
            Exit Do
        End If
    Loop

    If nLinCta <> 0 Then
        'INFO: HAY CUENTAS SEPARADAS
        MiLen1 = Len(Format(nTotCta, "STANDARD"))
        nProp = 0#
        
        nTotCta = nTotCta + nProp
        
        If Not rsPreCta.EOF Then
            For i = 1 To 10
                If HayPrinterLocal Then
                    Print2_OPOS_Dev Space(2)
                Else
                    Print #nFreefile, Space(2)
                End If
            Next
            If HayPrinterLocal Then
                LoginMesas.ImpresoraCuentas.CutPaper 100
                Print2_OPOS_Dev Format(Date, "short date")
                'Print2_OPOS_Dev "Mesero : " & cNomMesero
                Print2_OPOS_Dev Space(1)
                Print2_OPOS_Dev "Mesa : " & nMesa
                Print2_OPOS_Dev Space(1)
                Print2_OPOS_Dev "Cuenta # : " & nLinCta
            Else
                Print #nFreefile, Format(Date, "short date")
                'Print #nFreefile, "Mesero : " & cNomMesero
                Print #nFreefile, Space(2)
                Print #nFreefile, "Mesa : " & nMesa
                Print #nFreefile, Space(2)
                Print #nFreefile, "Cuenta # : " & nLinCta
            End If
        End If

    End If
    nTotCta = 0#
Loop
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If HayPrinterLocal Then
    'Print2_OPOS_Dev Space(nEspacio) & "------------------------------"
    '23AGO2019
    Print2_OPOS_Dev Space(nEspacio) & String(nl_Line, "-")
Else
    Print #nFreefile, Space(nEspacio) & "------------------------------"
End If

If lParc = 1 Then
    'INFO: HAY PAGO PARCIAL
    MiLen1 = -1
    Milen2 = Len(Format(rsParciales!Valor * (-1), "STANDARD"))
    If HayPrinterLocal Then
        Print2_OPOS_Dev "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & Space(10 - Milen2) & Format(rsParciales!Valor * (-1), "STANDARD")
    Else
        Print #nFreefile, "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & Space(10 - Milen2) & Format(rsParciales!Valor * (-1), "STANDARD")
    End If
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error Resume Next
If HayPrinterLocal Then
    Print2_OPOS_Dev Space(1)
Else
    Print #nFreefile, Space(1)
End If

Milen2 = Len(Format(SubTot, "STANDARD"))

'INFO: 23ENE2011. QUITANDO nTotProp de aqui, ya que no debe estar incluida en este Sub total
If HayPrinterLocal Then
    'Print2_OPOS_Dev _
        "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa + nTotProp), "STANDARD")
    Print2_OPOS_Dev "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa), "STANDARD")
Else
    'Print #nFreefile, "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa + nTotProp), "STANDARD")
    Print #nFreefile, "   * Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa), "STANDARD")
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MARZO2011
'INFO: TIENE LA OPCION DE DESCUENTO Y HAY DESCUENTO MARCADO
'ENTRA AQUI SI HAY DESCUENTO, PERO COMO EN LA PANTALLA DE MESEROS
'NO SE PUEDE DAR DESCUENTO, ASI QUE LA RUTINA QUE ESTABA AQUI, SE ELIMINA.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'INFO: NO HAY DESCUENTO MARCADO
Call GetGROUPTAX(rsTAX, DescPreCta)
'INFO: ENE 2011. DEJAR DE IMPRIMIR EL DETALLE DE LOS IMPUESTOS
'NADA MAS INCLUIR EL TOTAL DEL IMPUESTO CALCULADO
Do While Not rsTAX.EOF
    Milen2 = Len(Format(rsTAX!TAX, "STANDARD"))
    txtString = "   *ITBMS (" & Format(rsTAX!CON_TAX, "@@") & "%):" & Space(14 - Milen2) & Format(rsTAX!TAX, "STANDARD")
    'txtString = txtString & Format(rsTAX!TAX, "STANDARD")
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'SI ESTA ACTIVO IMPRIME EL DETALLE DEL IMPUESTO
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If bDetalleTax Then
        If HayPrinterLocal Then
            Print2_OPOS_Dev txtString
        Else
            Print #nFreefile, txtString
        End If
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    nTempISC = nTempISC + FormatCurrency(rsTAX!TAX, 2)
    rsTAX.MoveNext
Loop
rsTAX.Close
Set rsTAX = Nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Milen2 = Len(Format(nTempISC, "STANDARD"))
txtString = "   *ITBMS TOTAL:" & Space(14 - Milen2) & Format(nTempISC, "STANDARD")
'txtString = txtString & Format(nTempISC, "STANDARD")
'txtString = txtString & Format(iISCTransaccion, "STANDARD")
If HayPrinterLocal Then
    Print2_OPOS_Dev txtString
Else
    Print #nFreefile, txtString
End If

Milen2 = Len(Format(nSubtotalMesa, "STANDARD"))
If HayPrinterLocal Then
    Print2_OPOS_Dev "     Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa - nLDESCUENTO + nTempISC), "STANDARD")
    Print2_OPOS_Dev Space(1)
Else
    Print #nFreefile, "     Sub-Total :" & Space(14 - Milen2) & Format((nSubtotalMesa - nLDESCUENTO + nTempISC), "STANDARD")
    Print #nFreefile, Space(1)
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'If (OKProp = 1 And nLinCta = 0) Or (OPEN_PROPINA = False And mlincta = 0) Then
'INFO: MAR2011. ELIMINANDO RESTRICCION DE # CUENTA EN LA MESA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If (OKProp = 1) Or (OPEN_PROPINA = False And mlincta = 0) Then
    'SI HAY PROPINA
    'MiPropina = Format(Round(SubTot * 0.1, 1) * 1#, "STANDARD")
    If nMesa <> rs00!MESA_BARRA Then
        MiPropina = RoundToNearest((SBTot - nTempISC) * 0.1, 0.05, 1)
        'MiPropina = RoundToNearest((SBTot - iISCTransaccion) * 0.1, 0.05, 1)
        nProp = MiPropina * 100
        nProp = nProp / 100
        nTotProp = nProp
        NLEN = Len(PROPINA_DESCRIP)
        
        If HayPrinterLocal Then
            Print2_OPOS_Dev PROPINA_DESCRIP & " : " & Space(21 - NLEN) & Format(nProp, "##0.00")
            Print2_OPOS_Dev Space(1)
        Else
            Print #nFreefile, PROPINA_DESCRIP & " : " & Space(21 - NLEN) & Format(nProp, "##0.00")
            Print #nFreefile, Space(2)
        End If
    End If
    'EL SOLOINI DECIDE CUANTO ES LA PROPINA (31/10/2007)
    'OKProp = 0
End If

'Milen2 = Len(Format((SubTot + nTotProp - DescPreCta), "CURRENCY"))

'INFO: MAYO 2010
'nSUMA = SubTot + nTotProp - DescPreCta
'INFO: 3FEB2014. USANDO LOS TOTALES IMPRESOS EN VEZ DE LOS CALCULADOS DEL LVVIEW
nSUMA = nSubtotalMesa - nLDESCUENTO + nTempISC + nTotProp - DescPreCta
Milen2 = Len(Format((nSUMA), "CURRENCY"))
'nsuma2 = nSubtotalMesa - nLDESCUENTO + nTempISC + nTotProp - DescPreCta

If HayPrinterLocal Then
    'INFO: 3FEB2014. USANDO LOS TOTALES IMPRESOS EN VEZ DE LOS CALCULADOS DEL LVVIEW
    'Print2_OPOS_Dev "     SUMA      :" & Space(14 - Milen2) & Format((SubTot + nTotProp - DescPreCta), "CURRENCY")
    Print2_OPOS_Dev "     SUMA      :" & Space(14 - Milen2) & Format((nSUMA), "CURRENCY")
    Print2_OPOS_Dev Space(1)
    Print2_OPOS_Dev Chr(27) & "|3C" & "NO FISCAL / NO FISCAL"
Else
    'INFO: 3FEB2014. USANDO LOS TOTALES IMPRESOS EN VEZ DE LOS CALCULADOS DEL LVVIEW
    'Print #nFreefile, "     SUMA      :" & Space(14 - Milen2) & Format((SubTot + nTotProp - DescPreCta), "CURRENCY")
    Print #nFreefile, "     SUMA      :" & Space(14 - Milen2) & Format((nSUMA), "CURRENCY")
End If
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 2 JUNIO 2010
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If HayPrinterLocal Then
    If bPropinaAparte Then
        'IMPRIME LA PROPINA FUERA DEL TOTAL DE LA FACTURA
        'ENTONCES NO AFECTA EL TOTAL DE LA FACTURA
        Print2_OPOS_Dev Space(nEspacio)
        
        NLEN = Len(Space(1) & PROPINA_DESCRIP & Space(1))
        Print2_OPOS_Dev Space(nEspacio) & _
                String((28 - NLEN) / 2, "=") & Space(1) & PROPINA_DESCRIP & Space(1) & String((28 - NLEN) / 2, "=")
                
        For i = 0 To UBound(aPropina)
            'NLEN = Len(Format(aPropina(i) & " % " & " : ", "@@@@@@@@"))
        
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '25 JULIO 2010
            'CORRECCION: CALCULANDO LA PROPINA SOBRE EL SUBTOTAL - DESCUENTO (ANTES DEL IMPUESTO y EXONERACIONES)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            NLEN = Len("( " & Format(aPropina(i), "@@") & " % )")
            Print2_OPOS_Dev _
                "( " & Format(aPropina(i), "@@") & " % )" & Space(10) & _
                Format(Format((RoundToNearest((aPropina(i) / 100) * (nSubtotalMesa - nLDESCUENTO), 0.05, 1)), "##0.00"), "@@@@@@@")
        Next
        Print2_OPOS_Dev Space(nEspacio) & String(28, "=")
    End If
Else
    'INFO: 18SEP2013. HACIA FALTA CALCULAR CUANDO NO HAY IMPRESORA LOCAL
    'ANTES COMO NO SE IMPRIMIA EN NINGUN LADO NO ERA NECESARIO.
    If bPropinaAparte Then
        'IMPRIME LA PROPINA FUERA DEL TOTAL DE LA FACTURA
        'ENTONCES NO AFECTA EL TOTAL DE LA FACTURA
        Print #nFreefile, Space(nEspacio)
        
        NLEN = Len(Space(1) & PROPINA_DESCRIP & Space(1))
        Print #nFreefile, Space(nEspacio) & _
                String((28 - NLEN) / 2, "=") & Space(1) & PROPINA_DESCRIP & Space(1) & String((28 - NLEN) / 2, "=")
                
        For i = 0 To UBound(aPropina)
            'NLEN = Len(Format(aPropina(i) & " % " & " : ", "@@@@@@@@"))
        
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '25 JULIO 2010
            'CORRECCION: CALCULANDO LA PROPINA SOBRE EL SUBTOTAL - DESCUENTO (ANTES DEL IMPUESTO y EXONERACIONES)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            NLEN = Len("( " & Format(aPropina(i), "@@") & " % )")
            Print #nFreefile, _
                "( " & Format(aPropina(i), "@@") & " % )" & Space(10) & _
                Format(Format((RoundToNearest((aPropina(i) / 100) * (nSubtotalMesa - nLDESCUENTO), 0.05, 1)), "##0.00"), "@@@@@@@")
        Next
        Print #nFreefile, Space(nEspacio) & String(28, "=")
    End If
End If



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 11SEP2015
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
arrayMensaje = Split(rs00!Mensaje, ",")
If HayPrinterLocal Then
    If UBound(arrayMensaje) = -1 Then
        'NO HAY MENSAJES
    Else
        Print2_OPOS_Dev Space(1)
        Print2_OPOS_Dev Space(1)
        Print2_OPOS_Dev String(30, "~")
        Print2_OPOS_Dev String(30, "~")
        For i = 0 To UBound(arrayMensaje)
            Print2_OPOS_Dev Format(arrayMensaje(i), "@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        Next
        Print2_OPOS_Dev String(30, "~")
        Print2_OPOS_Dev String(30, "~")
    End If
Else
    If UBound(arrayMensaje) = -1 Then
        'NO HAY MENSAJES
    Else
        Print #nFreefile, Space(1)
        Print #nFreefile, Space(1)
        Print #nFreefile, String(30, "~")
        Print #nFreefile, String(30, "~")
        For i = 0 To UBound(arrayMensaje)
            Print #nFreefile, Format(arrayMensaje(i), "@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        Next
        Print #nFreefile, String(30, "~")
        Print #nFreefile, String(30, "~")
    End If
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For i = 1 To 10
    If HayPrinterLocal Then
        Print2_OPOS_Dev Space(2)
    Else
        Print #nFreefile, Space(1)
    End If
Next

If HayPrinterLocal Then

    'INFO: 7MAR2019. IMPRESORA BEMATECH EN ESTACION DE MESEROS
    'INFO: 10sep2021
    If NOM_PRN_FACTURA = "w" Or NOM_PRN_FACTURA = "w1" Then
        Printer.EndDoc
    Else
        
        'INFO: CORTE DE PAPEL DEBE ESTAR ANTES DEL FIN DE TRANSACTION PRINT
        LoginMesas.ImpresoraCuentas.CutPaper 100
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'INFO: print all the buffer data. and exit the batch processing mode. (7AGO2013)
        'LoginMesas.ImpresoraCuentas.TransactionPrint PtrSReceipt, PtrTpNormal
        Call OPOSTransactionPrint(LoginMesas.ImpresoraCuentas.Name, "END")
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
        'INFO: 22SEP2013
        'INFO: 28ABRIL2014
        If cHayTableta = "SI" Then
            Call OPOSRelease_CLOSE_LocalPrinter
        Else
        End If
    End If
Else
    Close #nFreefile
End If


On Error GoTo 0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If nErrCount >= 3 Then
    ShowMsg "POR FAVOR REVISE LA PRE-CUENTA", vbRed, vbYellow
    Resume Next
End If

If Not HayPrinterLocal Then PantallaPreCuenta.Show 1
On Error GoTo 0
Exit Sub

AdmErr:
nErrCount = nErrCount + 1
EscribeLog "Meseros.ImprPreCuenta: (" & Err.Number & ") - " & Err.Description
Milen2 = 10
If nErrCount < 3 Then
    Resume
Else
    Resume Next
End If
'CALL sOLOTRANS("BEGIN")
'msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'CALL sOLOTRANS("COMMIT")
End Sub
Private Function GetLargoDescrip() As Integer
'IMPRIME EL MEJOR LARGO SEGUN LA IMPRESORA
'INFO: 14OCT2013
GetLargoDescrip = 15
Select Case OPOS_DevName
    'INFO: 28SNRIL2014
    'INFO: 28ABRIL2014. PARA LAS T-20 MAS NUEVAS."T20-42C" y "T-20II-42C"
    Case "SRP-350plus", "MP4200TH", "TM-T20E", "TM-T20U", "TM-T20-42CU", "TM-T20-42CE", "TM-T20II-42CU", "TM-T20II-42CE"
        GetLargoDescrip = 25
    'INFO: 21ABRIL2014. UPDATE PARA BIXOLON OPOS NUEVO
    Case "LR3000", "TM-U220B", "TM-U200B", "SRP270", "SRP270P", "TM-U220B", "SEMOPOS.SO.SERIAL.POSPrinter", _
           "SRP-275", "SRP-275P", "SRP-270", "SRP-270P"
        GetLargoDescrip = 15
    Case "TM-U950P", "TM-U950"
        GetLargoDescrip = 15
    Case Else
        'DEFAULT SI ES UNA IMPRESORA DESCONOCIDA LO BAJA A 15
        'GetLargoDescrip = 25
        GetLargoDescrip = 15
End Select

End Function
Private Sub AddOpenDeptItem()
'Agrega registros a TMP_TRANS desde un Departamento Abierto
'REEMPLAZO DE ####.## por #0.00 CUANDO Sea Necesario
Dim SOLO_FECHA As String
Dim CadenaSql As String
Dim nTempoSingle As Single
'INFO: REVISADA ABRIL/2006
'INFO: AGREGANDO GetENCRYPTEDINI PARA leer el impuesto del soloini.ini
Dim cDeptoAbiertoItem As String
Dim cSQL As String

ValOpenDept = 0
TXT_OPEN_DEPT = ""

ActHost.Show 1

'If TXT_OPEN_DEPT = "" Then Exit Sub
On Error GoTo ErrAdmAddOpenDeptItem:
cDeptoAbiertoItem = RegRead("HKCU\Software\SoloSoftware\SoloMix\DeptoAbiertoItem")
If TXT_OPEN_DEPT = "" And cDeptoAbiertoItem = "" Then
    'INFO: VALIDA TAMBIEN QUE NO EXISTA DESCRIPCION ADICIONAL CON PRECIO 0.00
    On Error GoTo 0
    Exit Sub
End If

CajLin = CajLin + 1
SOLO_FECHA = Format(Date, "YYYYMMDD")
    
'INFO: POR DEFAULT CON_TAX DEBE ESTAR EN 0 (NO ES CIERTO)
'INFO: SE PONE EN VALOR DEFAULT DEL TAX EN SOLOINI.INI
CadenaSql = "INSERT INTO TMP_TRANS "
CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,"
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'INFO: ACTUALIZACION DE AREAS
                '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
'CadenaSql = CadenaSql & "PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX) VALUES ("
CadenaSql = CadenaSql & "PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX, AREA) VALUES ("
CadenaSql = CadenaSql & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & ",'"
'=================================================
'INFO: AGREGANDO LA DESCRIPCION DEL PRODUCTO DEL DEPARTAMENTO ABIERTO
' 29SEP2012
'=================================================

If cDeptoAbiertoItem <> "" Then
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\DeptoAbiertoItem", ""
    cDeptoAbiertoItem = "DA.." & cDeptoAbiertoItem
    CadenaSql = CadenaSql & cDeptoAbiertoItem & "'," & nMulti & "," & rs02!codigo & "," & 0 & "," & 0 & ","
    'INFO: 29SEP2012 /  Update 4NOV2012 (agrega precio)
    EscribeLog "Ventas. Producto Temporal: " & cDeptoAbiertoItem & ", Precio: " & Format((ValOpenDept / nMulti) / 100, "#0.00")
Else
    CadenaSql = CadenaSql & rs02!CORTO & TXT_OPEN_DEPT & "'," & nMulti & "," & rs02!codigo & "," & 0 & "," & 0 & ","
End If
'=================================================
'=================================================
CadenaSql = CadenaSql & Format((ValOpenDept / nMulti) / 100, "#0.00") & "," & Format(ValOpenDept / 100, "#0.00") & ",'"

Select Case TXT_OPEN_DEPT
    Case " RESTAURANTE"
        CadenaSql = CadenaSql & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & nCta & ",FALSE," & 1 & ","
    Case " BARRA"
        'INFO EL NUMERO CORRECTO DE LA IMPRESORA DE BARRA ES 2
        CadenaSql = CadenaSql & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & nCta & ",FALSE," & 2 & ","
    Case Else
        'SI ES OTRA COSA VA A LA IMPRESORA DE FACTURACION (26/10/2007)
        'CadenaSql = CadenaSql & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & nCta & ",FALSE," & 0 & ","
        'VALIDA EL CLICK QUE SE HIZO (RESTAURANTE, BARRA u OTRO 10DIC2012) CON nDeptoAbiertoSeleccionado
        CadenaSql = CadenaSql & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & nCta & ",FALSE," & nDeptoAbiertoSeleccionado & ","
End Select

CadenaSql = CadenaSql & GetFromINI("Facturacion", "PorcentajeImpuesto", App.Path & "\soloini.ini") & "," & nArea & ")"

Call SOLOTrans("BEGIN")
msConn.Execute CadenaSql
Call SOLOTrans("COMMIT")

'If CajLin = 1 Then msConn.Execute "UPDATE Mesas SET ocupada = TRUE,MESERO_ACTUAL = " & rs!numero & " WHERE numero = " & nMesa
If CajLin = 1 Then
    'msConn.Execute "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & nMesero & " WHERE numero = " & nMesa
    Call SQL_Update(False, "AddOpenDeptItem", "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & nMesero & " WHERE numero = " & nMesa)
    
End If

cSQL = "SELECT a.lin,a.descrip,a.cant,"
cSQL = cSQL & " format(precio_unit,'##0.00') as mPrecio_unit,"
cSQL = cSQL & " format(precio,'##0.00') as mPrecio,"
cSQL = cSQL & " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, "
cSQL = cSQL & " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, "
cSQL = cSQL & " a.caja "
cSQL = cSQL & " FROM tmp_trans as a "
cSQL = cSQL & " WHERE a.mesa = " & nMesa
''If bCuenta = True Then
    cSQL = cSQL & " AND a.CUENTA = " & nCta
''End If
cSQL = cSQL & " ORDER BY a.lin"

rs07.Open cSQL, msConn, adOpenStatic, adLockOptimistic

Set PlatosMesa.DataSource = rs07
SetupPantalla

nLineas = PlatosMesa.Rows - 1

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
    " WHERE MESA = " & nMesa & _
    " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
    
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)

'NEW SBTot = Format(SubTot, "STANDARD")
''''On Error Resume Next
''''nTempoSingle = (rs07!precio * iISC)
''''SubTot = FormatCURRENCY((SubTot + nTempoSingle), 2)
''''iISCTransaccion = rs07!precio * iISC
''''SBTot = Format(SubTot, "STANDARD")
''''On Error GoTo 0

Call ActualizaSUBTOTAL

rs07.Close
If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad
StatBar.Panels(4) = 1

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!Valor, "STANDARD") & Chr(9) & Format(rsParciales!Valor, "STANDARD")
    SubTot = Format(SubTot - rsParciales!Valor, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

'CALL sOLOTRANS("BEGIN")
'msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'CALL sOLOTRANS("COMMIT")
On Error GoTo 0
Exit Sub

ErrAdmAddOpenDeptItem:
MsgBox "Debe ESCRIBIR un Producto en Departamento Abierto", vbInformation, BoxTit

End Sub
Private Sub DescProducto()
'PROCEDIMIENTO PARA DAR DESCUENTO A UN PRODUCTO
'INFO: ACTUALIZADO ABRIL/2006
Dim MiDesc As Single
Dim nDescImpre As Single
Dim rsParciales As Recordset
Dim lParc As Integer
Dim sqltext As String

If PlatosMesa.Rows = 0 Then
    ShowMsg " No hay nada Marcado "
    Exit Sub
End If

If nCantidad > MAX_DESCUENTO Then
    ShowMsg "ES IMPOSIBLE DAR ESE DESCUENTO. INTENTE DAR UN PORCENTAJE MAS BAJO", vbRed, vbYellow
    Clear_Click
    Exit Sub
End If

'nCantidad es el valor del Cuadro de Numeros de la Derecha Abajo
If nCantidad > 1 Then
    MiDesc = Format(nCantidad / 100, "STANDARD")
Else
    'nDesc01 es el Descuento Marcado
    MiDesc = Format(nDesc01 / 100, "STANDARD")
End If

MiDesc = Format(MiDesc, "STANDARD")

Dim rsFixTmpTrans As New Recordset
Dim rsGetMaxLin As New Recordset
Dim nMaxLin As Integer
Dim txto As String
Dim nTempoSingle As Single

If nNLinSel <> 0 Then   'PREGUNTA SI HIZO CLICK A PLATOSMESAS
    txto = "SELECT * FROM tmp_trans "
    txto = txto & " WHERE mesa = " & nMesa & " AND lin = " & nNLinSel
Else
    txto = "SELECT MAX(LIN) AS MAX_LIN "
    txto = txto & " FROM TMP_TRANS "
    txto = txto & " WHERE MESA = " & nMesa
    
    rsGetMaxLin.Open txto, msConn, adOpenStatic, adLockOptimistic
    
    If (PlatosMesa.Rows - 1) >= 1 Then
        PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
    End If
    PlatosMesa.col = 0
    PlatosMesa.row = (PlatosMesa.Rows - 1)

    txto = "SELECT * FROM tmp_trans "
    txto = txto & " WHERE MESA = " & nMesa & " AND LIN = " & Val(PlatosMesa.Text)
    nMaxLin = rsGetMaxLin!MAX_LIN
    rsGetMaxLin.Close
    Set rsGetMaxLin = Nothing
    'txto = "SELECT * FROM tmp_trans " & _
        " WHERE MESA = " & nMesa & " AND LIN = " & Val(PlatosMesa.Text)
End If

rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

If rsFixTmpTrans.EOF = True Then
    rsFixTmpTrans.Close
    ShowMsg "Por Favor SELECCIONE un Producto"
    nCantidad = 1
    StatBar.Panels(4) = 1
    Exit Sub
End If

If rsFixTmpTrans!CANT < 0 Then
    'Si la Cantidad es 0 entonces...
    ShowMsg "NO puede dar DESCUENTO a este Producto", vbRed, vbYellow
    rsFixTmpTrans.Close
    nCantidad = 1
    StatBar.Panels(4) = 1
    Exit Sub
End If
    
If Mid(rsFixTmpTrans!descrip, 1, 9) = "DESCUENTO" Then
    ShowMsg "NO puede dar DESCUENTO a un Descuento", vbRed, vbYellow
    rsFixTmpTrans.Close
    nCantidad = 1
    StatBar.Panels(4) = 1
    Exit Sub
End If
    
If Mid(rsFixTmpTrans!TIPO, 1, 1) = "B" Then
    ShowMsg "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!TIPO, 5, 2)), vbRed, vbYellow
    rsFixTmpTrans.Close
    nCantidad = 1
    StatBar.Panels(4) = 1
    Exit Sub
End If
    
nCtaLinAnul = rsFixTmpTrans!CUENTA
CajLin = CajLin + 1

'------------REVISION DE PAGOS PARCIALES-------------------
Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR " & _
            " FROM TMP_PAR_PAGO " & _
            " WHERE MESA = " & nMesa & _
            " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1
'--------------------------------

Dim nTestDesc As Integer
nDescImpre = Format(MiDesc * rsFixTmpTrans!precio * (-1), "STANDARD")
nTestDesc = Val(Mid(nDescImpre, Len(nDescImpre) + 1, 1))
    
'Proceso que quita los centavos del Descuento y los redondea al mas bajo
'y Asigna su valor a nDescImpre
If nTestDesc = 0 Or nTestDesc = 5 Then
ElseIf nTestDesc < 5 Then
    nDescImpre = nDescImpre + (nTestDesc / 100)
ElseIf nTestDesc > 5 And nTestDesc <= 9 Then
    nDescImpre = nDescImpre + ((nTestDesc - 5) / 100)
End If
    
Dim SOLO_FECHA As String
SOLO_FECHA = Format(Date, "YYYYMMDD")

If nNLinSel <> 0 Then
    CadenaSql = "INSERT INTO TMP_TRANS "
    CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,"
    CadenaSql = CadenaSql & "DESCRIP,CANT,DEPTO,"
    CadenaSql = CadenaSql & "PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,"
    'CadenaSql = CadenaSql & "HORA,TIPO,DESCUENTO,CUENTA,CON_TAX) VALUES ("
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    'INFO: ACTUALIZACION DE AREAS
                    '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    CadenaSql = CadenaSql & "HORA,TIPO,DESCUENTO,CUENTA,CON_TAX, AREA) VALUES ("
    CadenaSql = CadenaSql & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & ","
    CadenaSql = CadenaSql & "'DESCUENTO : " & Format(MiDesc, "#.00") & "%'" & "," & 1 & "," & rsFixTmpTrans!DEPTO & ","
    CadenaSql = CadenaSql & rsFixTmpTrans!PLU & "," & rsFixTmpTrans!envase & "," & nDescImpre & "," & nDescImpre & ","
    CadenaSql = CadenaSql & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
    CadenaSql = CadenaSql & ",'DC-" & nNLinSel & "'," & MiDesc & "," & nCtaLinAnul & ","
    'ABRIL/2006 = CON_TAX
    CadenaSql = CadenaSql & rsFixTmpTrans!CON_TAX & "," & nArea & ")"

    sqltext = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str((CajLin))
    sqltext = sqltext & "' WHERE MESA = " & nMesa
    sqltext = sqltext & " AND LIN = " & nNLinSel
Else
    CadenaSql = "INSERT INTO TMP_TRANS "
    CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,"
    CadenaSql = CadenaSql & "DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    'INFO: ACTUALIZACION DE AREAS
                    '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'CadenaSql = CadenaSql & "PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,CON_TAX) "
    CadenaSql = CadenaSql & "PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,CON_TAX, AREA) "
    CadenaSql = CadenaSql & " VALUES ("
    CadenaSql = CadenaSql & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & nMaxLin + 1 & ","
    CadenaSql = CadenaSql & "'DESCUENTO : " & Format(MiDesc, "#.00") & "%'" & "," & 1 & "," & rsFixTmpTrans!DEPTO & "," & rsFixTmpTrans!PLU & ","
    CadenaSql = CadenaSql & rsFixTmpTrans!envase & "," & nDescImpre & "," & nDescImpre & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
    
    '***************************************
    'INFO: 25 JULIO 2010
    'CadenaSql = CadenaSql & ",'DC-" & Val(PlatosMesa.Text) & "'," & MiDesc & "," & nCtaLinAnul & ","
    CadenaSql = CadenaSql & ",'DC-" & nMaxLin & "'," & MiDesc & "," & nCtaLinAnul & ","
    '***************************************
    
    '24/8/2005 = CON_TAX
    CadenaSql = CadenaSql & rsFixTmpTrans!CON_TAX & "," & nArea & ")"
        
    sqltext = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str(Val(PlatosMesa.Text))
    sqltext = sqltext & "' WHERE MESA = " & nMesa
    sqltext = sqltext & "  AND LIN = " & nMaxLin
    CajLin = (nMaxLin + 1)
End If
    
Call SOLOTrans("BEGIN")
msConn.Execute CadenaSql
msConn.Execute sqltext
Call SOLOTrans("COMMIT")

'''''''''''msConnLoc.BeginTrans
'''''''''''msConnLoc.Execute CadenaSql
'''''''''''msConnLoc.Execute sqltext
'''''''''''msConnLoc.CommitTrans

''''rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
''''    " format(precio_unit,'##0.00') as mPrecio_unit," & _
''''    " format(precio,'##0.00') as mPrecio," & _
''''    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
''''    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
''''    " a.caja " & _
''''    " FROM tmp_trans as a " & _
''''    " WHERE a.mesa = " & nMesa & _
''''    " AND A.CUENTA = " & nCta & _
''''    " ORDER BY a.lin ", msConn, adOpenStatic, adLockOptimistic
''''
''''Set PlatosMesa.DataSource = rs07
''''SetupPantalla
    
Call OpenTMP_TRANS(True)
Call SetupPantalla

nLineas = PlatosMesa.Rows - 1

If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
    
rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & _
    " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly

SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
'NEW SBTot = Format(SubTot, "STANDARD")

''''On Error Resume Next
''''nTempoSingle = (rs07!precio * iISC)
''''SubTot = FormatCURRENCY((SubTot + nTempoSingle), 2)
''''iISCTransaccion = rs07!precio * iISC
''''SBTot = Format(SubTot, "STANDARD")
''''On Error GoTo 0

Call ActualizaSUBTOTAL

rs07.Close
rsFixTmpTrans.Close

nCantidad = 1: nPase = 0
Text1(2) = nCantidad
StatBar.Panels(4) = 1

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
    "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
    Format(rsParciales!Valor, "STANDARD") & Chr(9) & Format(rsParciales!Valor, "STANDARD")
    SubTot = Format(SubTot - rsParciales!Valor, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If
End Sub
Private Sub MuestraPLU_del_Envase(nElEnvase As Long)
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim iTam As Integer
Dim sqltext As String
'INFO: QUEDA IGUAL MAY/2006

'Busca PLUS x Enase Seleccionado
nGlobEnv = nElEnvase
sqltext = "SELECT a.depto,a.contenedor,b.codigo,b.descrip,b.corto,c.precio "
sqltext = sqltext & " FROM CONTEND_01 as a,PLU as b, CONTEND_02 as c "
sqltext = sqltext & " WHERE a.DEPTO = " & ElDepto
sqltext = sqltext & " AND a.contenedor = " & nElEnvase
sqltext = sqltext & " AND a.depto = b.depto "
sqltext = sqltext & " AND b.codigo = c.codigo "
'INFO: 6MAY2011
'QUE NADA MAS MUESTRE LOS DEL TAG DE DISPONIBLE
sqltext = sqltext & " AND B.DISPONIBLE = TRUE "
sqltext = sqltext & " AND c.contenedor = " & nElEnvase
sqltext = sqltext & " ORDER BY b.CORTO "
rs08.Open sqltext, msConn, adOpenStatic, adLockOptimistic
iTam = 0

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs08.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs08!descrip
        cmdPlus(numplu).Tag = rs08!codigo
        cmdPlus(numplu).ToolTipText = "Precio : " & Format(rs08!precio, "CURRENCY")
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs08!descrip
        cmdPlus(numplu).Tag = rs08!codigo
        cmdPlus(numplu).ToolTipText = "Precio : " & Format(rs08!precio, "CURRENCY")
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    'INFO: CAMBIO A 800X600
    ''''MiLeft = MiLeft + 1900
    MiLeft = MiLeft + 2400
    'If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
    'INFO: 1024 ---> If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
    If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
        'INFO: CAMBIO A 800X600
        ''''MiTop = MiTop + 460
        'MiTop = MiTop + 600
        'MiTop = MiTop + 670
        MiTop = MiTop + 865
        MiLeft = 0
    End If
    'If numplu = 18 Then Exit Do
    'INFO: 1024 ---> If numplu = 24 Then Exit Do
    If numplu = 24 Then Exit Do
    rs08.MoveNext
Loop
rs08.Close
End Sub
Private Sub SetupPantalla()
Dim i As Integer
'Formato de la Pantalla de Facturacion
'INFO: QUEDA IGUAL ABRIL/2006
On Error GoTo ErrAdm:
With PlatosMesa
'INFO: 22ENE2013. ELIMINANDO EL LOOP DE COLUMNAS CERO
'    For i = 0 To 5
'        DoEvents
'        .ColWidth(i) = 0
'    Next
    '.ColWidth(0) = 400       'LINEA
    .ColWidth(0) = 550       'LINEA
    '.ColWidth(1) = 3360     'PRODUCTO
    .ColWidth(1) = 3460     'PRODUCTO
    .ColWidth(2) = 410       'CANTIDAD
    '.ColWidth(3) = 800       'PRECIO UNIT
    .ColWidth(3) = 960        'PRECIO UNIT
    .ColWidth(4) = 1130     'PRECIO
    '.ColWidth(18) = 2600     'PRECIO
    'MsgBox .Cols
    'INFO: CAMBIO A 800X600
    ''''.ColWidth(0) = 300: .ColWidth(1) = 2800: .ColWidth(2) = 400
    ''''.ColWidth(3) = 650: .ColWidth(4) = 800:
    .ColAlignmentFixed(3) = flexAlignRightCenter
    .ColAlignmentFixed(4) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
End With
On Error GoTo 0
Exit Sub

ErrAdm:
    MsgBox "SetupPantalla. " & Err.Number & Space(1) & Err.Description, vbCritical
End Sub
Private Sub QuitarPLUS()
'INFO: QUEDA IGUAL ABRIL/2006
Dim nNum As Integer, lNum As Integer

nNum = rs03.RecordCount

'For lNum = 1 To 17
'INFO: 1024 ---> For lNum = 1 To 23
'For lNum = 1 To 23
For lNum = 1 To 23
    cmdPlus(1).Caption = ""
    If Not IsObject(cmdPlus(lNum)) Then
        Load cmdPlus(lNum)
    End If
    cmdPlus(lNum).Visible = False
Next
cmdPlus(0).Visible = True
End Sub
Private Sub MuestraPLU(ElDepto As Integer)
'VIENE DE HACER CLICK A LOS DEPARTAMENTOS
'INFO: ACTUALIZADO ABRIL/2006
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim iTam As Integer
Dim cSQL As String

On Error GoTo ErrAdm:
'Muestra los productos a Vender
Set rs03 = New Recordset
'Set rs04 = New Recordset
rs04.Close
'Busca PLUS del Depto
'21/08/2005
cSQL = "SELECT codigo,depto,descrip,corto,precio1,envases,IMPRESORA,CON_TAX "
cSQL = cSQL & " FROM PLU "
cSQL = cSQL & " WHERE depto = " & ElDepto
cSQL = cSQL & " AND DISPONIBLE = TRUE "
cSQL = cSQL & " ORDER BY DESCRIP"
rs03.Open cSQL, msConn, adOpenStatic, adLockReadOnly
'Busca Envases del Departamento
cSQL = "SELECT a.depto,a.contenedor,b.descrip "
cSQL = cSQL & " FROM contend_01 as a,contened as b "
cSQL = cSQL & " WHERE a.DEPTO = " & ElDepto & " AND "
cSQL = cSQL & " a.contenedor = b.contenedor "
cSQL = cSQL & " ORDER BY a.depto,a.contenedor"
rs04.Open cSQL, msConn, adOpenStatic, adLockOptimistic
iTam = 0
'Prepara los Envases del Departamento
For iTam = 0 To 3
    cmdEnvases(iTam).Enabled = True
    cmdEnvases(iTam).BackColor = &HC0C0C0
    cmdEnvases(iTam).Caption = ""
Next

iTam = 0

'If Not rs04.EOF Then FlashControl Frame2(4)
If rs04.EOF Then Frame2(4).BackColor = &H8000000F Else Frame2(4).BackColor = &HFFFF&

Do Until rs04.EOF
    cmdEnvases(iTam).Caption = rs04!descrip
    cmdEnvases(iTam).Tag = rs04!contenedor
    iTam = iTam + 1
    rs04.MoveNext
Loop

For iTam = 0 To 3
    If cmdEnvases(iTam).Caption = "" Then
        cmdEnvases(iTam).Enabled = False
    End If
Next

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0
'Si No hay productos, quitar los que estan visibles
If rs03.EOF Then
    Dim lNum As Integer
    cmdPlus(0).Tag = ""
    For lNum = 0 To 23
        cmdPlus(0).Caption = ""
        If Not IsObject(cmdPlus(lNum)) Then
            Load cmdPlus(lNum)
        End If
        cmdPlus(lNum).Visible = False
    Next
    cmdPlus(0).Visible = True
    rs02.MoveFirst
    rs02.Find "CODIGO = " & ElDepto
    If Not rs02.EOF Then
        If rs02!ABIERTO = True Then
            'MsgBox "DEPARTAMENTO ABIERTO", vbCritical, BoxTit
            If nMesa = 0 Or nMesero = 0 Then
                MsgBox "Antes Debe Seleccionar una Mesa y su Mesero", vbCritical, BoxTit
                Exit Sub
            End If
            AddOpenDeptItem
        End If
        
    End If
    
    Exit Sub
End If

'SI HAY PRODUCTOS EN EL DEPARTAMENTO, LOS MUESTRO
Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        cmdPlus(numplu).ToolTipText = "Precio : " & Format(rs03!precio1, "CURRENCY")
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        cmdPlus(numplu).ToolTipText = "Precio : " & Format(rs03!precio1, "CURRENCY")
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    'INFO: CAMBIO A 800X600
    ''''MiLeft = MiLeft + 1900
    MiLeft = MiLeft + 2400
    'If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
    'INFO: 1024 ---> If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
    If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
        'INFO: CAMBIO A 800X600
        ''''MiTop = MiTop + 460
        'MiTop = MiTop + 600
        'MiTop = MiTop + 670
        MiTop = MiTop + 865
        MiLeft = 0
    End If
    'If numplu = 18 Then Exit Do
    'INFO: 1024 ---> If numplu = 24 Then Exit Do
    If numplu = 24 Then Exit Do
    rs03.MoveNext
Loop
On Error GoTo 0
Exit Sub

ErrAdm:
    EscribeLog ("POSIBLE PROBLEMA CON LA TABLA PLU, CAMPO: CON_TAX, REVISAR BASE DE DATOS")
    EscribeLog (Err.Number & " - " & Err.Description)
    MsgBox "POSIBLE PROBLEMA CON LA TABLA PLU, CAMPO: CON_TAX, REVISAR BASE DE DATOS" & vbCrLf & _
        Err.Number & " - " & Err.Description & vbCrLf & _
        "SALGA DEL PROGRAMA Y CONTACTE A SOLO SOFTWARE DEVELOPMENT", vbCritical, BoxTit
    'Resume
End Sub
Private Sub Quita_Subrallado(var As Integer)
'INFO: ACTUALIZADO ABRIL/2006
Dim i As Integer

i = 0

For i = 0 To cmdDepto.Count - 1
    cmdDepto(i).BackColor = &H8000000F
    'INFO: ABRIL 2010
    'cmdDepto(i).BackColor = &HC0C0C0
Next
If var <> 67 Then
    cmdDepto(var).BackColor = &HFFFF80
End If
End Sub
Private Sub QuitarDeptos()
'INFO: ACTUALIZADO ABRIL/2006
Dim nNum As Integer

For nNum = 1 To cmdDepto.Count - 1
    cmdDepto(1).Caption = ""
    cmdDepto(nNum).Visible = False
Next
End Sub

Private Sub Clear_Click()
nPase = 0
nCantidad = 1
Text1(2) = nCantidad
StatBar.Panels(4) = 1
End Sub

Private Sub cmdAcomp_Click(Index As Integer)
'INFO: ACTUALIZADO ABRIL/2006
Dim SOLO_FECHA As String
Dim nTempoSingle As Single

If cmdAcomp(Index).Caption = "" Then Exit Sub

CajLin = CajLin + 1
SOLO_FECHA = Format(Date, "YYYYMMDD")

CadenaSql = "INSERT INTO TMP_TRANS "
CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,"
CadenaSql = CadenaSql & "ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    'INFO: ACTUALIZACION DE AREAS
                    '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CadenaSql = CadenaSql & "IMPRESO,IMPRESORA,CON_TAX) VALUES ("
CadenaSql = CadenaSql & "IMPRESO,IMPRESORA,CON_TAX, AREA) VALUES ("
CadenaSql = CadenaSql & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & ","
CadenaSql = CadenaSql & -1 & "," & CajLin & ",' @@ " & cmdAcomp(Index).Caption & "',"
CadenaSql = CadenaSql & 1 & "," & cmdAcomp(Index).Tag & "," & 0 & "," & 0 & ","
CadenaSql = CadenaSql & 0 & "," & 0 & ",'" & SOLO_FECHA & "','" & Time & "','  ',"
CadenaSql = CadenaSql & 0# & "," & nCta & ",FALSE," & nSeleccionCocina & ","
'CadenaSql = CadenaSql & GetENCRYPTEDINI("Facturacion", "PorcentajeImpuesto", App.path & "\soloini.ini") & ")"
CadenaSql = CadenaSql & GetFromINI("Facturacion", "PorcentajeImpuesto", App.Path & "\soloini.ini") & "," & nArea & ")"

Call SOLOTrans("BEGIN")
msConn.Execute CadenaSql
Call SOLOTrans("COMMIT")

If CajLin = 1 Then
    'msConn.Execute "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & rs!numero & " WHERE numero = " & nMesa
    Call SQL_Update(False, "cmdAcomp_Click", "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & rs!numero & " WHERE numero = " & nMesa)
End If

''''rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
''''    " format(precio_unit,'##0.00') as mPrecio_unit," & _
''''    " format(precio,'##0.00') as mPrecio," & _
''''    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
''''    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
''''    " a.caja " & _
''''    " FROM tmp_trans as a " & _
''''    " WHERE a.mesa = " & nMesa & _
''''    " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic
''''
''''Set PlatosMesa.DataSource = rs07

'INFO: 22ENE2013 / ERROR 3705 / CUANDO LA RED ESTA LENTA
If rs07.State = adStateOpen Then rs07.Close
'Debug.Print Time & " - cmdAcomp_Click.OpenTMP_TRANS"
Call OpenTMP_TRANS(True)
'Debug.Print Time & " - cmdAcomp_Click.OpenTMP_TRANS.Return"

'DoEvents       'INFO: 08AGO2016. REMOVIENDO DOEVENTS
'INFO: PONIENDO DATASOURCE = RS07 AQUI PARA QUE SE VEA LOS DATOS EN EL GRID
'YA QUE OPENTMP_TRANS ACTUALIZA EL RS07
On Error Resume Next
Set PlatosMesa.DataSource = rs07
On Error GoTo 0

SetupPantalla

nLineas = PlatosMesa.Rows - 1

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
    " WHERE MESA = " & nMesa & _
    " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

'INFO: 22ENE2012
'ERROR EN LA RED LENTA  3704 - La operación no está permitida si el objeto está cerrado.
'On Error Resume Next
If rs07.State = adStateOpen Then rs07.Close
'rs07.Close
'On Error GoTo 0
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)

'''''NEW SBTot = Format(SubTot, "STANDARD")
''''On Error Resume Next
''''nTempoSingle = (rs07!precio * iISC)
''''SubTot = FormatCURRENCY((SubTot + nTempoSingle), 2)
''''iISCTransaccion = rs07!precio * iISC
''''SBTot = Format(SubTot, "STANDARD")
''''On Error GoTo 0

Call ActualizaSUBTOTAL

rs07.Close
If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad
StatBar.Panels(4) = 1
'PUEDE SELECCIONAR MAS DE UN ACOMPAÑANTE
'POR ESO FRAME2(2)=ENABLED
'Frame2(2).Enabled = False

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!Valor, "STANDARD") & Chr(9) & Format(rsParciales!Valor, "STANDARD")
    SubTot = Format(SubTot - rsParciales!Valor, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

End Sub

Private Sub cmdAcomp_GotFocus(Index As Integer)
'INFO: 28NOV2013
cmdAcomp(Index).BackColor = &HFFFF00
End Sub

Private Sub cmdAcomp_LostFocus(Index As Integer)
'INFO: 28NOV2013
'cmdAcomp(Index).BackColor = &H8000000F
cmdAcomp(Index).BackColor = &HC0C0C0
End Sub

Private Sub cmdCtas_Click()
'INFO: ACTUALIZADO ABRIL/2006
Dim rsParciales As Recordset
Dim rsMaxLin As New ADODB.Recordset
Dim cSQL As String
Dim nTempoSingle As Single

On Error GoTo ErrAdm:

If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    
Set rsParciales = New Recordset
rsParciales.Open "SELECT CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO " & _
        " FROM TMP_PAR_PAGO " & _
        " WHERE MESA = " & nMesa, msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then
    rsParciales.Close
    cSQL = "SELECT DISTINCT TIPO FROM TMP_TRANS "
    cSQL = cSQL & " WHERE MESA = " & nMesa
    cSQL = cSQL & " AND CUENTA = 0 "
    rsParciales.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    If rsParciales.RecordCount > 1 Then
        ShowMsg "NO ES POSIBLE ASIGNAR CUENTAS. ESTA MESA YA TIENE CORRECCIONES, ANULACIONES o DESCUENTOS " & vbCrLf & _
            "SE LE SUGIERE ABRIR UNA NUEVA MESA", vbRed
        rsParciales.Close
        Set rsParciales = Nothing
        Exit Sub
    End If
    Set rsParciales = Nothing
    
    If nCta = 0 Then
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' BLOCK PARA PRAIA
        If cAllowSeparar = "SI" Then
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If ShowMsg("¿ DESEA ABRIR CUENTAS SEPARADAS PARA ESTA MESA ?", , , vbYesNo) = vbYes Then
            
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'INFO: 27/SEP/2010
                'REVISA SI HAY ALGO PENDIENTE QUE NO SE HA MARCADO Y LO ENVIA A CHEF, CUANDO PRESIONA MESAS
                Call ENVIAR_PLATOS("NO MOSTRAR MENSAJE DE PLATOS PENDIENTES, NI PASAR A MESAS")
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            
                FacCtaPlato.Show 1      'PANTALLA DE CUENTAS
                lbCuenta = nCta
            
                Call OpenTMP_TRANS(True)
                
                On Error Resume Next
                Set PlatosMesa.DataSource = rs07
                On Error GoTo 0
                
                SetupPantalla
                
                rs07.Close
                rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
                      " WHERE a.mesa = " & nMesa & _
                      " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
                SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
                
                Call ActualizaSUBTOTAL
            
                rs07.Close
                If (PlatosMesa.Rows - 1) >= 1 Then
                    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
                End If
                nCantidad = 1: nPase = 0
                nNLinSel = 0
                StatBar.Panels(4) = 1
            End If
        Else
            Exit Sub
        End If
    Else
    
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 27/SEP/2010
        'REVISA SI HAY ALGO PENDIENTE QUE NO SE HA MARCADO Y LO ENVIA A CHEF, CUANDO PRESIONA MESAS
        Call ENVIAR_PLATOS("NO MOSTRAR MENSAJE DE PLATOS PENDIENTES, NI PASAR A MESAS")
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        FacCtaPlato.Show 1      'PANTALLA DE CUENTAS
        lbCuenta = nCta
    
        Call OpenTMP_TRANS(True)
        
        On Error Resume Next
        Set PlatosMesa.DataSource = rs07
        On Error GoTo 0
        
        SetupPantalla
        
        rs07.Close
        rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
              " WHERE a.mesa = " & nMesa & _
              " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
        SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
        
        Call ActualizaSUBTOTAL
    
        rs07.Close
        If (PlatosMesa.Rows - 1) >= 1 Then
            PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
        End If
        nCantidad = 1: nPase = 0
        nNLinSel = 0
        StatBar.Panels(4) = 1
    End If
Else
    MsgBox "ESTA MESA NO SE PUEDE TRABAJAR POR CUENTAS, YA QUE TIENE PAGOS PARCIALES", vbExclamation, BoxTit
    Exit Sub
End If
On Error GoTo 0
Exit Sub

ErrAdm:
MsgBox Err.Number & " ->>>> " & Err.Description, vbCritical, BoxTit
Resume Next
End Sub
Private Sub cmdDepto_Click(Index As Integer)
'INFO: REVISADO ABRIL/2006
Dim iAcoCnt As Integer
Frame2(1).Caption = "MENU " & cmdDepto(Index).Caption
Quita_Subrallado (Index)

'=================================
'TRATAMIENTO DE ACOMPAÑANTES
'=================================
iAcoCnt = 0
If Frame2(2).Enabled = True Then
    For iAcoCnt = 0 To 3
        cmdAcomp(iAcoCnt).Caption = ""
        cmdAcomp(iAcoCnt).Tag = 0
    Next
    Frame2(2).Enabled = False
End If
'---------------------------------
TextEnv = ""
'=================================
'INFO: AL DESACTIVAR EL PAGEUP(cmdUpAcompa), SE CORRIGE EL DEPARTAMENTO ZERO
'30/NOV/2004
cmdRestoAco.Enabled = False
cmdUpAcompa.Enabled = False
'=================================
'FIN DE TRATAMIENTO DE ACOMPAÑANTES
'=================================
QuitarPLUS
ElDepto = Arreg_Deptos(Index)
MuestraPLU (Arreg_Deptos(Index))
nGlobEnv = 0
nNLinSel = 0
End Sub

Private Sub cmdEnvases_Click(Index As Integer)
'INFO: REVISADO ABRIL/2006
'cmdEnvases(Index).Tag
TextEnv = "-" + cmdEnvases(Index).Caption
QuitarPLUS
For i = 0 To 3
    cmdEnvases(i).BackColor = &HC0C0C0
Next
If cmdEnvases(Index).BackColor = &HC0C0C0 Then
    cmdEnvases(Index).BackColor = &HFFFF00
Else
    cmdEnvases(Index).BackColor = &HC0C0C0
End If
MuestraPLU_del_Envase (cmdEnvases(Index).Tag)
End Sub

Private Sub cmdNota_Click()
Dim cNotas As String
'INFO: DOMICILIO
cNotas = InputBox("ENVIAR INFORMACION DE PLATOS (50 LETRAS o MENOS) ", "NOTAS DEL PLATO")
If cNotas <> "" Then
    Call UpdateDOMINotas(cNotas)
End If
End Sub

Private Sub cmdPlus_GotFocus(Index As Integer)
cmdPlus(Index).BackColor = &HFFFF00
End Sub
Private Sub cmdPlus_LostFocus(Index As Integer)
'INFO: ABRIL 2010 (CORRECCION DEL COLOR DE FONDO)
'cmdPlus(Index).BackColor = &H00C0C0C0&
cmdPlus(Index).BackColor = &HC0C0C0
End Sub

Private Sub cmdRestoAco_Click()
'INFO: MOSTRAR EL RESTO DE LOS ACOMPAÑANTES
'INFO: REVISADO ABRIL/2006
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
rsTmpAco.Bookmark = nAcoBookMark
'SE MUEVE AL PROXIMO
rsTmpAco.MoveNext
Do Until rsTmpAco.EOF
    If iAcom = 0 Then
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Caption = rsTmpAco!descrip
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        'NUEVO CODIGO PARA SALIR DESPUES DE 4 o CUANDO RsTmpAco=EOF
        nAcoBookMark = rsTmpAco.Bookmark
        rsTmpAco.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    Else
        cmdAcomp(iAcom).Visible = True
        'cmdAcomp(iAcom).Top = nAcoTop + 540
        cmdAcomp(iAcom).Top = nAcoTop + 660
        cmdAcomp(iAcom).Caption = rsTmpAco!descrip
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        'nAcoTop = nAcoTop + 540
        nAcoTop = nAcoTop + 660
'NUEVO CODIGO PARA SALIR DESPUES DE 4 o CUANDO RsTmpAco=EOF
        nAcoBookMark = rsTmpAco.Bookmark
        rsTmpAco.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    End If
Loop
'cmdRestoAco.Enabled = False
On Error GoTo 0
Exit Sub

ErrAdm:
If Err.Number = 340 Then
    MsgBox "Este plato tiene demasiados acompañantes" & vbCrLf & "El maximo de acompañantes es 48", vbInformation, "Muchos Acompañantes"
Else
    Resume Next
End If
End Sub
Private Sub cmdSalir_Click()
If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
Set rsTmpAco = Nothing

'INFO: 3/SEP/2007
'REVISA SI HAY ALGO PENDIENTE QUE NO SE HA MARCADO Y LO ENVIA A CHEF, CUANDO PRESIONA SALIR
Call ENVIAR_PLATOS("NO MOSTRAR MENSAJE DE PLATOS PENDIENTES, NI PASAR A MESAS")

'INFO MAYO 2010: EL TEXTO DEL ENVASE NO SE ESTABA LIMPIANDO SI MARCABAN UN DEPARTAMENTO QUE
'TUVIERA ENVASES PERO EL MESERO NO MARCABA NADA,
'ASI CUANDO ENTRABA EL PROXIMO MESEROS, ESTE VALOR QUEDABA GRAVADO Y LO PONIA AL PROXIMO ITEM QUE SE MARCARA.

TextEnv = ""

Debug.Print "cmdSalir_Click Lineas: " & CajLin & " - CUENTA: " & nCta & " - MESA: " & nMesa
Call EvalMesaUpdate(CajLin, nCta, nMesa)

'StatMesa nMesa, 0
StatMesa nMesa, vbLibre, "PLU.cmdSalir"
Unload Me
'' --- > LoginMesas.Visible = True

LoginMesas.Show
End Sub
Private Sub PrintConsolidadoEnCaja()
'INFO: SI LA OPCION ESTA MARCADA EN SOLOINI.INI
'BAJO LA SECCION DE MESEROS, Y EL VALOR DE
'ConsolidadoEnCaja ES pereza, ENTONCES EL SISTEMA
'IMPRIME EL CONSOLIDADO EN LA IMPRESORA DE CAJA
'INFO: REVISADO ABRIL/2006
'+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=
'INFO: IMPRIMIENDO EN BASE DE DATOS, NO EN ARCHIVO DE TEXTO
'23/OCT/2006

Dim cSQL As String
Dim rsConsolidado As ADODB.Recordset
Dim nFreefile As Integer
Dim cConsecutivo As String
Dim rsReturn As ADODB.Recordset
Dim nHEADMesa As Integer
Dim cHEADMesero As String
Dim iIndicador As Byte

Set rsReturn = New ADODB.Recordset

On Error GoTo ErrAdm:
Set rsConsolidado = New ADODB.Recordset
cSQL = "SELECT A.MESERO,A.MESA,A.LIN,A.DESCRIP,A.CANT,"
cSQL = cSQL & " A.IMPRESO , A.IMPRESORA, "
cSQL = cSQL & " (B.NOMBRE + ' ' + B.APELLIDO) AS NOMBRE"
cSQL = cSQL & " FROM TMP_TRANS AS A, MESEROS AS B "
cSQL = cSQL & " WHERE A.MESA = " & nMesa
cSQL = cSQL & " AND A.MESERO = B.NUMERO "
cSQL = cSQL & " AND A.IMPRESO = FALSE "
cSQL = cSQL & " ORDER BY A.LIN"

rsConsolidado.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If rsConsolidado.EOF Then
    rsConsolidado.Close
    Set rsConsolidado = Nothing
    Exit Sub
End If

cConsecutivo = GetPrinterCounter(2)     'IMPRESORA DE FACTURACION
nHEADMesa = rsConsolidado!MESA
cHEADMesero = Left(rsConsolidado!nombre, 30)

msPED.BeginTrans
'INFO: DETALLE DEL PEDIDO
Do Until rsConsolidado.EOF
    cSQL = "INSERT INTO PEDIDO_DETALLE (IMPRESORA_ORDEN, LIN, CANT, DESCRIPCION) "
    cSQL = cSQL & " VALUES ('"
    cSQL = cSQL & "FACTURA_" & cConsecutivo & "'," & rsConsolidado!LIN & ","
    If Mid(LTrim(rsConsolidado!descrip), 1, 2) = "@@" Then
        cSQL = cSQL & rsConsolidado!CANT & ",'" & Space(3) & Mid(rsConsolidado!descrip, 1, 26) & "')"
    Else
        cSQL = cSQL & rsConsolidado!CANT & ",'" & Mid(rsConsolidado!descrip, 1, 26) & "')"
    End If
    'EscribeLog cSQL
    msPED.Execute cSQL
    rsConsolidado.MoveNext
Loop
msPED.CommitTrans

'INFO: ENCABEZADO PEDIDO
msPED.BeginTrans
If cConsecutivo <> "" Then
    cSQL = "INSERT INTO PEDIDO_MAIN (IMPRESORA_ORDEN, FECHA, HORA, NUM_ORDEN, MESA, MESERO, IS_PRINTED, IMPRESORA) "
    cSQL = cSQL & " VALUES ('"
    cSQL = cSQL & "FACTURA_" & cConsecutivo & "','"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM:SS") & "','"
    cSQL = cSQL & cConsecutivo & "'," & nHEADMesa & ",'"
    cSQL = cSQL & cHEADMesero & "'," & False & ",0)"
    'EscribeLog cSQL
    msPED.Execute cSQL
    Set rsReturn = Nothing
End If
msPED.CommitTrans

rsConsolidado.Close
Set rsConsolidado = Nothing
On Error GoTo 0
Exit Sub

ErrAdm:
    EscribeLog ("Error en Consolidado desde MESERO hacia Impresora de SOLOMIX: " & Err.Description & " - " & cSQL)
    'EscribeLog ("Error en Consolidado desde MESERO hacia Impresora de SOLOMIX")
    Sleep 150
    If iIndicador < 2 Then
        iIndicador = iIndicador + 1
        Resume
    Else
        EscribeLog ("Error en Consolidado desde MESERO hacia Impresora de SOLOMIX: " & Err.Description & " - " & cSQL)
        MsgBox "Error en Consolidado desde MESERO hacia Impresora de SOLOMIX" & " - " & Err.Number & " - " & Err.Description & vbCrLf & cSQL
    End If
    Set rsConsolidado = Nothing
End Sub
Private Sub cmdSlip_Click()
Call ENVIAR_PLATOS("MOSTRAR MENSAJE")
End Sub

Private Function ENVIAR_PLATOS(cShowMessage As String) As Boolean
'INFO: (28/MAR/2007)
'INFO: ACTUALIZADO ABRIL/2006
'INFO: ACTUALIZADO 26FEB2011. PARA QUE ENVIE LOS PRODUCTOS CORRECTAMENTE A CADA IMPRESORA
Dim rsCocina As New ADODB.Recordset
Dim nFlag As Boolean
Dim nFreefile As Integer
Dim bSeImprimioEnCocina As Boolean
Dim cSQL As String
Dim nImpresora As Integer
Dim nSelectedPrinter As Integer
Dim iIndicador As Byte
Dim jIndicador As Byte
Dim cConsecutivo As String
Dim nHEADMesa As Integer
Dim cHEADMesero As String

Dim ooError As Variant, ooErrorMDB As Variant
Dim ooDescrip As String

If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&

nFlag = False

'If UCase(GetENCRYPTEDINI("Meseros", "ConsolidadoEnCaja", App.path & "\soloini.ini")) = "PEREZA" Then
If UCase(GetFromINI("Meseros", "ConsolidadoEnCaja", App.Path & "\soloini.ini")) = "PEREZA" Then
    Call PrintConsolidadoEnCaja
End If

'PRIMERO SELECCIONO COCINA (COCINA_01, COCINA_02 y COCINA_03)
'nImpresora = Val(GetENCRYPTEDINI("Facturacion", "TotalImpresorasCocina", App.path & "\soloini.ini"))
'nImpresora = Val(GetFromINI("Facturacion", "TotalImpresorasCocina", App.Path & "\soloini.ini"))
'For i = 1 To nImpresora
'ProximaCocina:

On Error GoTo ErrAdm:

cSQL = "SELECT A.MESERO,A.MESA,A.LIN,A.DESCRIP,A.CANT,"
cSQL = cSQL & " A.IMPRESO , A.IMPRESORA, "
cSQL = cSQL & " (B.NOMBRE + ' ' + B.APELLIDO) AS NOMBRE"
cSQL = cSQL & " FROM TMP_TRANS AS A, MESEROS AS B "
cSQL = cSQL & " WHERE A.MESA = " & nMesa
cSQL = cSQL & " AND A.MESERO = B.NUMERO "
cSQL = cSQL & " AND A.IMPRESO = FALSE "
cSQL = cSQL & " AND A.IMPRESORA IN (1,3,4,5,6,7) "
cSQL = cSQL & " ORDER BY A.LIN, A.IMPRESORA "
'======================================================
    
rsCocina.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsCocina.EOF Then
    'MsgBox "NO HAY ARTICULOS PENDIENTES PARA ENVIAR A LA COCINA o LA BARRA", vbInformation, BoxTit
    nFlag = True
    rsCocina.Close
    GoTo ChequeaBarra:
End If

rsCocina.MoveFirst
nImpresora = rsCocina!IMPRESORA
nHEADMesa = rsCocina!MESA
cHEADMesero = Left(rsCocina!nombre, 30)

Do While Not rsCocina.EOF
    cConsecutivo = GetPrinterCounter(nImpresora + 2)
    
    msPED.BeginTrans
    Do While nImpresora = rsCocina!IMPRESORA
        jIndicador = 1
        cSQL = "INSERT INTO PEDIDO_DETALLE (IMPRESORA_ORDEN, LIN, CANT, DESCRIPCION) "
        cSQL = cSQL & " VALUES ('"
        cSQL = cSQL & "COCINA" & nImpresora & "_" & cConsecutivo & "'," & rsCocina!LIN & ","
        If Mid(LTrim(rsCocina!descrip), 1, 2) = "@@" Then
            cSQL = cSQL & rsCocina!CANT & ",'" & Space(3) & Mid(rsCocina!descrip, 1, 26) & "')"
        Else
            cSQL = cSQL & rsCocina!CANT & ",'" & Mid(rsCocina!descrip, 1, 26) & "')"
        End If
        msPED.Execute cSQL
        rsCocina.MoveNext
        If rsCocina.EOF Then Exit Do
    Loop
    
    msPED.CommitTrans
    
    jIndicador = 2
    
    cSQL = "INSERT INTO PEDIDO_MAIN (IMPRESORA_ORDEN, FECHA, HORA, NUM_ORDEN, MESA, MESERO, IS_PRINTED, IMPRESORA) "
    cSQL = cSQL & " VALUES ('"
    cSQL = cSQL & "COCINA" & nImpresora & "_" & cConsecutivo & "','"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM:SS") & "','"
    cSQL = cSQL & cConsecutivo & "'," & nHEADMesa & ",'"
    cSQL = cSQL & cHEADMesero & "'," & False & "," & nImpresora & ")"
    
    msPED.BeginTrans
    msPED.Execute cSQL
    msPED.CommitTrans
    
    If rsCocina.EOF Then Exit Do
    nImpresora = rsCocina!IMPRESORA
Loop

rsCocina.Close
bSeImprimioEnCocina = True

'=======================
'DESPUES SELECCIONO LA BARRA
ChequeaBarra:
'=======================
cSQL = "SELECT A.MESERO,A.MESA,A.LIN,A.DESCRIP,A.CANT,A.IMPRESO,A.IMPRESORA, "
cSQL = cSQL & " (B.NOMBRE + ',' + B.APELLIDO) AS NOMBRE"
'INFO: 31MAY2017. PARA IMPRESION EN MAQUINA DE PRECUENTA
cSQL = cSQL & " , A.IMPRESO "
cSQL = cSQL & " FROM TMP_TRANS AS A, MESEROS AS B "
cSQL = cSQL & " WHERE A.MESA = " & nMesa
cSQL = cSQL & " AND A.MESERO = B.NUMERO "
cSQL = cSQL & " AND A.IMPRESO = FALSE "
'cSQL = cSQL & " AND A.IMPRESORA = 2 "
'INFO: UPDATE 7ABR2018.   cSQL = cSQL & " AND A.IMPRESORA IN (0,2,4) "
cSQL = cSQL & " AND A.IMPRESORA IN (0,2) "
cSQL = cSQL & " ORDER BY A.LIN,A.IMPRESORA "
rsCocina.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsCocina.EOF Then
    rsCocina.Close
    Set rsCocina = Nothing
    If nFlag = True Then
        If bSeImprimioEnCocina = False Then
            If cShowMessage = "MOSTRAR MENSAJE" Then
                ShowMsg "NO HAY ARTICULOS PENDIENTES PARA ENVIAR A LA COCINA o LA BARRA", vbWhite
            End If
            
            'INFO: MAYO 2010
            Call SOLOTrans("BEGIN")
                msConn.Execute "UPDATE TMP_TRANS SET IMPRESO = TRUE WHERE MESA = " & nMesa
            Call SOLOTrans("COMMIT")
            
            Exit Function
        End If
    End If
    GoTo MarcaImpresos:
End If

nImpresora = rsCocina!IMPRESORA
nHEADMesa = rsCocina!MESA
cHEADMesero = Left(rsCocina!nombre, 30)

'Print #nFreeFile, "SOLICITUD DE BEBIDAS"
'INFO: CORRECION DE ERROR EN PEDIDOS DE FACTURACION Y BAR, BAR2
'26FEB2011

Do While Not rsCocina.EOF
    cConsecutivo = GetPrinterCounter(nImpresora + 2) '4 ES EL DEFAULT PARA BAR
    
    'INFO: 31MAY2017. RED LION MULTICENTRO. PARA QUE LLEVEN EL TICKET Y LE DESPACHEN EL PEDIDO
    'INFO: 20FEB2018. PARA RESOLVER EL CONFLICTO DE LAS IMPRESORAS, CUANDO HAY
    'PRECUENTA, BAR, COCINA, COCINA2, COCINA3
    'If PedidoBarEnLocalPrinter = "SI" Then
    'If PedidoBarEnLocalPrinter <> "" Then
    '    Exit Do
    'End If
    Select Case PedidoBarEnLocalPrinter
        Case "", "NO"
        Case "SI"
            Exit Do
    End Select
    
    msPED.BeginTrans
    Do While nImpresora = rsCocina!IMPRESORA
        jIndicador = 3
        cSQL = "INSERT INTO PEDIDO_DETALLE (IMPRESORA_ORDEN, LIN, CANT, DESCRIPCION) "
        cSQL = cSQL & " VALUES ('"
        cSQL = cSQL & "BAR" & nImpresora & "_" & cConsecutivo & "'," & rsCocina!LIN & ","
        If Mid(LTrim(rsCocina!descrip), 1, 2) = "@@" Then
            cSQL = cSQL & rsCocina!CANT & ",'" & Space(3) & Mid(rsCocina!descrip, 1, 26) & "')"
        Else
            cSQL = cSQL & rsCocina!CANT & ",'" & Mid(rsCocina!descrip, 1, 26) & "')"
        End If
        msPED.Execute cSQL
        rsCocina.MoveNext
        If rsCocina.EOF Then Exit Do
    Loop
    msPED.CommitTrans

    jIndicador = 4
    msPED.BeginTrans
    cSQL = "INSERT INTO PEDIDO_MAIN (IMPRESORA_ORDEN, FECHA, HORA, NUM_ORDEN, MESA, MESERO, IS_PRINTED, IMPRESORA) "
    cSQL = cSQL & " VALUES ('"
    cSQL = cSQL & "BAR" & nImpresora & "_" & cConsecutivo & "','"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM:SS") & "','"
    cSQL = cSQL & cConsecutivo & "'," & nHEADMesa & ",'"
    cSQL = cSQL & cHEADMesero & "'," & False & "," & nImpresora & ")"
    msPED.Execute cSQL
    msPED.CommitTrans
    
    If rsCocina.EOF Then Exit Do
    nImpresora = rsCocina!IMPRESORA
Loop

'INFO: 31MAY2017
If PedidoBarEnLocalPrinter = "SI" Then
    Call PrintPedidoBarEnPrinterLocal(rsCocina, cConsecutivo)
End If


rsCocina.Close
Set rsCocina = Nothing
'Seleccion_Impresora_Default

MarcaImpresos:
jIndicador = 5
Call SOLOTrans("BEGIN")
msConn.Execute "UPDATE TMP_TRANS SET IMPRESO = TRUE WHERE MESA = " & nMesa
Call SOLOTrans("COMMIT")

On Error GoTo 0
'Call BuscaMesa(False)
If cShowMessage = "MOSTRAR MENSAJE" Then
    'StatMesa nMesa, 0
    StatMesa nMesa, vbLibre, "PLU.ENVIAR_PLATOS"
    If bAutoLogin Then
        Call BuscaMesa(False)
    Else
        'INFO MAYO 2010: EL TEXTO DEL ENVASE NO SE ESTABA LIMPIANDO SI MARCABAN UN DEPARTAMENTO QUE
        'TUVIERA ENVASES PERO EL MESERO NO MARCABA NADA,
        'ASI CUANDO ENTRABA EL PROXIMO MESEROS, ESTE VALOR QUEDABA GRAVADO Y LO PONIA AL PROXIMO ITEM QUE SE MARCARA.
        TextEnv = ""
        Unload PLU
        Unload Mesas
        
        '' --- > LoginMesas.Visible = True
    End If
End If
Exit Function

ErrAdm:

ooError = Err.Number
ooErrorMDB = msPED.Errors(0).NativeError
ooDescrip = Err.Description


If Err.Number = -2147467259 Then
    'info ERROR AL HACER EL INSERT.
    'PUEDE SER QUE HAY UN PEDIDO ANTERIOR EN MESASPED CON EL MISMO IMPRESORA_ORDEN
    Select Case jIndicador
        Case 1
            OLD_EscribeLog ("Error ENVIAR_PLATOS PEDIDO_DETALLE (cocina) - " & ooDescrip)
            OLD_EscribeLog ("COMANDO SQL: " & cSQL)
        Case 2
            OLD_EscribeLog ("Error ENVIAR_PLATOS PEDIDO_MAIN (cocina) - " & ooDescrip)
            OLD_EscribeLog ("COMANDO SQL: " & cSQL)
        Case 3
            OLD_EscribeLog ("Error ENVIAR_PLATOS PEDIDO_DETALLE (BAR) - " & ooDescrip)
            OLD_EscribeLog ("COMANDO SQL: " & cSQL)
        Case 4
            OLD_EscribeLog ("Error ENVIAR_PLATOS PEDIDO_MAIN  (BAR) - " & ooDescrip)
            OLD_EscribeLog ("COMANDO SQL: " & cSQL)
        Case 5
            OLD_EscribeLog ("Error ENVIAR_PLATOS TMP_TRANS SET IMPRESO - " & ooDescrip)
            OLD_EscribeLog ("COMANDO SQL: " & cSQL)
        Case Else
            OLD_EscribeLog ("Error ENVIAR_PLATOS: " & ooError & " (" & ooErrorMDB & ") " & ooDescrip)
            OLD_EscribeLog ("COMANDO SQL: " & cSQL)
    End Select
Else
    OLD_EscribeLog ("Error DESCONOCIDO ENVIAR_PLATOS: " & ooError & " (" & ooErrorMDB & ") " & ooDescrip): OLD_EscribeLog ("COMANDO SQL: " & cSQL)
End If
Sleep 150
If iIndicador < 2 Then
    iIndicador = iIndicador + 1
    Resume
Else
    'OLD_EscribeLog ("Error GRAVE ENVIAR_PLATOS: " & ooError & " - " & ooDescrip)
    OLD_EscribeLog ("Error GRAVE ENVIAR_PLATOS: " & ooError & " (" & ooErrorMDB & ") " & ooDescrip)
    OLD_EscribeLog ("COMANDO SQL: " & cSQL)
    MsgBox "ERROR GRAVE AL ENVIAR LA IMPRESION DE PRODUCTOS, CONTACTE A SOLO SOFTWARE", vbCritical, "ERROR GRAVE"
End If
End Function
Private Sub cmdSelMesa_Click()
'INFO: REVISADO ABRIL/2006
If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&

'INFO: 3/SEP/2007
'REVISA SI HAY ALGO PENDIENTE QUE NO SE HA MARCADO Y LO ENVIA A CHEF, CUANDO PRESIONA MESAS
Call ENVIAR_PLATOS("NO MOSTRAR MENSAJE DE PLATOS PENDIENTES, NI PASAR A MESAS")


Debug.Print "cmdSelMesa_Click Lineas: " & CajLin & " - CUENTA: " & nCta & " - MESA: " & nMesa
Call EvalMesaUpdate(CajLin, nCta, nMesa)

Call BuscaMesa(False)
End Sub
Private Sub BuscaMesa(iOpc As Boolean)
'INFO: ACTUALIZADO ABRIL/2006
Dim rsParciales As Recordset
Dim rsCuentas As Recordset
Dim lParc As Integer
Dim nTempoSingle As Single

lGo = False
'StatMesa nMesa, 0
StatMesa nMesa, vbLibre, "PLU.BuscaMesa"

'Llama a la Pantalla que muestra las mesas (Ocupadas/Disponibles)
Mesas.Show 1

'ProgBar.Max = 10

'INFO: 8JUN2011
Frame2(2).Enabled = False
'ProgBar.value = 1: Text1(2).Text = 1
''''rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
''''        " format(precio_unit,'##0.00') as mPrecio_unit," & _
''''        " format(precio,'##0.00') as mPrecio," & _
''''        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
''''        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
''''        " a.caja " & _
''''        " FROM tmp_trans as a " & _
''''        " WHERE a.mesa = " & nMesa & _
''''        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic

Call OpenTMP_TRANS(False)
'ProgBar.value = 2: Text1(2).Text = 2

'INFO: ESTO ES LO QUE HACE QUE SE REPITAN LAS LINEAS
' SEPT 27 2010
'CajLin = rs07.RecordCount
'INFO: SEPT 27 2010
CajLin = GetLastLine()
'ProgBar.value = 3:: Text1(2).Text = 3
'INFO: DOMICILIO MAY/2009
If HAS_Domicilio Then
    If nMesa >= nDomicilio Then
        Dim cLocalPhone As String
        cLocalPhone = MesaAssigned()
        If cLocalPhone = "" Then
            'INFO: MESA NO HA SIDO ASIGNADA
            DomiClientes.Show 1
        Else
            If CajLin > 0 Then
                'SI YA HAY PRODUCTOS MARCADOS, ENTONCES SIGUE CON LA OPERACION NORMAL
            Else
                'DomiClientes.Show 1
                Call DomiClientes.ShowDomi(cLocalPhone)
            End If
        End If
    End If
End If

'ProgBar.value = 4:: Text1(2).Text = 4
Set rsCuentas = New Recordset
Set rsParciales = New Recordset

rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO WHERE MESA = " & nMesa & " GROUP BY MESA", _
        msConn, adOpenDynamic, adLockOptimistic
'ProgBar.value = 5:: Text1(2).Text = 5
If rsParciales.EOF Then lParc = 0 Else lParc = 1

'Si la linea de detalle esta en 0, llama al mesero
If CajLin = 0 Then
    If nFlag = 0 Then
        nMesero = rs!numero
        cNomMesero = rs!nombre
        PLU.Text1(1) = cNomMesero
        StatBar.Panels(1) = "Mesa: " & nMesa
        StatBar.Panels(2) = cNomMesero
        'Meseros.Show 1
        'INFO: PANTALLA PARA INTRODUCCION DE CLIENTE
        '8:25PM 21/01/2005
        If bISThisSocios And nMesa <> 0 Then
            Socios.Show 1
        End If
    Else
        nMesero = rs!numero
        cNomMesero = rs!nombre
        PLU.Text1(1) = cNomMesero
        StatBar.Panels(1) = "Mesa: " & nMesa
        StatBar.Panels(2) = cNomMesero
    End If
    'ProgBar.value = 30
Else
    nMesero = rs!numero
    cNomMesero = rs!nombre
    PLU.Text1(1) = cNomMesero
    StatBar.Panels(1) = "Mesa: " & nMesa
    StatBar.Panels(2) = cNomMesero
    'ProgBar.value = 35
End If
    
On Error Resume Next
    Set PlatosMesa.DataSource = rs07
On Error GoTo 0
'ProgBar.value = 40
SetupPantalla
'ProgBar.value = 45
If PlatosMesa.Rows <> 0 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
'ProgBar.value = 50
If rs07.RecordCount > 0 Then
    If npNumCaj <> rs07!CAJERO Then
        '--- PARA MODULO DE MESEROS
        'MsgBox "USTED ES UN CAJERO DIFERENTE, SE LE PASARA ESTA MESA A SU NOMBRE", vbInformation, BoxTit
        Call SOLOTrans("BEGIN")
        msConn.Execute "UPDATE TMP_TRANS SET CAJERO = " & npNumCaj & " WHERE MESA = " & nMesa
        Call SOLOTrans("COMMIT")
    End If
End If
rs07.Close
'ProgBar.value = 55
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a WHERE a.mesa = " & nMesa, msConn, adOpenStatic, adLockReadOnly
'ProgBar.value = 60
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)

'''''SBTot = Format(SubTot, "STANDARD")
''''On Error Resume Next
''''nTempoSingle = (rs07!precio * iISC)
''''SubTot = FormatCURRENCY((SubTot + nTempoSingle), 2)
''''iISCTransaccion = rs07!precio * iISC
''''SBTot = Format(SubTot, "STANDARD")
''''On Error GoTo 0

'================================================================================
'================================================================================
'INFO: MOVIENDO ESTA RUTINA AQUI, SE ARREGLA EL BUG MAS GRANDE QUE HEMOS
'TENIDO EN 5 AÑOS. EL PROBLEMA ERA QUE LA VARIABLE nCta se limpiaba
'DESPUES QUE SE ACTULIZABA EL SUB TOTAL, Y EN LA RUTINA DEL SUB TOTAL
'SE BUSCABA LA CUENTA DE LA MESA ANTERIOR, QUE PARA LA MESA ACTUAL NO EXISTIA
'02/may/2006
sqltxt = "SELECT MESA,CUENTA FROM TMP_CUENTAS "
sqltxt = sqltxt & " WHERE MESA = " & nMesa
sqltxt = sqltxt & " ORDER BY MESA,CUENTA"

rsCuentas.Open sqltxt, msConn, adOpenKeyset, adLockOptimistic

If rsCuentas.RecordCount > 0 Then lGo = True

If Not rsCuentas.EOF Then
    rsCuentas.MoveFirst
    nCta = rsCuentas!CUENTA
Else
    nCta = 0
End If
rsCuentas.Close
'ProgBar.value = 65
'================================================================================
'================================================================================

Call ActualizaSUBTOTAL
'ProgBar.value = 70
rs07.Close

Text1(1) = cNomMesero
Text1(0) = nMesa

If nMesa = 0 Then
    Frame2(1).Enabled = False
    lbMensaje.BackColor = &HFFFF&
    lbMensaje = "¡¡ DEBE SELECCIONAR UNA MESA !!"
Else
    Frame2(1).Enabled = True
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If
'ProgBar.value = 75
If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
    "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
    Format(rsParciales!Valor, "STANDARD") & Chr(9) & Format(rsParciales!Valor, "STANDARD")
    SubTot = Format(SubTot - rsParciales!Valor, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

lbCuenta = nCta
If lGo = True Then
    If iOpc = False Then
        StatBar.Panels(3) = "Cuentas Separadas"
        cmdCtas_Click
    Else
        StatBar.Panels(3) = ""
    End If
Else
    StatBar.Panels(3) = ""
End If
'ProgBar.value = 80
End Sub

Private Sub cmdUpAcompa_Click()
'INFO: MOSTRAR LOS ACOMPAÑANTES desde el inicio
'INFO: REVISADO ABRIL/2006
Dim iLoc As Integer
Dim iAcom As Integer
Dim nAcoTop As Integer

iLoc = 0: iAcom = 0: nAcoTop = 240

On Error Resume Next
rsTmpAco.MoveFirst
On Error GoTo 0

On Error GoTo ErrAdm:
If rsTmpAco.EOF Then
    If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
    cmdRestoAco.Enabled = False
    cmdAcomp(0).Caption = ""
    cmdAcomp(0).Tag = 0
    For iLocal = 1 To cmdAcomp.Count - 1
        cmdAcomp(iLocal).Visible = False
    Next
    Frame2(2).Enabled = False
Else
    'ACTIVA FRAME DE ACOMPAÑANTES
    Frame2(2).Enabled = True
    iAcom = 0: iLocal = 0
    'COMIENZA A CREAR LOS OTROS 3 BOTONES
    For iLocal = 1 To 3
        On Error Resume Next
            Load cmdAcomp(iLocal)
        On Error GoTo 0
    Next
    On Error Resume Next
    Do Until rsTmpAco.EOF
        'MUESTRA LOS PRIMEROS 4 ACOMPAÑANTES
        'LES ASIGNA EL DEPARTAMENTO DEL PRODUCTO MARCADO
        If iAcom = 0 Then
            cmdAcomp(iAcom).Visible = True
            cmdAcomp(iAcom).Caption = rsTmpAco!descrip
            cmdAcomp(iAcom).Tag = rs03!DEPTO
            iAcom = iAcom + 1
            rsTmpAco.MoveNext
            If rsTmpAco.EOF Then Exit Do
        End If
        cmdAcomp(iAcom).Visible = True
        'cmdAcomp(iAcom).Top = nAcoTop + 540
        cmdAcomp(iAcom).Top = nAcoTop + 660
        cmdAcomp(iAcom).Caption = rsTmpAco!descrip
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        'nAcoTop = nAcoTop + 540
        nAcoTop = nAcoTop + 660
        nAcoBookMark = rsTmpAco.Bookmark
        rsTmpAco.MoveNext
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
    MsgBox "Este plato tiene demasiados acompañantes" & vbCrLf & "El maximo de acompañantes es 8", vbInformation, "Muchos Acompañantes"
Else
    Resume Next
End If
End Sub

''''Private Sub Command1_Click()
'''''INFO: ESUNA PRUEBA
'''''CREAR UN COMMAND1 OBJECT PARA CORRERLA
''''Dim i As Integer
''''DoEvents
''''For i = 1 To 200
''''rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
''''        " format(precio_unit,'##0.00') as mPrecio_unit," & _
''''        " format(precio,'##0.00') as mPrecio," & _
''''        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
''''        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
''''        " a.caja " & _
''''        " FROM tmp_trans as a " & _
''''        " WHERE a.mesa = " & nMesa & _
''''        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic
''''On Error GoTo ErrAdm:
''''    Set PlatosMesa.DataSource = rs07
''''    SetupPantalla
''''    rs07.Close
''''On Error GoTo 0
''''Next
''''MsgBox "200 ok"
''''Exit Sub
''''
''''ErrAdm:
''''MsgBox Err.Source & "---" & Err.Description, vbCritical, BoxTit
''''End Sub

Private Sub Command1_Click()

txtInfo = "Clave para DESCUENTOS Mesa"
AskClave.Show 1
If OkAnul = 1 Then
    DescuentoMesa.Show 1
    OkAnul = 0
    'Call OpenTMP_TRANS(False)
    DoEvents
    On Error Resume Next
    rs07.Requery
    Set PlatosMesa.DataSource = rs07
    On Error GoTo 0
    Call SetupPantalla
Else
    MsgBox "NO Tiene AUTORIZACION para DESCUENTOS DE MESA", vbExclamation, BoxTit
End If
End Sub

Private Sub Command13_Click(Index As Integer)
'INFO: REVISADO ABRIL/2006
Dim DescResp As Variant

Select Case Index
'~~~~~~~~~~~~~~~~~~~~~~
Case 0  'DESCUENTO
'~~~~~~~~~~~~~~~~~~~~~~
    'Descuento al ultimo producto de la lista
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    txtInfo = "Escriba Clave para DESCUENTO"
    AskClave.Show 1
    If OkAnul = 1 Then
        
        Call DescProducto
        
        OkAnul = 0
        Call SetupPantalla
    End If
'~~~~~~~~~~~~~~~~~~~~~~
Case 1  'ANULACION DE LINEA
'~~~~~~~~~~~~~~~~~~~~~~
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    txtInfo = "Escriba Clave para ANULAR Linea"
    AskClave.Show 1
    If OkAnul = 1 Then
        BorraLin.Show 1
        OkAnul = 0
        Call SetupPantalla
    Else
        MsgBox "NO Tiene AUTORIZACION para ANULAR esta linea", vbExclamation, BoxTit
    End If
'~~~~~~~~~~~~~~~~~~~~~~
Case 2  'Reporte de X
'~~~~~~~~~~~~~~~~~~~~~~
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    If REPCAJAX_OK = True Then
        RptCajas.RepCajX
        MsgBox "GENERACION DEL REPORTE EN (X) HA FINALIZADO", vbInformation, BoxTit
        cmdSalir_Click
    Else
        MsgBox "ESTA OPCION NO ESTA DISPONIBLE. CONTACTE A SU ADMINISTRADOR", vbInformation, BoxTit
    End If
    'HacerRep_X
'~~~~~~~~~~~~~~~~~~~~~~
Case 3  'IMPRESION DE PRECUENTA
'~~~~~~~~~~~~~~~~~~~~~~
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    '28/MAR/2007
    'INFO: LOS MESEROS ESTABAN IMPRIMIENDO LA PRECUENTA ANTES DE ENVIAR LOS PLATOS
    '      A LA COCINA Y ROBANDOLE DE ESA MANERA A LOS CLIENTES
    Call ENVIAR_PLATOS("NO MOSTRAR MENSAJE DE PLATOS PENDIENTES, NI PASAR A MESAS")

    'Call DGI_ImprPreCta
    
    'INFO: 24AGO2013
   
    Call ImprPreCta
    
    'INFO: ENERO 2010
    'INFO: UPDATE PRAIA 08/OCT/2023
    'SE MARCA EN LA CAJA
    
    If CajLin > 0 Then
        Call MarcaPrecuenta(1)
    End If

    'StatMesa nMesa, 0
    
    StatMesa nMesa, vbLibre, "PRE-CUENTA"
    
    Debug.Print "PRECUENTA Lineas: " & CajLin & " - CUENTA: " & nCta & " - MESA: " & nMesa
    Call EvalMesaUpdate(CajLin, nCta, nMesa)
    If bAutoLogin Then
    Else
        Unload PLU
        ''''-----------PLU.Hide
        ''''-----------Mesas.Hide
        Unload Mesas
        '' --- > LoginMesas.Visible = True
    End If
'~~~~~~~~~~~~~~~~~~~~~~
Case 4  'PAGOS
Case 5
'~~~~~~~~~~~~~~~~~~~~~~
Case 6
'~~~~~~~~~~~~~~~~~~~~~~
    'CORTESIA DE LA CASA. EL PLATO DEBE APARECER CON PRECIO 0.00
    'COLOR ORIGINAL = &HFF&
    'nCortesia = 0
    'Command13(6).BackColor = &HFF00&
    txtInfo = "Escriba Clave para MARCAR Cortesia"
    AskClave.Show 1
    If OkAnul = 1 Then
        nCortesia = 0
        Command13(6).BackColor = &HFF00&
    End If
End Select
nNLinSel = 0
End Sub

Private Sub Command2_Click()
'AVANZA HACIA ABAJO
'INFO: REVISADO ABRIL/2006
Dim num As Integer
num = 0
nNLinSel = 0

'Llamar proc. de limpiar deptos anteriores
Quita_Subrallado (67)
QuitarDeptos

If rs02.EOF = True Then
    rs02.MovePrevious
    cmdDepto(num).Caption = rs02!CORTO
End If

Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!CORTO
        Arreg_Deptos(num) = rs02!codigo
    Else
        If Not IsObject(cmdDepto(num)) Then
           Load cmdDepto(num)
        End If
        cmdDepto(num).Caption = rs02!CORTO
        Arreg_Deptos(num) = rs02!codigo
        'cmdDepto(num).Left = 120
        cmdDepto(num).Left = 90
        'INFO: CAMBIO A 800X600
        ''''cmdDepto(num).Top = cmdDepto(num - 1).Top + 540
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = 10 Then Exit Do
    'If num = 11 Then Exit Do
    rs02.MoveNext
Loop

End Sub

Private Sub Command3_Click()
'AVANZA HACIA ARRIBA
'INFO: REVISADO ABRIL/2006
Dim nNum As Integer
num = 0
nNLinSel = 0

rs02.MoveFirst
rs02.Find "codigo = " & Arreg_Deptos(0)

If rs02.EOF Then
    'El PG-DOWN llego a la ultima pantalla
    rs02.MovePrevious
    'cmdDepto(num).Caption = rs02!CORTO
End If

rs02.Move -10
If rs02.BOF Then rs02.MoveFirst
'If (nNum - 11) <= 0 Then
'    ' Desde el principio
'    rs02.MoveFirst
'Else
'    'rs02.Move (-12)
'    rs02.Move (-11)
'End If
'Llamar proc. de limpiar deptos anteriores
Quita_Subrallado (67)
QuitarDeptos

'CARGANDO EL CODIGO DE LOS DEPARTAMENTO EN LOS BOTONES DISPONIBLES
Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!CORTO
        Arreg_Deptos(num) = rs02!codigo
    Else
        If Not IsObject(cmdDepto(num)) Then
           Load cmdDepto(num)
        End If
        cmdDepto(num).Caption = rs02!CORTO
        Arreg_Deptos(num) = rs02!codigo
        'cmdDepto(num).Left = 120
        cmdDepto(num).Left = 90
        'INFO: CAMBIO A 800X600
        ''''cmdDepto(num).Top = cmdDepto(num - 1).Top + 540
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = 10 Then Exit Do
    rs02.MoveNext
Loop

End Sub

Private Sub cmdPlus_Click(Index As Integer)
'INFO: ACTUALIZADO MAYO/2006
Dim CadenaSql As String
Dim nLineas As Long
Dim i As Integer
Dim rsParciales As Recordset
Dim lParc As Integer
Dim SOLO_FECHA As String
Dim nTempoSingle  As Single
Dim cSQL As String

If cmdPlus(Index).Tag = "" Then Beep: Exit Sub

'MsgBox cmdPlus(Index).Width & " x " & cmdPlus(Index).Height

'INFO: 19DIC2010
If nCantidad < 1 Then
    EscribeLog "MARCACION DE CERO(EN CANTIDAD). NO SE PERMITE EN ESTA OPERACION"
    ShowMsg "LA CANTIDAD ES INVALIDA", vbRed, vbYellow
    
    nPase = 0
    nCantidad = 1
    Text1(2) = nCantidad
    Exit Sub
End If

i = 0
'Si quiere marcar PRODUCTOS y no hay Mesero, EXIGIRLO!!!!
Do Until nMesero > 0
    nMesero = rs!numero
    cNomMesero = rs!nombre
    StatBar.Panels(1) = "Mesa: " & nMesa
    StatBar.Panels(2) = cNomMesero
    PLU.Text1(1) = cNomMesero
    'Meseros.Show 1
Loop

On Error Resume Next
    rs03.MoveFirst
On Error GoTo 0

rs03.Find "codigo = " + cmdPlus(Index).Tag

nAcoTop = 240

If rs03!ENVASES = True Then    'El producto tiene Envase(s)
    If nGlobEnv < 1 Then
        MsgBox "Por Favor Seleccione ENVASE/TAMAÑO", vbInformation, BoxTit
        nCortesia = 1: Command13(6).BackColor = &HFF&
        Exit Sub
    End If

    CadenaSql = "SELECT a.contenedor,a.codigo,a.precio, "
    CadenaSql = CadenaSql & " b.depto,b.descrip,b.corto,b.IMPRESORA, b.CON_TAX  "
    CadenaSql = CadenaSql & " FROM CONTEND_02 as a, PLU as b "
    CadenaSql = CadenaSql & " WHERE a.CODIGO = " & rs03!codigo
    CadenaSql = CadenaSql & " AND a.CONTENEDOR = " & nGlobEnv
    CadenaSql = CadenaSql & " AND a.codigo = b.codigo "
    rs09.Open CadenaSql, msConn, adOpenStatic, adLockReadOnly
    
    If rs09.EOF Then
        MsgBox "Por Favor Seleccione ENVASE/TAMAÑO", vbInformation, BoxTit
        rs09.Close
        nCortesia = 1: Command13(6).BackColor = &HFF&
        Exit Sub
    End If

    CajLin = CajLin + 1
    
    SOLO_FECHA = Format(Date, "YYYYMMDD")
    
    'LOG DE CORTESIAS (31ENE2011)
    If nCortesia = 0 Then
        EscribeLog "CORTESIA. " & RegRead("HKCU\Software\SoloSoftware\SoloMix\LastAuthorization") & _
            " Cant (" & nCantidad & "). Mesa: " & nMesa & " - " & rs09!CORTO + TextEnv & ", valor: " & Format(rs09!precio, "CURRENCY")
    End If


    CadenaSql = "INSERT INTO TMP_TRANS "
    CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE, "
    CadenaSql = CadenaSql & "PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,"
    'CadenaSql = CadenaSql & "IMPRESORA,CON_TAX) VALUES ("
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        'INFO: ACTUALIZACION DE AREAS
                        '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    CadenaSql = CadenaSql & "IMPRESORA,CON_TAX, AREA) VALUES ("
    CadenaSql = CadenaSql & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'"
    CadenaSql = CadenaSql & rs09!CORTO + TextEnv & "'" & "," & nCantidad & "," & rs09!DEPTO & "," & rs09!codigo & ","
    CadenaSql = CadenaSql & nGlobEnv & "," & (rs09!precio * nCortesia) & "," & (rs09!precio * nCortesia * nCantidad) & "," & "'"
    CadenaSql = CadenaSql & SOLO_FECHA & "'" & "," & "'" & Time & "'"
    CadenaSql = CadenaSql & ",'  '," & 0# & "," & nCta & ",FALSE," & rs09!IMPRESORA & ","
    '02/MAY/2006 = CON_TAX
    CadenaSql = CadenaSql & rs09!CON_TAX & "," & nArea & ")"
    nSeleccionCocina = rs09!IMPRESORA
Else
    CajLin = CajLin + 1

    SOLO_FECHA = Format(Date, "YYYYMMDD")
    
    'LOG DE CORTESIAS (31ENE2011)
    If nCortesia = 0 Then
        EscribeLog "CORTESIA. " & RegRead("HKCU\Software\SoloSoftware\SoloMix\LastAuthorization") & _
            " Cant (" & nCantidad & "). Mesa: " & nMesa & " - " & rs03!CORTO + TextEnv & ", valor: " & Format(rs03!precio1, "CURRENCY")
    End If

    CadenaSql = "INSERT INTO TMP_TRANS "
    CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,"
    CadenaSql = CadenaSql & "PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    'INFO: ACTUALIZACION DE AREAS
                    '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'CadenaSql = CadenaSql & "IMPRESORA,CON_TAX) VALUES ("
    CadenaSql = CadenaSql & "IMPRESORA,CON_TAX, AREA) VALUES ("
    CadenaSql = CadenaSql & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'"
    CadenaSql = CadenaSql & rs03!descrip + TextEnv & "'" & "," & nCantidad & "," & rs03!DEPTO & "," & rs03!codigo & ","
    CadenaSql = CadenaSql & 0 & "," & (rs03!precio1 * nCortesia) & "," & (rs03!precio1 * nCortesia * nCantidad) & "," & "'"
    CadenaSql = CadenaSql & SOLO_FECHA & " '" & "," & "'" & Time & "'"
    CadenaSql = CadenaSql & ",'  '," & 0# & "," & nCta & ",FALSE," & rs03!IMPRESORA & ","
    '02/MAY/2006 = CON_TAX
    'CadenaSql = CadenaSql & rs03!CON_TAX & ")"
    CadenaSql = CadenaSql & rs03!CON_TAX & "," & nArea & ")"
    nSeleccionCocina = rs03!IMPRESORA
End If

On Error GoTo ErrAdm:
Call SOLOTrans("BEGIN")
msConn.Execute CadenaSql
Call SOLOTrans("COMMIT")
On Error GoTo 0
'---------  CORTESIA ----------------
nCortesia = 1: Command13(6).BackColor = &HFF&
'------------------------------------
' ------ TRATAMIENTO DE ACOMPAÑANTES
'----------------------------------------------------------------
If rsTmpAco.State = adStateOpen Then rsTmpAco.Close


cSQL = "SELECT A.PLU_ID,A.ACOMP_ID,B.DESCRIP "
cSQL = cSQL & " FROM PLU_ACOMP AS A, ACOMPA AS B "
cSQL = cSQL & " WHERE A.PLU_ID = " & rs03!codigo
cSQL = cSQL & " AND A.ACOMP_ID = B.CODIGO "
cSQL = cSQL & " ORDER BY B.DESCRIP "

rsTmpAco.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'BORRA TODOS LOS ACOMPAÑANTES
'INFO: ABRIL2010
For iLocal = 1 To cmdAcomp.Count - 1
    cmdAcomp(iLocal).Caption = ""
    cmdAcomp(iLocal).Tag = 0
    cmdAcomp(iLocal).Visible = False
Next

If rsTmpAco.EOF Then
    If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
    cmdRestoAco.Enabled = False
    cmdAcomp(0).Caption = ""
    cmdAcomp(0).Tag = 0
    For iLocal = 1 To cmdAcomp.Count - 1
        cmdAcomp(iLocal).Visible = False
    Next
    Frame2(2).Enabled = False
Else
    Frame2(2).Enabled = True
    iAcom = 0: iLocal = 0
    
    For iLocal = 1 To 3
        On Error Resume Next
        Load cmdAcomp(iLocal)
        On Error GoTo 0
    Next
    On Error Resume Next
    Do Until rsTmpAco.EOF
        If iAcom = 0 Then
            cmdAcomp(iAcom).Visible = True
            cmdAcomp(iAcom).Caption = rsTmpAco!descrip
            cmdAcomp(iAcom).Tag = rs03!DEPTO
            iAcom = iAcom + 1
            rsTmpAco.MoveNext
            If rsTmpAco.EOF Then Exit Do
        End If
        
        cmdAcomp(iAcom).Visible = True
        'cmdAcomp(iAcom).Top = nAcoTop + 540
        cmdAcomp(iAcom).Top = nAcoTop + 660
        cmdAcomp(iAcom).Caption = rsTmpAco!descrip
        cmdAcomp(iAcom).Tag = rs03!DEPTO
        iAcom = iAcom + 1
        'nAcoTop = nAcoTop + 540
        nAcoTop = nAcoTop + 660
        nAcoBookMark = rsTmpAco.Bookmark
        rsTmpAco.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    Loop
    On Error GoTo 0
End If
'---------------------------------------------------------------------------------------------------------------
' ------------------FIN TRATAMIENTO DE ACOMPAÑANTES---------------------------------
'---------------------------------------------------------------------------------------------------------------
If rs03!ENVASES = True Then rs09.Close

If CajLin = 1 Then
    'msConn.Execute "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & rs!numero & " WHERE numero = " & nMesa
    cSQL = "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & rs!numero & " WHERE numero = " & nMesa
    Call SQL_Update(False, "cmdPlus_Click", cSQL)
End If

If rs07.State = adStateOpen Then rs07.Close
''''rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
''''        " format(precio_unit,'##0.00') as mPrecio_unit," & _
''''        " format(precio,'##0.00') as mPrecio," & _
''''        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
''''        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
''''        " a.caja " & _
''''        " FROM tmp_trans as a " & _
''''        " WHERE a.mesa = " & nMesa & _
''''        " AND A.CUENTA = " & nCta & _
''''        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic

Call OpenTMP_TRANS(True)

'DoEvents       'INFO: 08AGO2016. REMOVIENDO DOEVENTS

On Error Resume Next
Set PlatosMesa.DataSource = rs07
On Error GoTo 0

'DoEvents
SetupPantalla

nLineas = PlatosMesa.Rows - 1

If cHAYClientes = "SI" Then
    'INFO: PONIENDO CLIENTES EN LA MESA
    If nLineas + 1 = 1 Then Call PutClientesOnMesa(nClientesOnMesa)
End If

Set rsParciales = New Recordset

cSQL = "SELECT MESA,SUM(MONTO) AS VALOR "
cSQL = cSQL & " FROM TMP_PAR_PAGO "
cSQL = cSQL & " WHERE MESA = " & nMesa
cSQL = cSQL & " GROUP BY MESA"

rsParciales.Open cSQL, msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

If rs07.State = adStateClosed Then Else rs07.Close

cSQL = "SELECT sum(a.precio) as precio "
cSQL = cSQL & " FROM tmp_trans as a "
cSQL = cSQL & " WHERE a.mesa = " & nMesa
cSQL = cSQL & " AND A.CUENTA = " & nCta

rs07.Open cSQL, msConn, adOpenStatic, adLockReadOnly

SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)

Call ActualizaSUBTOTAL

rs07.Close
'SBTot = Format(SubTot, "STANDARD")

'''''On Error Resume Next
''''nTempoSingle = (rs07!precio * iISC)
''''SubTot = FormatCURRENCY((SubTot + nTempoSingle), 2)
''''iISCTransaccion = rs07!precio * iISC
''''SBTot = Format(SubTot, "STANDARD")
'''''On Error GoTo 0

'DoEvents
If (PlatosMesa.Rows - 1) >= 1 Then
    'DoEvents
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
nCantidad = 1: nPase = 0
'INFO: CAMBIO A 800X600
StatBar.Panels(4) = 1
nNLinSel = 0
Text1(2) = nCantidad

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!Valor, "STANDARD") & Chr(9) & Format(rsParciales!Valor, "STANDARD")
    SubTot = Format(SubTot - rsParciales!Valor, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

On Error GoTo 0
Exit Sub

ErrAdm:
    If Err.Number = 3704 Or Err.Number = 3705 Then
    Else
        'MsgBox Err.Number & Space(1) & Err.Description, vbCritical, nLin
        ShowMsg Err.Number & Space(1) & Err.Description, vbYellow, vbRed
        'INFO: 18DIC2017. UPDATE, VA PARA EL LOG
        EscribeLog "Meseros." & cMachineName & ".cmdPLUS_Click. " & Err.Number & Space(1) & Err.Description
        EscribeLog "cSQL: " & cSQL
        EscribeLog "CadenaSql : " & CadenaSql
        MsgBox CadenaSql
        
    End If
    Resume Next
End Sub

Private Sub Command5_Click()
'INFO: REVISADO MAYO/2006
Dim nNum As String
num = 0
'Tengo que saber quien se ve de primero para mostrar
'los 11 anteriores

If rs03.EOF Then
    If rs03.BOF Then Exit Sub
    rs03.MovePrevious
    cmdPlus(num).Caption = rs03!descrip
    cmdPlus(numplu).Tag = rs03!codigo
End If
nNum = rs03!codigo

'If (numplu - 17) <= 0 Then
'INFO: 1024 ---> If (numplu - 23) <= 0 Then
If (numplu - 23) <= 0 Then
    ' Desde el principio
    rs03.MoveFirst
Else
    'rs03.Move (-17)
    'INFO: 1024 ---> rs03.Move (-23)
    rs03.Move (-23)
    If rs03.BOF Then
        rs03.MoveFirst
    End If
End If
'Llamar proc. de limpiar deptos anteriores
'Quita_Subrallado (67)
'QuitarDeptos
'QuitarPLUS
MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    'INFO: CAMBIO A 800X600
    ''''MiLeft = MiLeft + 1900
    MiLeft = MiLeft + 2400

    'If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
    'INFO: 1024 ---> If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
    If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
        'INFO: CAMBIO A 800X600
        ''''MiTop = MiTop + 460
        'MiTop = MiTop + 600
        'MiTop = MiTop + 670
        MiTop = MiTop + 865
        MiLeft = 0
    End If
    'If numplu = 18 Then Exit Do
    'INFO: 1024 ---> If numplu = 24 Then Exit Do
    If numplu = 24 Then Exit Do
    rs03.MoveNext
Loop
nNLinSel = 0
End Sub

Private Sub Command6_Click()
'INFO: REVISADO MAYO/2006

'If cmdPlus(17).Visible = False Then
'INFO: 1024 ---> If cmdPlus(23).Visible = False Then
If cmdPlus(23).Visible = False Then
    Exit Sub
End If

numplu = 0
'Llamar proc. de limpiar PLUS anteriores
QuitarPLUS
If rs03.EOF Then
    rs03.MovePrevious
    cmdPlus(numplu).Caption = rs03!descrip
End If

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    'INFO: CAMBIO A 800X600
    ''''MiLeft = MiLeft + 1900
    MiLeft = MiLeft + 2400
    'If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
    'INFO: 1024 ---> If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
    If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
        'INFO: CAMBIO A 800X600
        ''''MiTop = MiTop + 460
        'MiTop = MiTop + 600
        'MiTop = MiTop + 670
        MiTop = MiTop + 865
        MiLeft = 0
    End If
    'If numplu = 18 Then Exit Do
    'INFO: 1024 ---> If numplu = 24 Then Exit Do
    If numplu = 24 Then Exit Do
    
    rs03.MoveNext
Loop
nNLinSel = 0
End Sub

Private Sub Command8_Click(Index As Integer)
'INFO: SE MARCA LA CANTIDAD DE ARTCULOS
'INFO: REVISADO MAYO/2006
Dim cCant As String

On Error GoTo FixError:

If nPase = 0 Then
    nCantidad = Command8(Index).Index
Else
    cCant = Str(nCantidad)
    cCant = cCant & Command8(Index).Index
    nCantidad = Val(cCant)
    If Len(cCant) = 4 Then
        ShowMsg "CANTIDAD NO ES VALIDA, ESTABLECIENDO UNO (1)" & vbCrLf & vbCrLf & _
            "INSERTE UNA CANTIDAD MENOR DE 100"
        nPase = 0
        nCantidad = 1
        Text1(2) = nCantidad
        StatBar.Panels(4) = 1
        Exit Sub
    End If
End If
On Error GoTo 0

StatBar.Panels(4) = nCantidad
Text1(2) = nCantidad
nPase = nPase + 1
Exit Sub

FixError:
    nCantidad = 1
    StatBar.Panels(4) = 1
    Resume Next
End Sub

Private Sub Command8_GotFocus(Index As Integer)
'INFO: 28NOV2013
Command8(Index).BackColor = &HFFFF00
End Sub
Private Sub Command8_LostFocus(Index As Integer)
'INFO: 28NOV2013
Command8(Index).BackColor = &H8000000F
End Sub

Private Sub Correccion_Click()
'------------------- CORRECCION / ERROR CORRECT ----------------
'INFO: NO SE PUEDE HACER CORRECCION SI LA ULTIMA
'LINEA ES UN PAGO PARCIAL
'INFO: ACTUALIZADO MAYO/2006
Dim rsFixTmpTrans As Recordset
Dim txto As String
Dim rsParciales As Recordset
Dim lParc As Integer
Dim sqltext As String
Dim SSD As Single
Dim nTp  As Integer
Dim nn, i As Integer
Dim zz As Integer
Dim SOLO_FECHA As String
Dim nVeriCant As Integer
'---------------------
Dim nLocLin As Integer
Dim nLocCan As Integer
Dim nTempoSingle  As Single
Dim cSQL As String
Dim bPrintERRORCorrect As Boolean
'INFO: 08AGO2016
Dim ccError As String

bPrintERRORCorrect = True
nNLinSel = 0: nTp = 0

''''''If UCase(GetENCRYPTEDINI("Facturacion", "AllowFreeCorrection", App.path & "\soloini.ini")) = "PEREZA" Then
'''''If UCase(GetFromINI("Facturacion", "AllowFreeCorrection", App.Path & "\soloini.ini")) = "PEREZA" Then
'''''Else
'''''    txtInfo = "Escriba Clave para CORREGIR"
'''''    AskClave.Show 1
'''''    If OkAnul = 1 Then
'''''    Else
'''''        MsgBox "NO Tiene AUTORIZACION para CORREGIR esta linea", vbExclamation, BoxTit
'''''        Exit Sub
'''''    End If
'''''End If

On Error Resume Next
    PlatosMesa.row = PlatosMesa.Rows - 1
    PlatosMesa.col = 0
    nLocLin = Val(PlatosMesa.Text)
    PlatosMesa.col = 2
    nLocCan = PlatosMesa.Text
On Error GoTo 0

On Error GoTo ErrAdm:
Set rsFixTmpTrans = New Recordset

txto = "SELECT * FROM tmp_trans "
txto = txto & " WHERE mesa = " & nMesa
txto = txto & " AND lin = " & nLocLin
rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

If rsFixTmpTrans.EOF = True Then
    rsFixTmpTrans.Close
    Exit Sub
End If

If rsFixTmpTrans!CANT < 0 Then
    ShowMsg "NO puede CORREGIR este Producto", vbRed, vbYellow
    rsFixTmpTrans.Close
    Exit Sub
End If

If Mid(rsFixTmpTrans!TIPO, 1, 1) = "B" Then
    ShowMsg "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!TIPO, 5, 2)), vbRed, vbYellow
    rsFixTmpTrans.Close
    Exit Sub
End If

'INFO: AHORA EL PROGRAMA CONTROLA LO QUE SE PUEDE CORREGIR
'SI EL PRODUCTO NO HA SIDO ENVIADO A CHEF (YA SE IMPRIMIO
'EN EL BAR o COCINA), ENTONCES PERMITE ANULACION SIN PROBLEMAS
'SI YA FUE ENVIADO A CHEF, EL PRODUCTO QUE SE DESEA CORREGIR SOLAMENTE
'SE PUEDE CAMBIAR SU ESTATUS, SI ES ANULADO
'(03/SEP/2006)
If rsFixTmpTrans!IMPRESO Then
    ShowMsg "NO puede CORREGIR un Producto que ya fue enviado a CHEF" & vbCrLf & _
           "DEBE DE ANULARLO EN LA CAJA!", vbRed, vbYellow
    rsFixTmpTrans.Close
    Exit Sub
End If

'***********************************
'INFO: 24JUL2010
'***********************************
If Mid(rsFixTmpTrans!TIPO, 1, 3) = "DC-" Then
    'ESTA LINEA ES UN DESCUENTO, BUSCAR LA LINEA A LA QUE SE LE ESTA DANDO EL
    'DESCUENTO y QUITARLE ESA INFORMACION
    If GetEnteroFromString(rsFixTmpTrans!TIPO) = "" Then
        'DO NOTHING
    Else
        'INFO: QUITA EL DESCUENTO MARCADO EN EL CAMPO (TIPO) EN ESA LINEA
        Call SOLOTrans("BEGIN")
            '7OCT2014
            'msConn.Execute "UPDATE TMP_TRANS SET TIPO = ' ' WHERE MESA = " & nMesa & " AND LIN = " & Val(GetEnteroFromString(rsFixTmpTrans!TIPO))
            msConn.Execute "UPDATE TMP_TRANS SET TIPO = '  ' WHERE MESA = " & nMesa & " AND LIN = " & Val(GetEnteroFromString(rsFixTmpTrans!TIPO))
        Call SOLOTrans("COMMIT")
    End If
End If

'21/08/2005
'ERROR CORRECT se elimine antes de enviarlo a CHEF (se elimina el plato y el ERROR CORRECT)
'Solamente cuando el producto NO se ha enviado a CHEF.
'Si ya se envio, entoces la linea fisica no se puede eliminar y se opera normalmente.
'NO aplica para Anulacion (esta sigue trabajando igual)

If Not rsFixTmpTrans!IMPRESO Then
    'INFO: SI AUN NO SE HA ENVIADO A SU IMPRESORA CORRESPONDIENTE
    'ENTONCES ELIMINARLO DIRECTAMENTE
    cSQL = "DELETE FROM TMP_TRANS "
    cSQL = cSQL & " WHERE MESA = " & rsFixTmpTrans!MESA
    cSQL = cSQL & " AND LIN = " & rsFixTmpTrans!LIN
    msConn.Execute cSQL
    bPrintERRORCorrect = False
    CajLin = CajLin - 1
End If

'---------------------------------------
nn = 0: i = 1: zz = 0
'Pregunta si hay un Numero en TIPO, si hay significa que tiene Desc
For i = i To 9
    nn = InStr(1, rsFixTmpTrans!TIPO, i)
    If nn <> 0 Then Exit For
Next
If nn <> 0 Then zz = Val(Mid(rsFixTmpTrans!TIPO, nn, 2))
'----------------------------------------

SSD = rsFixTmpTrans!precio * (-1)

SOLO_FECHA = Format(Date, "YYYYMMDD")


'21/08/2005
If bPrintERRORCorrect Then
    CajLin = CajLin + 1
    
    'INSERTA LA LINEA DE CORRECCION
    CadenaSql = "INSERT INTO TMP_TRANS "
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        'INFO: ACTUALIZACION DE AREAS
                        '5MAY2023
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX) VALUES ("
    CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX,AREA) VALUES ("
    CadenaSql = CadenaSql & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'EC-"
    CadenaSql = CadenaSql & rsFixTmpTrans!descrip & "'" & "," & rsFixTmpTrans!CANT * (-1) & "," & rsFixTmpTrans!DEPTO & "," & rsFixTmpTrans!PLU & ","
    CadenaSql = CadenaSql & rsFixTmpTrans!envase & "," & rsFixTmpTrans!precio_unit * (-1) & "," & SSD & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
    CadenaSql = CadenaSql & ",'EC-" & CajLin - 1 & "'," & 0# & "," & rsFixTmpTrans!CUENTA & ",FALSE," & rsFixTmpTrans!IMPRESORA & ","
    CadenaSql = CadenaSql & rsFixTmpTrans!CON_TAX & "," & nArea & ")"
    
    'MARCA LA LINEA QUE SE ESTA CORRIGIENDO PARA QUE NO PUEDA SER CORREGIDA/ANULADA
    'DE NUEVO
    sqltext = "UPDATE TMP_TRANS SET VALID = 0,TIPO = 'BEC" & Str(CajLin)
    sqltext = sqltext & "' WHERE MESA = " & nMesa
    sqltext = sqltext & " AND LIN = " & (nLocLin)
    
    Call SOLOTrans("BEGIN")
    ''''''''''''''' msConnLoc.BeginTrans
    If zz > 0 Then
        '7OCT2014
        'sqltxt = "UPDATE TMP_TRANS SET TIPO = ' ' WHERE MESA = " & nMesa
        sqltxt = "UPDATE TMP_TRANS SET TIPO = '  ' WHERE MESA = " & nMesa
        sqltxt = sqltxt & " AND LIN = " & zz
    End If
    msConn.Execute CadenaSql
    msConn.Execute sqltext
    If zz > 0 Then
        msConn.Execute sqltxt
    End If
    Call SOLOTrans("COMMIT")
End If

''''rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
''''    " format(precio_unit,'##0.00') as mPrecio_unit," & _
''''    " format(precio,'##0.00') as mPrecio," & _
''''    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
''''    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
''''    " a.caja " & _
''''    " FROM tmp_trans as a " & _
''''    " WHERE a.mesa = " & nMesa & _
''''    " AND A.CUENTA = " & nCta & _
''''    " ORDER BY a.lin ", msConn, adOpenStatic, adLockOptimistic

Call OpenTMP_TRANS(True)

On Error Resume Next
Set PlatosMesa.DataSource = rs07
On Error GoTo 0

SetupPantalla

nLineas = PlatosMesa.Rows - 1

If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR" & _
        " FROM TMP_PAR_PAGO " & _
        " WHERE MESA = " & nMesa & _
        " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
      " WHERE a.mesa = " & nMesa & _
      " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
'NEW SBTot = Format(SubTot, "STANDARD")

''''On Error Resume Next
''''nTempoSingle = (rs07!precio * iISC)
''''SubTot = FormatCURRENCY((SubTot + nTempoSingle), 2)
''''iISCTransaccion = rs07!precio * iISC
''''SBTot = Format(SubTot, "STANDARD")
''''On Error GoTo 0

Call ActualizaSUBTOTAL

rs07.Close
rsFixTmpTrans.Close

nCantidad = 1: nPase = 0
Text1(2) = nCantidad
StatBar.Panels(4) = 1

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!Valor, "STANDARD") & Chr(9) & Format(rsParciales!Valor, "STANDARD")
    SubTot = Format(SubTot - rsParciales!Valor, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If


'=================================
'TRATAMIENTO DE ACOMPAÑANTES
'=================================
Dim iAcoCnt As Integer
iAcoCnt = 0
If Frame2(2).Enabled = True Then
    For iAcoCnt = 0 To 3
        cmdAcomp(iAcoCnt).Caption = ""
        cmdAcomp(iAcoCnt).Tag = 0
    Next
    Frame2(2).Enabled = False
End If
'---------------------------------
TextEnv = ""
'=================================
'INFO: AL DESACTIVAR EL PAGEUP(cmdUpAcompa), SE CORRIGE EL DEPARTAMENTO ZERO
'30/NOV/2004
cmdRestoAco.Enabled = False
cmdUpAcompa.Enabled = False
'=================================
'FIN DE TRATAMIENTO DE ACOMPAÑANTES
'=================================

On Error GoTo 0
Exit Sub

ErrAdm:
'INFO: 08AGO2016
ccError = Err.Number & " - " & Err.Description
MsgBox ccError, vbCritical, BoxTit
EscribeLog "Meseros. " & cMachineName & ".Error CORRECCION. " & ccError
Resume Next
End Sub
Private Sub Form_Load()
'INFO: ACTUALIZADO MAYO/2006
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim iTam As Integer
Dim rsCuentas As Recordset

'~~~~~~~~~~ CAMBIANDO VAMOS A USAR EL RECORSET PUBLICO RS01 ~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~ 23ENE2016 ~~~~~~~~~~
'Set rs01 = New Recordset
Set rs01 = New ADODB.Recordset

Set rs02 = New Recordset
Set rs03 = New Recordset
Set rs04 = New Recordset
Set rs06 = New Recordset
Set rs07 = New Recordset
Set rs08 = New Recordset 'Para Precios de PLU con Envases
Set rs09 = New Recordset 'Para Precios de PLU con Envases
Set rsParciales = New Recordset
Set rsCuentas = New Recordset

nCantidad = 1: cCaja = 1: Text1(2) = nCantidad
StatBar.Panels(4) = 1
num = 0: iTam = 0: nMesa = 0: CajLin = 0: nPase = 0:
nGlobEnv = 0: nCta = 0: nCliNum = 0
nCortesia = 1

nFlag = 0
OkAnul = 0
nNLinSel = 0    'Linea Seleccionada

If UCase(GetFromINI("Facturacion", "AutoPropinaEnPreCuenta", App.Path & "\soloini.ini")) = "PEREZA" Then
    OKProp = 1
Else
    OKProp = 0
End If
OKDesc = 0
OKCancelar = 0

SetupPantalla

Show    'Muestra la Pantalla de Facturacion

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: MOSTRAR PANTALLA DE BUSQUEDA
' 26ABR2023
If UCase(GetFromINI("Facturacion", "Search", App.Path & "\soloini.ini")) = "SI" Then
    Me.ImageLUPA.Visible = True
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



If UCase(MESEROS_TIENEN_BANCO) = "PEREZA" Then
    cmdFacturacion.Visible = True
Else
    cmdFacturacion.Visible = False
End If
'Mesas
'rs01.Open "SELECT numero, iif(ocupada=TRUE,'Ocupada','Libre') AS status FROM mesas", msConn, adOpenDynamic, adLockOptimistic
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO. 16OCT2015.
' PASANDO A STATIC YA QUE NO ES NECESARIO QUE SE DINAMICO
rs01.Open "SELECT numero, iif(ocupada=TRUE,'Ocupada','Libre') AS status FROM mesas", msConn, adOpenStatic, adLockOptimistic
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Departamentos
rs02.Open "SELECT codigo, corto, abierto FROM depto ORDER BY ORDEN", msConn, adOpenDynamic, adLockOptimistic
If rs02.EOF Then
    MsgBox "NO EXISTEN DEPARTAMENTOS. ES NECESARIO CREAR DEPARTAMENTOS DE VENTAS. EL PROGRAMA TERMINARA AHORA", vbCritical, BoxTit
    Unload Me
    End
End If
'PLUS del Primer Departamento
rs03.Open "SELECT codigo,depto,descrip,corto,precio1,envases,IMPRESORA,CON_TAX " & _
        " FROM PLU " & _
        " WHERE depto = " & rs02!codigo & _
        " AND DISPONIBLE = TRUE " & _
        " ORDER BY DESCRIP", msConn, adOpenStatic, adLockOptimistic
'Contened_01
rs04.Open "SELECT a.depto,a.contenedor,b.descrip FROM contend_01 as a,contened as b WHERE a.DEPTO = " & rs02!codigo & " and a.contenedor = b.contenedor ORDER BY a.depto,a.contenedor", msConn, adOpenStatic, adLockOptimistic
'Cajas
rs06.Open "SELECT caja_cod, descrip FROM cajas WHERE caja_cod = " & cCaja, msConn, adOpenStatic, adLockReadOnly

If TipoApplicacion <> "" Then
    Command13(1).Enabled = False
    Correccion.Enabled = False
End If

If Not bPermiteAnular Then Command13(1).Enabled = False

If ON_LINE Then
    'PLU.Caption = PLU.Caption + "." + rs06!descrip + "." + rs00!descrip + ".ON-LINE" + TipoApplicacion
    PLU.Caption = PLU.Caption + ".CAJERO-MESERO." + rs00!descrip + ".ON-LINE" + TipoApplicacion
Else
    PLU.Caption = PLU.Caption + "." + rs06!descrip + "." + rs00!descrip + ".OFF-LINE" + TipoApplicacion
    PLU.Caption = PLU.Caption + ".CAJERO-MESERO." + rs00!descrip + ".OFF-LINE" + TipoApplicacion
End If

Call BuscaMesa(True)    'Pantalla de Seleccion de Mesa

' Lo maximo que puede caber en el frame de Departamentos son 11
' botones. Indice es entonces 10

'CARGA LOS DATOS EN LOS BOTONES DEPARTAMENTALES DISPONIBLE
'Y SU VALOR EN EL ARREGLO ARREG_DEPTOS
Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!CORTO
        Arreg_Deptos(num) = rs02!codigo
        ElDepto = rs02!codigo
    Else
        Load cmdDepto(num)
        cmdDepto(num).Caption = rs02!CORTO
        Arreg_Deptos(num) = rs02!codigo
        'cmdDepto(num).Left = 120
        cmdDepto(num).Left = 90
        'INFO: CAMBIO A 800X600
        ''''cmdDepto(num).Top = cmdDepto(num - 1).Top + 530
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = 10 Then Exit Do
    'If num = 11 Then Exit Do
    rs02.MoveNext
Loop

'MiTop = 240: StayLeft = 120
MiTop = 240: StayLeft = 90
MiLeft = 0: numplu = 0

'Muestra los PLUs(Botones) del primer departamento

'For i = 1 To 18
'INFO: 1024 ---> For i = 1 To 24
For i = 1 To 24
    Load cmdPlus(i)
Next

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        cmdPlus(numplu).ToolTipText = "Precio : " & Format(rs03!precio1, "CURRENCY")
        'Arreg_Plu(numplu) = numplu
        'Muestra los PLUs del primer departamento
    Else
        'Load cmdPlus(numplu)
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!descrip
        cmdPlus(numplu).Tag = rs03!codigo
        cmdPlus(numplu).ToolTipText = "Precio : " & Format(rs03!precio1, "CURRENCY")
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
    End If
    numplu = numplu + 1
    'INFO: CAMBIO A 800X600
    ''''MiLeft = MiLeft + 1900
    MiLeft = MiLeft + 2400
    'If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
    'INFO: 1024 ---> If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
    If numplu = 4 Or numplu = 8 Or numplu = 12 Or numplu = 16 Or numplu = 20 Then
        ''''MiTop = MiTop + 460
        'INFO: CAMBIO A 800X600
        'MiTop = MiTop + 600
        'MiTop = MiTop + 670
        MiTop = MiTop + 865
        MiLeft = 0
    End If
    
    'If numplu = 18 Then Exit Do
    'INFO: 1024 ---> If numplu = 24 Then Exit Do
    If numplu = 24 Then Exit Do
    rs03.MoveNext
Loop

Do Until rs04.EOF
    cmdEnvases(iTam).Caption = rs04!descrip
    cmdEnvases(iTam).Tag = rs04!contenedor
    iTam = iTam + 1
    rs04.MoveNext
Loop

For iTam = 0 To 3
    If cmdEnvases(iTam).Caption = "" Then
        cmdEnvases(iTam).Enabled = False
    End If
Next

Text1(3) = cNomCaj
Text1(0) = nMesa
If nMesa = 0 Then
    Frame2(1).Enabled = False
    'lbMensaje.BackColor = &HFFFF&
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "! Debe Seleccionar una Mesa !"
Else
    Frame2(1).Enabled = True
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

sqltxt = "SELECT MESA,CUENTA " & _
        " FROM TMP_CUENTAS " & _
        " WHERE MESA = " & nMesa & _
        " ORDER BY MESA,CUENTA"
rsCuentas.Open sqltxt, msConn, adOpenDynamic, adLockOptimistic

If Not rsCuentas.EOF Then
    rsCuentas.MoveFirst
    nCta = rsCuentas!CUENTA
Else
    nCta = 0
End If

rsCuentas.Close
lbCuenta = nCta
If lGo = True Then
    StatBar.Panels(3) = "Cuentas Separadas"
    cmdCtas_Click
Else
    StatBar.Panels(3) = ""
End If

'INFO: DOMICILIO
If HAS_Domicilio Then
    cmdNota.Visible = True
Else
    cmdNota.Visible = False
End If

'INFO: 1024 ---> Me.Width = (1024 * 14.4) - 300
'INFO: 1024 ---> Frame2(2).Left = 11800
'INFO: 1024 ---> Frame2(1).Width = 9840
'INFO: 1024 ---> cmdUpAcompa.Left = 11800
'INFO: 1024 ---> cmdRestoAco.Left = 13195
'INFO: 1024 ---> Command8(1).Left = 11800
'INFO: 1024 ---> Command8(4).Left = 11800
'INFO: 1024 ---> Command8(7).Left = 11800
'INFO: 1024 ---> Command8(2).Left = 12655
'INFO: 1024 ---> Command8(5).Left = 12655
'INFO: 1024 ---> Command8(8).Left = 12655
'INFO: 1024 ---> Command8(3).Left = 13510
'INFO: 1024 ---> Command8(6).Left = 13510
'INFO: 1024 ---> Command8(9).Left = 13510
'INFO: 1024 ---> Command8(0).Left = 12655
End Sub

Private Sub Form_Unload(Cancel As Integer)
'INFO: JUNIO 2010
CajLin = 0
cNomMesero = ""
nMesa = 0
nMesero = 0
cNomCaj = ""
npNumCaj = 0
'INFO: ACTIBANDO PARA RELOJ. 7ENE2019
LoginMesas.Timer1.Enabled = True
End Sub

Private Sub GridDOWN_Click()
On Error Resume Next
'Debug.Print "rOW: " & PlatosMesa.Row & " TOP ROW: " & PlatosMesa.TopRow
PlatosMesa.TopRow = PlatosMesa.TopRow + 1
On Error GoTo 0
End Sub
Private Sub GridUP_Click()
On Error Resume Next
'PlatosMesa.Row = PlatosMesa.Row + 1
PlatosMesa.TopRow = PlatosMesa.TopRow - 1
On Error GoTo 0
End Sub

Private Sub ImageLUPA_Click()


If nMesa = 0 Then
    ShowMsg "PRIMERO DEBE SELECCIONAR MESA", vbYellow, vbRed
    Exit Sub
End If

AdmPLUBusqueda.Show 1

On Error GoTo ErrAdm:
If rsItemAdd.RecordCount > 0 Then
    'SE AGREGO UN ITEM
    
    rsItemAdd.MoveFirst
    Do While Not rsItemAdd.EOF
    
        CajLin = CajLin + 1

        SOLO_FECHA = Format(Date, "YYYYMMDD")
    
        CadenaSql = "INSERT INTO TMP_TRANS "
        CadenaSql = CadenaSql & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,"
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            'INFO: ACTUALIZACION DE AREAS
                            '5MAY2023
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'CadenaSql = CadenaSql & "FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX) VALUES ("
        CadenaSql = CadenaSql & "FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA,CON_TAX,AREA) VALUES ("
        CadenaSql = CadenaSql & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'"
        CadenaSql = CadenaSql & rsItemAdd!descrip + TextEnv & "'" & "," & rsItemAdd!CANT & "," & rsItemAdd!DEPTO & "," & rsItemAdd!codigo & ","
        CadenaSql = CadenaSql & 0 & "," & (rsItemAdd!precio * nCortesia) & "," & (rsItemAdd!precio * nCortesia * rsItemAdd!CANT) & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
        CadenaSql = CadenaSql & ",'  '," & 0# & "," & nCta & ",FALSE," & rsItemAdd!IMPRESORA & ","
        '24/AGO/2005 = CON_TAX
        If IsNull(rsItemAdd!CON_TAX) Then
            'FIX: 11/OCT/2008
            EscribeLog ("Producto NO TIENE IMPUESTO, USANDO PorcentajeImpuesto DEL SOLOINI.INI: " & rsItemAdd!descrip)
            CadenaSql = CadenaSql & GetFromINI("Facturacion", "PorcentajeImpuesto", App.Path & "\soloini.ini") & "," & nArea & ")"
        Else
            CadenaSql = CadenaSql & rsItemAdd!CON_TAX & "," & nArea & ")"
        End If
    
        Call SOLOTrans("BEGIN")
            msConn.Execute CadenaSql
        Call SOLOTrans("COMMIT")
    
    
        rsItemAdd.MoveNext
    Loop
    
    Call Finaliza_INSERT_TMP_TRANS
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: NORMALIZA DISPLAY DEL GRID
    If (PlatosMesa.Rows - 1) >= 1 Then
        PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
End If
On Error GoTo 0
Exit Sub

ErrAdm:

Select Case Err.Number
    Case Is = 91
        Exit Sub
    Case Is = 3704
        'OBJ CERRADO, NO SE SELEECIONO NADA
        Exit Sub
    Case Else
        ShowMsg Err.Number & " - " & Err.Description, vbRed, vbYellow
End Select
End Sub

Private Sub PlatosMesa_Click()
'INFO: REVISADO MAYO/2006
If PlatosMesa.Rows = 0 Then
    ShowMsg "DEBE MARCAR UN PLATO", vbBlue, vbYellow
    Exit Sub
End If
nNLinSel = Val(PlatosMesa.Text)
PlatosMesa.col = 16
nCta = Val(PlatosMesa.Text)
lbCuenta = nCta
PlatosMesa.col = 0
End Sub

Private Sub ActualizaSUBTOTAL()
Dim nTempoSingle As Single
Dim nLocalSubTot As Single
'INFO: SE GUARDA EL DATO EN nTempoSingle, YA QUE AL HACER LA MULTIPLICACION
'SubTot = FormatCURRENCY((SubTot + (rs07!precio * iISC)), 2)
'EL VALOR TENIA DECIMALES EXTRA QUE AFECTABA LA FUNCION FormatCURRENCY
'Y HACIA QUE EL VALOR SE REDONDEARA PARA ARRIBA
'***************************************************************************
'INFO: AQUI ES DONDE ESTA EL CALCULO DEL IMPUESTO
'INFO: REVISADO MAYO/2006

On Error Resume Next
'nTempoSingle = (rs07!precio * iISC)
nTempoSingle = GetTAX()
nLocalSubTot = SubTot
SubTot = FormatCurrency((nLocalSubTot + nTempoSingle), 2)
'iISCTransaccion = rs07!precio * iISC
iISCTransaccion = nTempoSingle
SBTot = Format(nLocalSubTot + nTempoSingle, "STANDARD")
On Error GoTo 0

End Sub

Private Function GetTAX() As Single
'INFO: NUEVO CALCULO DE IMPUESTO
'INFO: REVISADO MAYO/2006
Dim nTAX As Single
Dim rsTAX As ADODB.Recordset
Dim cSQL As String

On Error GoTo ErrAdm:
cSQL = "SELECT SUM(PRECIO * (CON_TAX/100)) AS TAX "
cSQL = cSQL & " FROM TMP_TRANS "
cSQL = cSQL & " WHERE MESA = " & nMesa
cSQL = cSQL & " AND CUENTA = " & nCta
Set rsTAX = New ADODB.Recordset
rsTAX.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If IsNull(rsTAX!TAX) Then
    nTAX = 0
Else
    nTAX = rsTAX!TAX
End If
rsTAX.Close
Set rsTAX = Nothing
GetTAX = nTAX
'INFO: REVISION DEL TAX
''''' Debug.Print nMesa & ") TAX: " & Format(nTAX, "STANDARD") & " - " & nTAX
On Error GoTo 0
Exit Function

ErrAdm:
    'lbMensaje = "ERROR EN PLU.GetTAX()"
    EscribeLog ("ERROR EN FUNCION PLU.GetTAX()")
End Function

'---------------------------------------------------------------------------------------
' Procedimiento : OpenTMP_TRANS
' Autor       : hsequeira
' Fecha       : 18/10/2013
' Proposito   : ABRE LA TABLA DE TMP_TRANS
'---------------------------------------------------------------------------------------
'
Private Sub OpenTMP_TRANS(Optional bCuenta As Boolean)
Dim cSQL As String
Dim oErrorDescrip As String

'Debug.Print Time & " - OpenTMP_TRANS.SQL"
'INFO: REVISADO MAYO/2006. CARGA EN rs07 TODAS LAS TRANSACCIONES DE LA MESA Y CUENTA
On Error GoTo OpenTMP_TRANS_Error

cSQL = "SELECT A.LIN,A.DESCRIP,A.CANT,"
cSQL = cSQL & " FORMAT(A.PRECIO_UNIT,'##0.00') AS MPRECIO_UNIT,"
cSQL = cSQL & " FORMAT(A.PRECIO,'##0.00') AS MPRECIO,"
cSQL = cSQL & " A.CAJERO, A.MESERO, A.DEPTO, A.PLU, A.MESA, A.VALID, "
cSQL = cSQL & " A.ENVASE, A.FECHA, A.HORA, A.TIPO, A.DESCUENTO, A.CUENTA, A.CAJA "
'cSQL = cSQL & " A.CANT & ' x ' & FORMAT(PRECIO_UNIT,'##0.00') & ' = '  & FORMAT(A.PRECIO,'##0.00') AS XXX"
cSQL = cSQL & " FROM TMP_TRANS AS A "
cSQL = cSQL & " WHERE A.MESA = " & nMesa
If bCuenta = True Then
    cSQL = cSQL & " AND A.CUENTA = " & nCta
End If
cSQL = cSQL & " ORDER BY A.LIN "


'INFO: 22ENE2013
' ERROR EN ESTACION DE MESEROS CUANDO LA RED ESTA LENTA
' 3705 - La operación no está permitida si el objeto está abierto.
'Debug.Print Time & " - OpenTMP_TRANS.BeforeOPEN"
rs07.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'Debug.Print Time & " - OpenTMP_TRANS.AfterOPEN"
'Debug.Print Time & " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

On Error GoTo 0
Exit Sub

OpenTMP_TRANS_Error:
    oErrorDescrip = "(" & Err.Number & ") " & Err.Description
    ShowMsg "Meseros. Error: " & oErrorDescrip & vbCrLf & " PLU.OpenTMP_TRANS", vbBlue, vbWhite
    EscribeLog "Meseros." & cMachineName & " Error. PLU.OpenTMP_TRANS: " & oErrorDescrip
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintPedidoBarEnPrinterLocal
' Author    : hsequeira
' Date      : 31/05/2017
' Purpose   : EN VEZ DE ENVIAR EL PEDIDO AL BAR, LO IMPRIME LOCALMENTE
'---------------------------------------------------------------------------------------
'
Private Function PrintPedidoBarEnPrinterLocal(oRS As ADODB.Recordset, cC As String) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(LoginMesas.ImpresoraCuentas.Name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Print2_OPOS_Dev Space(1)
Print2_OPOS_Dev Chr(27) & "|3C" & "ORDEN : " & Format(Right(cC, 2), "00")
Print2_OPOS_Dev Space(1)
Print2_OPOS_Dev Format(Date, "dd MMM yyyy") & Space(4) & Chr(27) & "|3C" & Time
Print2_OPOS_Dev Space(1)
Print2_OPOS_Dev "PEDIDO IMPRESORA BAR LOCAL"
Print2_OPOS_Dev "Mesero : " & Chr(27) & "|3C" & oRS!nombre
Print2_OPOS_Dev Chr(27) & "|2C" & "Mesa # : " & oRS!MESA
Print2_OPOS_Dev "---------------------------"

Do While Not oRS.EOF
    If Not oRS!IMPRESO Then
    
        If Mid(LTrim(oRS!descrip), 1, 2) = "@@" Then
            Print2_OPOS_Dev Space(3) & oRS!descrip
        Else
            Print2_OPOS_Dev Chr(27) & "|3C" & Format(oRS!CANT, "##") & Space(2) & oRS!descrip
        End If
    End If
    oRS.MoveNext
Loop

For i = 1 To 10
    Print2_OPOS_Dev Space(3)
Next

    'INFO: 7MAR2019. IMPRESORA BEMATECH EN ESTACION DE MESEROS
    'INFO: 10sep2021
    If NOM_PRN_FACTURA = "w" Or NOM_PRN_FACTURA = "w1" Then
        Printer.Print Chr(46)
        Printer.EndDoc
    Else
        LoginMesas.ImpresoraCuentas.CutPaper (100)
    End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(LoginMesas.ImpresoraCuentas.Name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

End Function

'---------------------------------------------------------------------------------------
' Procedure : Finaliza_INSERT_TMP_TRANS
' Author    : hsequeira
' Date      : 19/04/2023
' Purpose   : ACTUALIZA EL DISPLAY DE LA PANTALLA PLU
'---------------------------------------------------------------------------------------
'
Private Sub Finaliza_INSERT_TMP_TRANS()
Dim rsParciales As ADODB.Recordset

On Error GoTo ErrAdm:
If CajLin = 1 Then
    CadenaSql = "UPDATE Mesas SET ocupada = TRUE, MESERO_ACTUAL = " & nMesero & " WHERE numero = " & nMesa
    msConn.Execute CadenaSql
End If

Call OpenTMP_TRANS(True)

'------------------------------------------------------
''Call GetPlatosMesaDataSource
'INFO: 27ABR2023
'MANEJO PARA MESEROS
'------------------------------------------------------
            On Error Resume Next
            Set PlatosMesa.DataSource = rs07
            On Error GoTo 0
'------------------------------------------------------

''nLineas = LV.ListItems.Count

Call SetupPantalla
nLineas = PlatosMesa.Rows - 1

If nLineas = 1 Then Call PutClientesOnMesa(nClientesOnMesa)

CadenaSql = "SELECT MESA, SUM(MONTO) AS VALOR "
CadenaSql = CadenaSql & " FROM TMP_PAR_PAGO "
CadenaSql = CadenaSql & " WHERE MESA = " & nMesa
CadenaSql = CadenaSql & " GROUP BY MESA"

Set rsParciales = New ADODB.Recordset
rsParciales.Open CadenaSql, msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

'INFO: UPDATE 21ABR2016 (CadenaSql)
CadenaSql = "SELECT sum(a.precio) as precio "
CadenaSql = CadenaSql & " FROM TMP_TRANS as a "
CadenaSql = CadenaSql & " WHERE a.mesa = " & nMesa
CadenaSql = CadenaSql & " AND A.CUENTA = " & nCta

rs07.Close
rs07.Open CadenaSql, msConn, adOpenStatic, adLockReadOnly

SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
'SBTot = Format(SubTot, "STANDARD")
On Error GoTo 0

Call ActualizaSUBTOTAL
rs07.Close

nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad

If lParc = 1 Then
'
'    Call Add_PAGOPARCIAL(CajLin + 1, Format(rsParciales!Valor, "STANDARD"))
'
'    SubTot = FormatCurrency(SubTot - rsParciales!Valor, 2)
'    SubTot = FormatCurrency(SubTot - rsParciales!Valor, 2)
'    lbMensaje.BackColor = &HFFFF00
'    lbMensaje = "MESA CON PAGOS PARCIALES"
'Else
'    lbMensaje.BackColor = &HFFFFFF
'    lbMensaje = ""
    'INFO: HAY PAGO PARCIAL
    MiLen1 = -1
    Milen2 = Len(Format(rsParciales!Valor * (-1), "STANDARD"))
    If HayPrinterLocal Then
        Print2_OPOS_Dev "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & Space(10 - Milen2) & Format(rsParciales!Valor * (-1), "STANDARD")
    Else
        Print #nFreefile, "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & Space(10 - Milen2) & Format(rsParciales!Valor * (-1), "STANDARD")
    End If
End If

Exit Sub

ErrAdm:
    If Err.Number = -2147417848 Then
        EscribeLog ("Finaliza_INSERT_TMP_TRANS  Error # (" & Err.Number & ") - " & Err.Description)
    Else
        EscribeLog ("Finaliza_INSERT_TMP_TRANS Error # (" & Err.Number & ") - " & Err.Description)
    End If
    EscribeLog ("Finaliza_INSERT_TMP_TRANS Mesa # " & nMesa & " - SQL = " & CadenaSql)
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EvalMesaUpdate
' Author    : hsequeira
' Date      : 05/10/2023
' Purpose   : EVALUA CUANTAS LINEAS HAY MARCADAS EN LA MESA y CUENTA
'---------------------------------------------------------------------------------------
'
Private Function EvalMesaUpdate(nMiLineas As Integer, nMiCuenta As Integer, nMiMesa As Integer) As Boolean
If nMiLineas = 0 And nMiCuenta = 0 Then
    cSQL = "UPDATE MESAS SET MESERO_ACTUAL = 0, OCUPADA =  False,  X_COUNT = 0 WHERE NUMERO = " & nMiMesa
    msConn.BeginTrans
    msConn.Execute cSQL
    msConn.CommitTrans
End If
End Function

