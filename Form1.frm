VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UHF INTERROGATOR DEMO"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Left            =   1680
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   720
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   1680
   End
   Begin Project1.jcFrames jcFrames8 
      Height          =   3375
      Left            =   3120
      Top             =   4560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5953
      BackColor       =   16760832
      FillColor       =   16777215
      TxtBoxShadow    =   1
      Style           =   4
      Caption         =   "OPERATION INFO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   16761024
      GradientHeaderStyle=   2
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         ItemData        =   "Form1.frx":1DF72
         Left            =   0
         List            =   "Form1.frx":1DF74
         TabIndex        =   64
         Top             =   450
         Width           =   8865
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   8040
      Width           =   1575
   End
   Begin Project1.jcFrames jcFrames1 
      Height          =   4335
      Left            =   3120
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7646
      BackColor       =   16777215
      TxtBoxShadow    =   1
      Caption         =   "ISO18000-6B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   16761024
      ColorTo         =   16777215
      HeaderStyle     =   1
      GradientHeaderStyle=   2
      Begin Project1.jcFrames jcFrames3 
         Height          =   1215
         Left            =   120
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2143
         BackColor       =   16760832
         Style           =   5
         Caption         =   "IDENTIFY"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   5
         ColorTo         =   0
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin VB.ComboBox Combo5 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Form1.frx":1DF76
            Left            =   960
            List            =   "Form1.frx":1DF89
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Form1.frx":1DFB4
            Left            =   960
            List            =   "Form1.frx":1DFCA
            TabIndex        =   27
            Text            =   "Combo3"
            Top             =   360
            Width           =   975
         End
         Begin Project1.isButton isButton8 
            Height          =   300
            Left            =   2160
            TabIndex        =   31
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Form1.frx":1DFF4
            Style           =   5
            Caption         =   "IDENTIFY"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton9 
            Height          =   300
            Left            =   2160
            TabIndex        =   32
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Form1.frx":1E010
            Style           =   5
            Caption         =   "STOP"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Times:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1155
         End
      End
      Begin Project1.jcFrames jcFrames4 
         Height          =   2490
         Left            =   120
         Top             =   1740
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4392
         BackColor       =   16760832
         Style           =   5
         Caption         =   "READ AND WRITE"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   5
         ColorTo         =   0
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            MaxLength       =   23
            TabIndex        =   38
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            TabIndex        =   37
            Text            =   "1"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   36
            Text            =   "0"
            Top             =   480
            Width           =   495
         End
         Begin Project1.isButton isButton10 
            Height          =   420
            Left            =   1800
            TabIndex        =   39
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E02C
            Style           =   5
            Caption         =   "READ"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton11 
            Height          =   420
            Left            =   2640
            TabIndex        =   40
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E048
            Style           =   5
            Caption         =   "WRITE"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton12 
            Height          =   420
            Left            =   1800
            TabIndex        =   41
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E064
            Style           =   5
            Caption         =   "LOCK"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton13 
            Height          =   420
            Left            =   2640
            TabIndex        =   42
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E080
            Style           =   5
            Caption         =   "QUERY"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Data(Hex):"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "ByteCnt:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1140
            Width           =   735
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ByteAddr:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ANTENNA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   7
      Top             =   6840
      Width           =   2895
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ANT1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ANT2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ANT3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   660
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ANT4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   660
         Width           =   1215
      End
      Begin Project1.isButton isButton6 
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "Form1.frx":1E09C
         Style           =   5
         Caption         =   "SET"
         IconAlign       =   0
         iNonThemeStyle  =   0
         BackColor       =   16761024
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
      End
      Begin Project1.isButton isButton5 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "Form1.frx":1E0B8
         Style           =   5
         Caption         =   "QUERY"
         IconAlign       =   0
         iNonThemeStyle  =   0
         BackColor       =   16761024
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RF SETTINGS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   2895
      Begin VB.ComboBox Combo12 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Form1.frx":1E0D4
         Left            =   1320
         List            =   "Form1.frx":1E0E4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   675
      End
      Begin Project1.isButton isButton3 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "Form1.frx":1E108
         Style           =   5
         Caption         =   "QUERY"
         IconAlign       =   0
         iNonThemeStyle  =   0
         BackColor       =   16761024
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
      End
      Begin Project1.isButton isButton4 
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "Form1.frx":1E124
         Style           =   5
         Caption         =   "SET"
         IconAlign       =   0
         iNonThemeStyle  =   0
         BackColor       =   16761024
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Power:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "dBm"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "FreqType:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   820
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CONNECTING"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2895
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Form1.frx":1E140
         Left            =   1440
         List            =   "Form1.frx":1E142
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin Project1.isButton isButton1 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "Form1.frx":1E144
         Style           =   5
         Caption         =   "CONNECT"
         IconAlign       =   0
         iNonThemeStyle  =   4
         BackColor       =   16761024
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttBackColor     =   16761024
         ttForeColor     =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16761024
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Form1.frx":1E160
         Left            =   1440
         List            =   "Form1.frx":1E173
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Form1.frx":1E1A9
         Left            =   1440
         List            =   "Form1.frx":1E1C5
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin Project1.isButton isButton2 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "Form1.frx":1E1F9
         Style           =   5
         Caption         =   "DISCON"
         IconAlign       =   0
         iNonThemeStyle  =   0
         BackColor       =   16761024
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Addr Code:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "BaudRate  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comm Port:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
   End
   Begin Project1.jcFrames jcFrames2 
      Height          =   4335
      Left            =   6960
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7646
      BackColor       =   16777215
      TxtBoxShadow    =   1
      Caption         =   "EPC-GEN2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   16761024
      ColorTo         =   16777215
      HeaderStyle     =   1
      GradientHeaderStyle=   2
      Begin Project1.jcFrames jcFrames7 
         Height          =   780
         Left            =   120
         Top             =   3480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1376
         BackColor       =   16760832
         Style           =   5
         Caption         =   "Kill"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderStyle     =   1
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   11
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   300
            Width           =   1245
         End
         Begin Project1.isButton isButton20 
            Height          =   375
            Left            =   3360
            TabIndex        =   63
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Icon            =   "Form1.frx":1E215
            Style           =   5
            Caption         =   "KILL"
            IconAlign       =   0
            iNonThemeStyle  =   0
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1335
         End
      End
      Begin Project1.jcFrames jcFrames5 
         Height          =   1155
         Left            =   120
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2037
         BackColor       =   16760832
         Style           =   5
         Caption         =   "IDENTIFY"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   5
         ColorTo         =   0
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin VB.ComboBox Combo7 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Form1.frx":1E231
            Left            =   1320
            List            =   "Form1.frx":1E247
            TabIndex        =   44
            Text            =   "Combo3"
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Form1.frx":1E271
            Left            =   1320
            List            =   "Form1.frx":1E284
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   720
            Width           =   1095
         End
         Begin Project1.isButton isButton14 
            Height          =   300
            Left            =   2880
            TabIndex        =   45
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Form1.frx":1E2AF
            Style           =   5
            Caption         =   "IDENTIFY"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton15 
            Height          =   300
            Left            =   2880
            TabIndex        =   46
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Form1.frx":1E2CB
            Style           =   5
            Caption         =   "STOP"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Times:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
      End
      Begin Project1.jcFrames jcFrames6 
         Height          =   1755
         Left            =   120
         Top             =   1680
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3096
         BackColor       =   16760832
         Style           =   5
         Caption         =   "ACCESS"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   5
         ColorTo         =   0
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1080
            MaxLength       =   47
            TabIndex        =   56
            Top             =   1320
            Width           =   3615
         End
         Begin VB.ComboBox Combo10 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Form1.frx":1E2E7
            Left            =   1080
            List            =   "Form1.frx":1E303
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox Combo9 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Form1.frx":1E31F
            Left            =   1080
            List            =   "Form1.frx":1E33B
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   650
            Width           =   1095
         End
         Begin VB.ComboBox Combo8 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Form1.frx":1E357
            Left            =   1080
            List            =   "Form1.frx":1E367
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   320
            Width           =   1095
         End
         Begin Project1.isButton isButton16 
            Height          =   420
            Left            =   2520
            TabIndex        =   57
            Top             =   285
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E385
            Style           =   5
            Caption         =   "READ"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton17 
            Height          =   420
            Left            =   3720
            TabIndex        =   58
            Top             =   285
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E3A1
            Style           =   5
            Caption         =   "WRITE"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton18 
            Height          =   420
            Left            =   2520
            TabIndex        =   59
            Top             =   780
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E3BD
            Style           =   5
            Caption         =   "LOCK"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin Project1.isButton isButton19 
            Height          =   420
            Left            =   3720
            TabIndex        =   60
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   741
            Icon            =   "Form1.frx":1E3D9
            Style           =   5
            Caption         =   "INIT"
            IconAlign       =   0
            iNonThemeStyle  =   4
            BackColor       =   16761024
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttBackColor     =   16761024
            ttForeColor     =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16761024
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Data(Hex):"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "MemBank:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "WordPtr:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   680
            Width           =   735
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "WordCnt:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1000
            Width           =   975
         End
      End
   End
   Begin Project1.isButton isButton7 
      Height          =   495
      Left            =   10440
      TabIndex        =   26
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "Form1.frx":1E3F5
      Style           =   5
      Caption         =   "CLEAR"
      IconAlign       =   0
      iNonThemeStyle  =   0
      BackColor       =   16761024
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TAG COUNT:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Left            =   120
      Picture         =   "Form1.frx":1E411
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2940
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Function Display(str$)
    Dim tempstr As String
   If flagTag < 10 Then
    tempstr = "0" & flagTag & " :" & time & str
   Else
    tempstr = flagTag & " :" & time & str
    End If
    List1.AddItem (tempstr)
    flagTag = flagTag + 1
    List1.Selected(List1.ListCount - 1) = True
End Function

Private Sub Combo3_Change()
    If Combo3.Text = "" Or Val(Combo3.Text) > 10000 Then
     Combo3.Text = "10"
   End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
     If Chr(KeyAscii) Like "[!0-9]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Combo7_Change()
    If Combo7.Text = "" Or Val(Combo7.Text) > 10000 Then
     Combo7.Text = "10"
   End If
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
     If Chr(KeyAscii) Like "[!0-9]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Combo4.AddItem "Generic", 0
    For I = 1 To 240
        Combo4.AddItem I, I
    Next

    Combo1.ListIndex = 0
    Combo2.ListIndex = 4
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
    Combo5.ListIndex = 0
    Combo6.ListIndex = 0
    Combo7.ListIndex = 0
    Combo8.ListIndex = 1
    Combo9.ListIndex = 0
    Combo10.ListIndex = 0
    Combo12.ListIndex = 1
    
    isButton1.Enabled = True
    isButton2.Enabled = False
    
    Frame3.Enabled = False
    Text1.Enabled = False
    Text1.BackColor = &HFFC0C0
    Combo12.Enabled = False
    Combo12.BackColor = &HFFC0C0
    isButton3.Enabled = False
    isButton4.Enabled = False
    
    Frame6.Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    isButton5.Enabled = False
    isButton6.Enabled = False
    
    jcFrames3.Enabled = False
    Combo3.Enabled = False
    Combo3.BackColor = &HFFC0C0
    Combo5.Enabled = False
    Combo5.BackColor = &HFFC0C0
    isButton8.Enabled = False
    isButton9.Enabled = False
    
    jcFrames4.Enabled = False
    Text3.Enabled = False
    Text3.BackColor = &HFFC0C0
    Text4.Enabled = False
    Text4.BackColor = &HFFC0C0
    isButton10.Enabled = False
    isButton11.Enabled = False
    isButton12.Enabled = False
    isButton13.Enabled = False
    Text5.Enabled = False
    Text5.BackColor = &HFFC0C0
    
    jcFrames5.Enabled = False
    Combo7.Enabled = False
    Combo7.BackColor = &HFFC0C0
    Combo6.Enabled = False
    Combo6.BackColor = &HFFC0C0
    isButton14.Enabled = False
    isButton15.Enabled = False
    
    jcFrames6.Enabled = False
    Combo8.Enabled = False
    Combo8.BackColor = &HFFC0C0
    Combo9.Enabled = False
    Combo9.BackColor = &HFFC0C0
    Combo10.Enabled = False
    Combo10.BackColor = &HFFC0C0
    isButton16.Enabled = False
    isButton17.Enabled = False
    isButton18.Enabled = False
    isButton19.Enabled = False
    Text6.Enabled = False
    Text6.BackColor = &HFFC0C0
    
    jcFrames7.Enabled = False
    Text7.Enabled = False
    Text7.BackColor = &HFFC0C0
    isButton20.Enabled = False
    
    Text2.Enabled = False
    Text2.BackColor = &HFFC0C0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim result As Integer
    result = SetBaudRate(hCom, 0, 255)
    result = CommClose(hCom)
    Unload Me
End Sub

Private Sub isButton1_Click()
    Dim temp&, temp1&, tempstr$, strComm$, major As Byte, minor As Byte
    Dim result As Integer
    Dim addr As Integer

    If Combo4.ListIndex = 0 Then
        addr = 255
    Else
        addr = Combo4.ListIndex
    End If
        
       bButton2 = True
       temp = Combo2.ListIndex
       strComm = Combo1.Text
       
       result = CommOpen(hCom, strComm)

       If result = 0 Then
         Display (" Open SerialPort" & (Combo1.ListIndex + 1) & " success!")
    
         temp1 = SetBaudRate(hCom, temp, addr)
           If temp1 = 0 And result = 0 Then
               Display (" Set BaudRate success!")
               
                isButton1.Enabled = False
                isButton2.Enabled = True
                
                Frame3.Enabled = True
                Text1.Enabled = True
                Text1.BackColor = &HFFFFFF
                Combo12.Enabled = True
                Combo12.BackColor = &HFFFFFF
                isButton3.Enabled = True
                isButton4.Enabled = True
                
                Frame6.Enabled = True
                Check1.Enabled = True
                Check2.Enabled = True
                Check3.Enabled = True
                Check4.Enabled = True
                isButton5.Enabled = True
                isButton6.Enabled = True
                
                jcFrames3.Enabled = True
                Combo3.Enabled = True
                Combo3.BackColor = &HFFFFFF
                Combo5.Enabled = True
                Combo5.BackColor = &HFFFFFF
                isButton8.Enabled = True
                isButton9.Enabled = False
                
                jcFrames4.Enabled = True
                Text3.Enabled = True
                Text3.BackColor = &HFFFFFF
                Text4.Enabled = True
                Text4.BackColor = &HFFFFFF
                isButton10.Enabled = True
                isButton11.Enabled = True
                isButton12.Enabled = True
                isButton13.Enabled = True
                Text5.Enabled = True
                Text5.BackColor = &HFFFFFF
                
                jcFrames5.Enabled = True
                Combo7.Enabled = True
                Combo7.BackColor = &HFFFFFF
                Combo6.Enabled = True
                Combo6.BackColor = &HFFFFFF
                isButton14.Enabled = True
                isButton15.Enabled = False
                
                jcFrames6.Enabled = True
                Combo8.Enabled = True
                Combo8.BackColor = &HFFFFFF
                Combo9.Enabled = True
                Combo9.BackColor = &HFFFFFF
                Combo10.Enabled = True
                Combo10.BackColor = &HFFFFFF
                isButton16.Enabled = True
                isButton17.Enabled = True
                isButton18.Enabled = True
                isButton19.Enabled = True
                Text6.Enabled = True
                Text6.BackColor = &HFFFFFF
                
                jcFrames7.Enabled = True
                Text7.Enabled = True
                Text7.BackColor = &HFFFFFF
                isButton20.Enabled = True
                
                Text2.Enabled = True
                Text2.BackColor = &HFFFFFF
               
    
              result = GetFirmwareVersion(hCom, major, minor, 255)
              If result = 0 Then
                    If minor < 10 Then
                    tempstr = " Firmware version is V" & major & ".0" & minor & "!"
                    Else
                    tempstr = " Firmware version is V" & major & "." & minor & "!"
                    End If
                    Display (tempstr)
              Else
                    Display (" Query Firmware version failed!")
              End If
              
                    isButton3_Click
                    isButton5_Click
           Else
              Display (" Set BaudRate failed!")
              result = CommClose(hCom)
           End If
      Else
        tempstr = " Open SerialPort" & (Combo1.ListIndex + 1) & " failed!"
        Display (tempstr)
      End If
      
End Sub

Private Sub isButton10_Click()
    Dim Values(40) As Byte
    Dim tempstr As String, result As Integer, addr As Integer
    Dim n As Integer, j As Integer
    addr = Val(Text3.Text)
    For j = 0 To 1
         List1.Clear
         flagTag = 1
         If Text3.Text <> "" And Text4.Text <> "" Then
             result = IsoSigleTagRead(hCom, addr, Values(0), 255)
            If result = 0 Then
                 List1.AddItem " Read Tag success, Data is:"
                 tempstr = "     "
                 For n = 0 To Val(Text4.Text) - 1
                     tempstr = tempstr & " " & Right$("00" & Hex(Values(n + 1)), 2)
                 Next
                 Display (tempstr)
                 If Val(Text4.Text) > 22 Then
                    Call SendMessage(List1.hwnd, LB_SETHORIZONTALEXTENT, 550 + (Val(Text4.Text) - 22) * 20, ByVal 0&)
                 End If
                 tempstr = ""
             Else
                 List1.AddItem " Read Tag failed!"
             End If
        Else
            If Text3.Text = "" Then List1.AddItem " ByteAddr can't be blank!"
            If Text4.Text = "" Then List1.AddItem " ByteCnt can't be blank!"
        End If
     Next
End Sub

Private Sub isButton11_Click()
    Dim result As Integer, address As Integer, datalength As Integer
    Dim tempstr As String
    Dim I As Integer
    If Text3.Text <> "" Then
        address = Val(Text3.Text)
        If Text5.Text <> "" Then
            datalength = IIf(Len(Text5.Text) Mod 3 = 0, Len(Text5.Text) \ 3, Len(Text5.Text) \ 3 + 1)
            Dim Data1() As String
            Dim Data2(8) As Byte
            Data1 = Split(Text5.Text, " ")
            For I = 0 To datalength - 1
                Data2(I) = CByte("&H" & Data1(I))
                result = IsoWriteTag(hCom, address + I, Data2(I), 255)
            Next
            If result = 0 Then
                Display (" Write Tag success!")
            Else
                Display (" Write Tag failed!")
            End If
        Else
            Display (" Please input Data!")
        End If
    Else
        Display (" Please input ByteAddr!")
    End If
End Sub

Private Sub isButton12_Click()
    Dim tempstr As String, I As Integer, result As Integer
    result = IsoLockTag(hCom, Val(Text3.Text), 255)
    List1.Clear
    flagTag = 1
    If result = 0 Then
       tempstr = " Lock Tag success!"
          Display (tempstr)
    Else
       tempstr = " Lock Tag failed!"
       Display (tempstr)
   End If
End Sub

Private Sub isButton13_Click()
    Dim Value As Byte
    Dim tempstr As String, I As Integer, result As Integer
    List1.Clear
    flagTag = 1
    result = IsoQueryLock(hCom, Val(Text3.Text), Value, 255)
    If result = 0 Then
       tempstr = " Query success!"
               tempstr = tempstr + "StorageAddress " + Text3.Text
               Select Case Value
                  Case Is = 0
                     tempstr = tempstr + " Not Locked!"
                  Case Is = 1
                     tempstr = tempstr + " Locked!"
              End Select
               Display (tempstr)
    Else
       tempstr = " Query failed!"
       Display (tempstr)
   End If
End Sub

Private Sub isButton14_Click()
    Dim Values(100) As TagIds
        Dim tempstr As String, result As Integer
        Dim I As Long, j As Integer
        Dim Count As Long
        result = ClearIDBuffer(hCom, 255)
        List1.Clear
        flagTag = 1
        Dim period(5) As Integer
        period(0) = 10
        period(1) = 50
        period(2) = 100
        period(3) = 500
        period(4) = 1000
    
    If Combo7.Text <> "continurous" Then
        If Combo7.Text > 0 Then
          Timeflag = 1
          Timer3.Interval = period(Combo6.ListIndex)
          Timer3.Enabled = True
          isButton14.Enabled = False
          isButton15.Enabled = True
       Else
          result = Gen2MultiTagIdentify(hCom, Count, Values(0), 255)
          If result = 0 And Count > 0 Then
            List1.AddItem " Read success, TagID is:"
             tempstr = "     "
             For I = 0 To Count - 1
                For j = 0 To 11
                    If (Values(I).Ids(j) And &HF0) / 16 < &H10 Then
                        tempstr = tempstr & Hex((Values(I).Ids(j) And &HF0) / 16)
                    Else
                        tempstr = tempstr & Hex((Values(I).Ids(j) And (&HF0) / 16 + &H7))
                    End If
                    If Values(I).Ids(j) Mod 16 < &H10 Then
                        tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16)
                    Else
                        tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16 + &H7)
                    End If
                     tempstr = tempstr + " "
                Next j
                   tempstr = tempstr + "!"
                   Display (tempstr)
                   tempstr = "     "
             Next I
           End If
        End If
    Else
          Timeflag = 1
          Timer4.Interval = period(Combo6.ListIndex)
          Timer4.Enabled = True
          isButton14.Enabled = False
          isButton15.Enabled = True
    End If
    Text2.Text = flagTag - 1
End Sub

Private Sub isButton15_Click()
    Timer3.Enabled = False
    Timer4.Enabled = False
    isButton14.Enabled = True
    isButton15.Enabled = False
    Combo7.ListIndex = 0
End Sub

Private Sub isButton16_Click()
    Dim tempstr As String, result As Integer
    Dim I As Long, j As Integer
    Dim bank As Integer, ptr As Integer, cnt As Integer
    Dim ID(100)  As Byte
    List1.Clear
    flagTag = 1
    Dim Memb(4) As Integer
            Memb(0) = 0
            Memb(1) = 1
            Memb(2) = 2
            Memb(3) = 3
    bank = Memb(Combo8.ListIndex)
     ptr = CByte(Combo9.Text)
     cnt = CByte(Combo10.Text)
        result = Gen2Read(hCom, bank, ptr, cnt, ID(0), 255)
           If result = 0 Then
             List1.AddItem " Read success, Data is:"
             tempstr = " "
             For I = 0 To cnt * 2 - 1
                tempstr = tempstr & " " & Right$("00" & Hex(ID(I + 1)), 2)
             Next
             Display (tempstr)
            Else
                tempstr = " Read failed!"
                Display (tempstr)
           End If
End Sub

Private Sub isButton17_Click()
    Dim result As Integer, bank As Integer, address As Integer, datalength As Integer, n As Integer
    Dim tempstr As String
    Dim MemBank(3) As Integer
    Dim a() As Long, k As Integer
    Dim wordcnt As Integer
    
'''''    If Combo8.ListIndex = 0 Then
'''''        Display (" Reserve can't be written!")
'''''        Exit Sub
'''''    End If

    If Combo8.Text = "EPC" And (Combo9.ListIndex = 0 Or Combo9.ListIndex = 1) Then
        tempstr = " Address " & Combo9.Text & " of EPC is forbidden to write!"
        Display (tempstr)
        Exit Sub
    End If
    
    MemBank(0) = 0
    MemBank(1) = 1
    MemBank(2) = 2
    MemBank(3) = 3
    bank = MemBank(Combo8.ListIndex)
    address = CByte(Combo9.Text)
    datalength = IIf(Len(Text6.Text) Mod 3 = 0, Len(Text6.Text) \ 3, Len(Text6.Text) \ 3 + 1)
    
    wordcnt = Combo10.ListIndex + 1
    If Text6.Text <> "" And datalength < wordcnt * 2 Then
        Display (" The input data is less than needed!")
        Exit Sub
    End If
    
    If Text6.Text <> "" And datalength > wordcnt * 2 Then
        Display (" The input data is more than needed!")
        Exit Sub
    End If
        
    Dim Data1() As String

    Data1 = Split(Text6.Text, " ")
    
    For n = 0 To datalength - 1 Step 2
        ReDim Preserve a(k)
        a(k) = (CByte("&H" & Data1(n))) * 2 ^ 8 + (CByte("&H" & Data1(n + 1)))
        k = k + 1
    Next
    
    If Text6.Text <> "" Then
        For k = 0 To IIf(datalength Mod 2 <> 0, datalength \ 2 + 1, datalength \ 2) - 1
            result = Gen2Write(hCom, bank, address + k, a(k), 255)
        Next
    Else
        Display (" Please input Data!")
        Exit Sub
    End If
    
    If result = 0 Then
        tempstr = " Write Data success!"
        Display (tempstr)
    Else
        tempstr = " Write Data failed!"
        Display (tempstr)
    End If
    
End Sub

Private Sub isButton18_Click()
    Dim tempstr As String, I As Integer, result As Integer
    result = Gen2LockTag(hCom, 1, 255)
    flagTag = 1
    List1.Clear
    If result = 0 Then
        tempstr = " Lock Tag success!"
    Else
        tempstr = " Lock Tag failed!"
   End If
   Display (tempstr)
End Sub

Private Sub isButton19_Click()
    Dim result As Integer
    Dim tempstr As String
    result = Gen2InitEPC(hCom, 6, 255)
    List1.Clear
    flagTag = 1
    If result = 0 Then
       tempstr = " Tag Init success!"
        Display (tempstr)
    Else
       tempstr = " Tag Init failed!"
        Display (tempstr)
    End If
End Sub

Private Sub isButton2_Click()
    Dim result As Integer
    bButton2 = False
    result = SetBaudRate(hCom, 0, 255)
    
    isButton1.Enabled = True
    isButton2.Enabled = False
    
    Frame3.Enabled = False
    Text1.Enabled = False
    Text1.BackColor = &HFFC0C0
    Combo12.Enabled = False
    Combo12.BackColor = &HFFC0C0
    isButton3.Enabled = False
    isButton4.Enabled = False
    
    Frame6.Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    isButton5.Enabled = False
    isButton6.Enabled = False
    
    jcFrames3.Enabled = False
    Combo3.Enabled = False
    Combo3.BackColor = &HFFC0C0
    Combo5.Enabled = False
    Combo5.BackColor = &HFFC0C0
    isButton8.Enabled = False
    isButton9.Enabled = False
    
    jcFrames4.Enabled = False
    Text3.Enabled = False
    Text3.BackColor = &HFFC0C0
    Text4.Enabled = False
    Text4.BackColor = &HFFC0C0
    isButton10.Enabled = False
    isButton11.Enabled = False
    isButton12.Enabled = False
    isButton13.Enabled = False
    Text5.Enabled = False
    Text5.BackColor = &HFFC0C0
    
    jcFrames5.Enabled = False
    Combo7.Enabled = False
    Combo7.BackColor = &HFFC0C0
    Combo6.Enabled = False
    Combo6.BackColor = &HFFC0C0
    isButton14.Enabled = False
    isButton15.Enabled = False
    
    jcFrames6.Enabled = False
    Combo8.Enabled = False
    Combo8.BackColor = &HFFC0C0
    Combo9.Enabled = False
    Combo9.BackColor = &HFFC0C0
    Combo10.Enabled = False
    Combo10.BackColor = &HFFC0C0
    isButton16.Enabled = False
    isButton17.Enabled = False
    isButton18.Enabled = False
    isButton19.Enabled = False
    Text6.Enabled = False
    Text6.BackColor = &HFFC0C0
    
    jcFrames7.Enabled = False
    Text7.Enabled = False
    Text7.BackColor = &HFFC0C0
    isButton20.Enabled = False
    
    Text2.Enabled = False
    Text2.BackColor = &HFFC0C0
    
    result = CommClose(hCom)
End Sub

Private Sub isButton20_Click()
    Dim result As Integer
    Dim pwd As Long
    Dim n As Integer
    Dim str() As String
    Dim s As String
    str = Split(Text7.Text, " ")
    For n = 0 To 3
        s = s & str(n)
    Next
'    For n = 0 To 3
'''''        data(n) = CByte("&H" & PassWord(n))
'''''    Next

    pwd = "&H" & Val(s)
    If Text7.Text <> "" Then
        result = Gen2KillTag(hCom, pwd, 255)
        If result = 0 Then
            Display (" Kill the tag success!")
        Else
            Display (" Kill the tag failed!")
        End If
    Else
        Display (" Please input Password!")
    End If
End Sub

Private Sub isButton3_Click()
    Dim power As Byte, freq_type As Byte
    Dim result As Integer
    Dim tempstr As String
    result = GetRf(hCom, power, freq_type, 255)
      If result = 0 Then
            tempstr = " Querry RF Parameter success!"
            Display (tempstr)
            Text1.Text = power
            tempstr = " Power is " & power & " dBm!"
            Display (tempstr)
            Select Case freq_type
                    Case 0
                        Combo12.Text = Combo12.List(0)
                        tempstr = " Reader Frequency is " & Combo12.Text & " Frequency!"
                        Display (tempstr)
                    Case 1
                        Combo12.Text = Combo12.List(1)
                        tempstr = " Reader Frequency is " & Combo12.Text & " Frequency!"
                        Display (tempstr)
                    Case 2
                        Combo12.Text = Combo12.List(2)
                        tempstr = " Reader Frequency is " & Combo12.Text & " Frequency!"
                        Display (tempstr)
                    Case Else
                        Combo12.Text = Combo12.List(3)
                        tempstr = " Unknown Frequency!"
                        Display (tempstr)
            End Select
        Else
            tempstr = " Query Reader Frequency failed!"
            Display (tempstr)
        End If
End Sub

Private Sub isButton4_Click()
    Dim result As Integer, tempstr As String
    Dim a(3) As Byte
    a(0) = 0
    a(1) = 1
    a(2) = 2
    If Text1.Text <> "" And Combo12.ListIndex <> 3 Then
            result = SetRf(hCom, CByte(Text1.Text), a(Combo12.ListIndex), 255)
            If result = 0 Then
                tempstr = " Set Reader RF Parameter success!"
            Else
                tempstr = " Set Reader RF Parameter failed!"
            End If
            Display (tempstr)
    End If
End Sub

Private Sub isButton5_Click()
Dim result As Integer
Dim ant As Byte
Dim tempstr As String
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    result = GetAnt(hCom, ant, 255)
    If result = 0 Then
        If (ant And &H1) = 1 Then
            Check1.Value = 1
            tempstr = " Ant is " & Check1.Caption & " !"
            Display (tempstr)
        End If
        If (ant And &H2) = &H2 Then
            Check2.Value = 1
            tempstr = " Ant is " & Check2.Caption & " !"
            Display (tempstr)
        End If
        If (ant And &H4) = &H4 Then
            Check3.Value = 1
            tempstr = " Ant is " & Check3.Caption & " !"
            Display (tempstr)
        End If
        If (ant And &H8) = &H8 Then
            Check4.Value = 1
            tempstr = " Ant is " & Check4.Caption & " !"
            Display (tempstr)
        End If
    Else
        Display (" Query Ant failed!")
    End If
End Sub

Private Sub isButton6_Click()
Dim ant As Byte
Dim result As Integer
Dim tempstr As String
    ant = 0
    If Check1.Value = 1 Then ant = ant Or &H1
    If Check2.Value = 1 Then ant = ant Or &H2
    If Check3.Value = 1 Then ant = ant Or &H4
    If Check4.Value = 1 Then ant = ant Or &H8
    result = SetAnt(hCom, ant, 255)
    If result = 0 Then
        tempstr = " Set Ant success!"
        Display (tempstr)
    Else
        tempstr = " Set Ant failed!"
        Display (tempstr)
    End If
End Sub

Private Sub isButton7_Click()
    List1.Clear
    flagTag = 1
    Text2.Text = ""
End Sub

Private Sub isButton8_Click()
        Dim Values(100) As TagIds
        Dim tempstr As String, result As Integer
        Dim I As Long, j As Integer
        Dim Count As Long
        List1.Clear
        flagTag = 1
        result = ClearIDBuffer(hCom, 255)

        Dim period(5) As Integer
        period(0) = 10
        period(1) = 50
        period(2) = 100
        period(3) = 500
        period(4) = 1000
    
    If Combo3.Text <> "continurous" Then
        If Combo3.Text > 0 Then
          Timeflag = 1
          Timer1.Interval = period(Combo5.ListIndex)
          Timer1.Enabled = True
          isButton8.Enabled = False
          isButton9.Enabled = True
       Else
          result = IsoMultiTagIdentify(hCom, Count, Values(0), 255)
          If result = 0 And Count > 0 Then
            List1.AddItem " Read success, TagID is:"
             tempstr = "     "
             For I = 0 To Count - 1
                For j = 0 To 7
                    If (Values(I).Ids(j) And &HF0) / 16 < &H10 Then
                        tempstr = tempstr & Hex((Values(I).Ids(j) And &HF0) / 16)
                    Else
                        tempstr = tempstr & Hex((Values(I).Ids(j) And (&HF0) / 16 + &H7))
                    End If
                    If Values(I).Ids(j) Mod 16 < &H10 Then
                        tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16)
                    Else
                        tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16 + &H7)
                    End If
                     tempstr = tempstr + " "
                Next j
                   tempstr = tempstr + "!"
                   Display (tempstr)
                   tempstr = "     "
             Next I
           End If
        End If
    Else
          Timeflag = 1
          Timer2.Interval = period(Combo5.ListIndex)
          Timer2.Enabled = True
          isButton8.Enabled = False
          isButton9.Enabled = True
    End If
    Text2.Text = flagTag - 1
End Sub

Private Sub isButton9_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    isButton8.Enabled = True
    isButton9.Enabled = False
    Combo3.ListIndex = 0
End Sub

Private Sub Text1_Change()
    If Val(Text1.Text) > 30 Then
        Text1.Text = "0"
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text3_Change()
    If Val(Text3.Text) > 255 Then
     Text3.Text = "10"
   End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text4_Change()
    If Val(Text4.Text) > 32 Then
     Text4.Text = "0"
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text5_Change()
    If Text5.Text = " " Then Text5.Text = ""
    If Len(Text5.Text) Mod 3 = 2 Then Text5.Text = Text5.Text & " "
    If Right(Text5.Text, 2) = "  " Then Text5.Text = Left(Text5.Text, Len(Text5.Text) - 1)
    If Len(Text5.Text) = 2 And Right(Text5.Text, 1) = " " Then
        Text5.Text = "0" & Text5.Text
    Else
        If Len(Text5.Text) > 3 Then
            If Mid(Text5.Text, Len(Text5.Text) - 2, 1) = " " And Right(Text5.Text, 1) = " " Then
                Text5.Text = Mid(Text5.Text, 1, Len(Text5.Text) - 2) & "0" & Right(Text5.Text, 2)
            End If
        End If
    End If
    SendKeys "{end}"
    SendKeys "{DOWN}"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9a-fA-F]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 And KeyAscii <> 32 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text6_Change()
On Error GoTo Err
    If Text6.Text = " " Then Text6.Text = ""
    If Len(Text6.Text) Mod 3 = 2 Then Text6.Text = Text6.Text & " "
    If Right(Text6.Text, 2) = "  " Then Text6.Text = Left(Text6.Text, Len(Text6.Text) - 1)
    If Len(Text6.Text) = 2 And Right(Text6.Text, 1) = " " Then
        Text6.Text = "0" & Text6.Text
    Else
        If Len(Text6.Text) > 3 Then
            If Mid(Text6.Text, Len(Text6.Text) - 2, 1) = " " And Right(Text6.Text, 1) = " " Then
                Text6.Text = Mid(Text6.Text, 1, Len(Text6.Text) - 2) & "0" & Right(Text6.Text, 2)
            End If
        End If
    End If
    SendKeys "{end}"
    SendKeys "{DOWN}"
Err:
    Exit Sub
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9a-fA-F]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 And KeyAscii <> 32 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text7_Change()
    If Text7.Text = " " Then Text7.Text = ""
    If Len(Text7.Text) Mod 3 = 2 Then Text7.Text = Text7.Text & " "
    If Right(Text7.Text, 2) = "  " Then Text7.Text = Left(Text7.Text, Len(Text7.Text) - 1)
    If Len(Text7.Text) = 2 And Right(Text7.Text, 1) = " " Then
        Text7.Text = "0" & Text7.Text
    Else
        If Len(Text7.Text) > 3 Then
            If Mid(Text7.Text, Len(Text7.Text) - 2, 1) = " " And Right(Text7.Text, 1) = " " Then
                Text7.Text = Mid(Text7.Text, 1, Len(Text7.Text) - 2) & "0" & Right(Text7.Text, 2)
            End If
        End If
    End If
    SendKeys "{end}"
    SendKeys "{DOWN}"
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9a-fA-F]" Then
        If KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 And KeyAscii <> 32 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Dim tempstr As String, result As Integer
    Dim Values(100) As TagIds, I As Long, j As Integer
    Dim number As Integer, cnt As Integer
    Dim Count As Long
    number = Combo3.Text
    If number > 0 Then
        result = IsoMultiTagIdentify(hCom, Count, Values(0), 255)
        If result = 0 And Count > 0 Then
            List1.AddItem " Times No." & Timeflag & " Read success!TagID is:"
            Timeflag = Timeflag + 1
            tempstr = "     "
            flagTag = 1
           For I = 0 To Count - 1
              For j = 0 To 7
                  If (Values(I).Ids(j) And &HF0) / 16 < &H10 Then
                      tempstr = tempstr & Hex((Values(I).Ids(j) And &HF0) / 16)
                  Else
                      tempstr = tempstr & Hex((Values(I).Ids(j) And (&HF0) / 16 + &H7))
                  End If
                  If Values(I).Ids(j) Mod 16 < &H10 Then
                      tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16)
                  Else
                      tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16 + &H7)
                  End If
                   tempstr = tempstr + " "
                Next j
                   tempstr = tempstr + "!"
                   Display (tempstr)
                   tempstr = "     "
            Next I
        End If
        Combo3.Text = Combo3.Text - 1
        result = ClearIDBuffer(hCom, 255)
    Else
      Timer1.Enabled = False
      isButton8.Enabled = True
      isButton9.Enabled = False
    End If
    Text2.Text = flagTag - 1
End Sub

Private Sub Timer2_Timer()
    Dim IDNUM(100) As TagIds
    Dim FFZ As Integer
    Dim CountID As Long
    Dim ii As Integer, jj As Integer, kk  As Integer
    Dim tempstr As String, temp1 As String
    Dim rr As Integer
    Dim AddID As Boolean

    FFZ = IsoMultiTagIdentify(hCom, CountID, IDNUM(0), 255)
    If FFZ <> 0 Then
'         --
        Exit Sub
    End If
    If List1.ListCount = 0 Then
        List1.AddItem " Read success, TagID is:"
    End If
    For ii = 0 To CountID - 1
        AddID = False
        tempstr = "     "
        For jj = 0 To 7
            If (IDNUM(ii).Ids(jj) And &HF0) / 16 < &H10 Then
                tempstr = tempstr & Hex((IDNUM(ii).Ids(jj) And &HF0) / 16)
            Else
                tempstr = tempstr & Hex((IDNUM(ii).Ids(jj) And (&HF0) / 16 + &H7))
            End If
            If IDNUM(ii).Ids(jj) Mod 16 < &H10 Then
                tempstr = tempstr & Hex(IDNUM(ii).Ids(jj) Mod 16)
            Else
                tempstr = tempstr & Hex(IDNUM(ii).Ids(jj) Mod 16 + &H7)
            End If
         tempstr = tempstr & " "
        Next
        For kk = 0 To List1.ListCount                  '过滤重复标签
''            If List1.List(kk) <> "" Then
''                temp1 = Mid(List1.List(kk), 4, Len(List1.List(kk)) - 3)
''                If temp1 <> tempStr Then
''                    AddID = True
''                Else
''                    AddID = False
''                    Exit For
''                End If
''            End If
        Next
        
'        If AddID Then
'        List1.AddItem List1.ListCount & ": " & tempStr
          Display (tempstr + "!")
          tempstr = "     "
     Next
     Text2.Text = flagTag - 1
End Sub

Private Sub Timer3_Timer()
    Dim tempstr As String, result As Integer
    Dim Values(100) As TagIds, I As Long, j As Integer
    Dim number As Integer, cnt As Integer
    Dim Count As Long
    
     number = Combo7.Text
    If number > 0 Then
        result = Gen2MultiTagIdentify(hCom, Count, Values(0), 255)
        If result = 0 And Count > 0 Then
            List1.AddItem " Times No." & Timeflag & " Read success!TagID is:"
            Timeflag = Timeflag + 1
            flagTag = 1
           tempstr = "     "
           For I = 0 To Count - 1
              For j = 0 To 11
                  If (Values(I).Ids(j) And &HF0) / 16 < &H10 Then
                      tempstr = tempstr & Hex((Values(I).Ids(j) And &HF0) / 16)
                  Else
                      tempstr = tempstr & Hex((Values(I).Ids(j) And (&HF0) / 16 + &H7))
                  End If
                  If Values(I).Ids(j) Mod 16 < &H10 Then
                      tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16)
                  Else
                      tempstr = tempstr & Hex(Values(I).Ids(j) Mod 16 + &H7)
                  End If
                   tempstr = tempstr + " "
                Next j
                   tempstr = tempstr + "!"
                   Display (tempstr)
                   tempstr = "     "
            Next I
        End If
        Combo7.Text = Combo7.Text - 1
        result = ClearIDBuffer(hCom, 255)
    Else
      Timer3.Enabled = False
      isButton14.Enabled = True
      isButton15.Enabled = False
    End If
    Text2.Text = flagTag - 1
End Sub

Private Sub Timer4_Timer()
    Dim IDNUM(100) As TagIds
    Dim FFZ As Integer
    Dim CountID As Long
    Dim ii As Integer, jj As Integer, kk  As Integer
    Dim tempstr As String, temp1 As String
    Dim rr As Integer
    Dim AddID As Boolean
    FFZ = Gen2MultiTagIdentify(hCom, CountID, IDNUM(0), 255)
    If FFZ <> 0 Then
'         --
        Exit Sub
    End If

    If List1.ListCount = 0 Then
       List1.AddItem "ID of Read Tags:"

    End If
    For ii = 0 To CountID - 1
        AddID = False
        tempstr = "     "
        For jj = 0 To 11
            If (IDNUM(ii).Ids(jj) And &HF0) / 16 < &H10 Then
                tempstr = tempstr & Hex((IDNUM(ii).Ids(jj) And &HF0) / 16)
            Else
                tempstr = tempstr & Hex((IDNUM(ii).Ids(jj) And (&HF0) / 16 + &H7))
            End If
            If IDNUM(ii).Ids(jj) Mod 16 < &H10 Then
                tempstr = tempstr & Hex(IDNUM(ii).Ids(jj) Mod 16)
            Else
                tempstr = tempstr & Hex(IDNUM(ii).Ids(jj) Mod 16 + &H7)
            End If
         tempstr = tempstr & " "
        Next
        
        For kk = 0 To List1.ListCount                  '过滤重复标签
''            If List1.List(kk) <> "" Then
''                temp1 = Mid(List1.List(kk), 4, Len(List1.List(kk)) - 3)
''                If temp1 <> tempStr Then
''                    AddID = True
''                Else
''                    AddID = False
''                    Exit For
''                End If
''            End If
        Next
          Display (tempstr + "!")
          tempstr = "     "
     Next
     Text2.Text = flagTag - 1
End Sub
