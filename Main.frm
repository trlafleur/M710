VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7800
   ClientLeft      =   1830
   ClientTop       =   1545
   ClientWidth     =   10185
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10185
   Begin VB.CommandButton cmdPreset 
      Caption         =   "MM Net 14313"
      Height          =   315
      Index           =   13
      Left            =   5640
      TabIndex        =   137
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "MM Net 14300"
      Height          =   315
      Index           =   12
      Left            =   5640
      TabIndex        =   136
      Top             =   3615
      Width           =   1200
   End
   Begin VB.Frame Frame9 
      Caption         =   "PTC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   5880
      TabIndex        =   132
      Top             =   5160
      Width           =   1575
      Begin VB.ComboBox PTC_Baud 
         Height          =   315
         ItemData        =   "Main.frx":0442
         Left            =   60
         List            =   "Main.frx":047B
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox PTC_Select 
         Caption         =   "PTC II"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   133
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "PTC Baud Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.TextBox radioAddress 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   129
      Top             =   6360
      Width           =   375
   End
   Begin VB.Frame Frame13 
      Caption         =   "Radio Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   4080
      TabIndex        =   125
      Top             =   5160
      Width           =   1695
      Begin VB.OptionButton RadioType 
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   131
         Top             =   1200
         Width           =   1080
      End
      Begin VB.OptionButton RadioType 
         Caption         =   "M700Pro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   128
         Top             =   300
         Width           =   1080
      End
      Begin VB.OptionButton RadioType 
         Caption         =   "M710"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   127
         Top             =   600
         Width           =   900
      End
      Begin VB.OptionButton RadioType 
         Caption         =   "M710RT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   126
         Top             =   900
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   240
      Top             =   6900
   End
   Begin VB.Frame Frame12 
      Caption         =   "Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   8580
      TabIndex        =   98
      Top             =   60
      Width           =   1455
      Begin VB.CommandButton cmdStore 
         Caption         =   "24"
         Height          =   315
         Index           =   23
         Left            =   780
         TabIndex        =   122
         Top             =   4260
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "23"
         Height          =   315
         Index           =   22
         Left            =   780
         TabIndex        =   121
         Top             =   3900
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "22"
         Height          =   315
         Index           =   21
         Left            =   780
         TabIndex        =   120
         Top             =   3540
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "21"
         Height          =   315
         Index           =   20
         Left            =   780
         TabIndex        =   119
         Top             =   3180
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "20"
         Height          =   315
         Index           =   19
         Left            =   780
         TabIndex        =   118
         Top             =   2820
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "19"
         Height          =   315
         Index           =   18
         Left            =   780
         TabIndex        =   117
         Top             =   2460
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "18"
         Height          =   315
         Index           =   17
         Left            =   780
         TabIndex        =   116
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "17"
         Height          =   315
         Index           =   16
         Left            =   780
         TabIndex        =   115
         Top             =   1740
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "16"
         Height          =   315
         Index           =   15
         Left            =   780
         TabIndex        =   114
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "15"
         Height          =   315
         Index           =   14
         Left            =   780
         TabIndex        =   113
         Top             =   1020
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "14"
         Height          =   315
         Index           =   13
         Left            =   780
         TabIndex        =   112
         Top             =   660
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "13"
         Height          =   315
         Index           =   12
         Left            =   780
         TabIndex        =   111
         Top             =   300
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "12"
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   110
         Top             =   4260
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "11"
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   109
         Top             =   3900
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "10"
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   108
         Top             =   3540
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "9"
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   107
         Top             =   3180
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "8"
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   106
         Top             =   2820
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "7"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   105
         Top             =   2460
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "6"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   104
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "5"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   103
         Top             =   1740
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "4"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   102
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "3"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   101
         Top             =   1020
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "2"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   100
         Top             =   660
         Width           =   435
      End
      Begin VB.CommandButton cmdStore 
         Caption         =   "1"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   99
         ToolTipText     =   """Test"""
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Recall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   7020
      TabIndex        =   73
      Top             =   60
      Width           =   1455
      Begin VB.CommandButton cmdRecall 
         Caption         =   "1"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   97
         ToolTipText     =   "Test"
         Top             =   300
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "2"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   96
         Top             =   660
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "3"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   95
         Top             =   1020
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "4"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   94
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "5"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   93
         Top             =   1740
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "6"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   92
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "7"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   91
         Top             =   2460
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "8"
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   90
         Top             =   2820
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "9"
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   89
         Top             =   3180
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "10"
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   88
         Top             =   3540
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "11"
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   87
         Top             =   3900
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "12"
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   86
         Top             =   4260
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "13"
         Height          =   315
         Index           =   12
         Left            =   780
         TabIndex        =   85
         Top             =   300
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "14"
         Height          =   315
         Index           =   13
         Left            =   780
         TabIndex        =   84
         Top             =   660
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "15"
         Height          =   315
         Index           =   14
         Left            =   780
         TabIndex        =   83
         Top             =   1020
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "16"
         Height          =   315
         Index           =   15
         Left            =   780
         TabIndex        =   82
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "17"
         Height          =   315
         Index           =   16
         Left            =   780
         TabIndex        =   81
         Top             =   1740
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "18"
         Height          =   315
         Index           =   17
         Left            =   780
         TabIndex        =   80
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "19"
         Height          =   315
         Index           =   18
         Left            =   780
         TabIndex        =   79
         Top             =   2460
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "20"
         Height          =   315
         Index           =   19
         Left            =   780
         TabIndex        =   78
         Top             =   2820
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "21"
         Height          =   315
         Index           =   20
         Left            =   780
         TabIndex        =   77
         Top             =   3180
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "22"
         Height          =   315
         Index           =   21
         Left            =   780
         TabIndex        =   76
         Top             =   3540
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "23"
         Height          =   315
         Index           =   22
         Left            =   780
         TabIndex        =   75
         Top             =   3900
         Width           =   435
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "24"
         Height          =   315
         Index           =   23
         Left            =   780
         TabIndex        =   74
         Top             =   4260
         Width           =   435
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1635
      Left            =   7560
      TabIndex        =   67
      Top             =   5160
      Width           =   2475
      Begin VB.Label lblVersion 
         Caption         =   "Version:"
         Height          =   195
         Left            =   195
         TabIndex        =   124
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Melbourne, FL 32901"
         Height          =   195
         Left            =   195
         TabIndex        =   71
         Top             =   1200
         Width           =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "1208 East River Drive #302"
         Height          =   195
         Left            =   195
         TabIndex        =   70
         Top             =   1005
         Width           =   2100
      End
      Begin VB.Label Label4 
         Caption         =   "Copyright 2001 - W5SMM"
         Height          =   195
         Left            =   195
         TabIndex        =   69
         Top             =   795
         Width           =   2100
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Icom IC-M710 Control Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame pnlSerialPort 
      Caption         =   "Serial Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   180
      TabIndex        =   58
      Top             =   5160
      Width           =   3795
      Begin VB.CommandButton cmdSetSerialPort 
         Caption         =   "Update Serial Port"
         Height          =   400
         Left            =   840
         TabIndex        =   72
         Top             =   1020
         Width           =   2295
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   66
         Top             =   240
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1060
         TabIndex        =   65
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1940
         TabIndex        =   64
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2820
         TabIndex        =   63
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   62
         Top             =   600
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1060
         TabIndex        =   61
         Top             =   600
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1940
         TabIndex        =   60
         Top             =   600
         Width           =   900
      End
      Begin VB.OptionButton optPort 
         Caption         =   "Port 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2810
         TabIndex        =   59
         Top             =   600
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "&Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4440
      TabIndex        =   57
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2820
      TabIndex        =   34
      Top             =   3960
      Width           =   1515
      Begin VB.CommandButton cmdUpFive 
         Appearance      =   0  'Flat
         Caption         =   " +5 KHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   780
         TabIndex        =   31
         Top             =   230
         Width           =   615
      End
      Begin VB.CommandButton cmdDnFive 
         Appearance      =   0  'Flat
         Caption         =   " -5 KHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   120
         TabIndex        =   30
         Top             =   230
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1500
      TabIndex        =   27
      Top             =   3960
      Width           =   1215
      Begin VB.OptionButton optSQLOff 
         Caption         =   "SQL Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optSQLOn 
         Caption         =   "SQL On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame6 
      Height          =   795
      Left            =   180
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
      Begin VB.OptionButton optNBOff 
         Caption         =   "NB Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optNBOn 
         Caption         =   "NB On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Height          =   960
      Left            =   4080
      TabIndex        =   47
      Top             =   960
      Width           =   2775
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   0
         Left            =   2400
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   1
         Left            =   2130
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   2
         Left            =   1860
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   3
         Left            =   1260
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   4
         Left            =   975
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   5
         Left            =   690
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   6
         Left            =   405
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTx 
         Height          =   675
         Index           =   7
         Left            =   120
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      Height          =   960
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   2775
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   0
         Left            =   2400
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   1
         Left            =   2130
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   2
         Left            =   1860
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   3
         Left            =   1260
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   4
         Left            =   975
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   5
         Left            =   690
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   6
         Left            =   405
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRx 
         Height          =   675
         Index           =   7
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1191
         _Version        =   393216
         Enabled         =   -1  'True
      End
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "WRCC 7268"
      Height          =   315
      Index           =   11
      Left            =   5640
      TabIndex        =   23
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "WWV 15"
      Height          =   315
      Index           =   10
      Left            =   5640
      TabIndex        =   22
      Top             =   2985
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "WWV 10"
      Height          =   315
      Index           =   9
      Left            =   5640
      TabIndex        =   21
      Top             =   2670
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "160 Meters"
      Height          =   315
      Index           =   8
      Left            =   5640
      TabIndex        =   20
      Top             =   2355
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "80 Meters"
      Height          =   315
      Index           =   7
      Left            =   5640
      TabIndex        =   18
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "40 Meters"
      Height          =   315
      Index           =   6
      Left            =   4440
      TabIndex        =   17
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "30 Meters"
      Height          =   315
      Index           =   5
      Left            =   4440
      TabIndex        =   16
      Top             =   3615
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "20 Meters"
      Height          =   315
      Index           =   4
      Left            =   4440
      TabIndex        =   15
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "17 Meters"
      Height          =   315
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   2985
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "15 Meters"
      Height          =   315
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   2670
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "12 Meters"
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   56
      Top             =   2355
      Width           =   1200
   End
   Begin VB.CommandButton cmdPreset 
      Caption         =   "10 Meters"
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   12
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5760
      TabIndex        =   32
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Emission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   180
      TabIndex        =   1
      Top             =   1980
      Width           =   1215
      Begin VB.OptionButton optEmission 
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   6
         Top             =   1560
         Width           =   900
      End
      Begin VB.OptionButton optEmission 
         Caption         =   "CW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   1245
         Width           =   900
      End
      Begin VB.OptionButton optEmission 
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   930
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optEmission 
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   615
         Width           =   900
      End
      Begin VB.OptionButton optEmission 
         Caption         =   "USB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "RF Gain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1500
      TabIndex        =   9
      Top             =   3000
      Width           =   2835
      Begin MSComctlLib.Slider sldRFGain 
         Height          =   450
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   794
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   10
         Value           =   51
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AF Gain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1500
      TabIndex        =   7
      Top             =   1980
      Width           =   2835
      Begin MSComctlLib.Slider sldAFGain 
         Height          =   450
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   794
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   5
         Value           =   50
      End
   End
   Begin VB.CommandButton btnTxToRx 
      Caption         =   "<<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   660
      Width           =   915
   End
   Begin VB.CommandButton btnRxToTx 
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   360
      Width           =   915
   End
   Begin VB.CommandButton btnRxTxLock 
      Caption         =   "Simplex"
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
      Left            =   3060
      TabIndex        =   11
      Top             =   1200
      Width           =   915
   End
   Begin VB.TextBox txtTransmitter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "29999.999"
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtReceiver 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "29999.999"
      Top             =   360
      Width           =   2775
   End
   Begin MSCommLib.MSComm spMain 
      Left            =   780
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   4800
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   375
      Left            =   4800
      TabIndex        =   130
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Use the mouse wheel as an RIT when the application is highlighted..."
      Height          =   255
      Left            =   180
      TabIndex        =   123
      Top             =   6900
      Width           =   9855
   End
   Begin VB.Line lnUpdate 
      X1              =   180
      X2              =   10020
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Transmit Kilohertz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   36
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Receive Kilohertz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   35
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Project..........Icom Control Panel
' File Name........MAIN.FRM
' File Version.....4/3/01
' Contents.........Main form...
'
' Copyright (c) 2001 - All Rights Reserved
' Victor Poor, W5SMM
' 1208 East River Drive, #302
' Melbourne, FL 32901
'
' PTC II control's by Tom Lafleur, KA6IQA
'

Option Explicit

Public lPort As Long
Dim rdo As New CIcom710
Dim lRxIndex As Long
Dim lTxIndex As Long
Dim bTxSelect As Boolean
Dim bRxTxLock As Boolean
Dim bRxChange As Boolean
Dim bTxChange As Boolean
Dim bNB As Boolean
Dim bSQL As Boolean
Dim bRIT As Boolean
Dim bStepFive As Boolean
Dim bInitialized As Boolean

Private Sub cmdPreset_Click(Index As Integer)
   Select Case Index
      Case 0
         bRxTxLock = False
         txtReceiver.Text = "28000.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
      Case 1
         bRxTxLock = False
         txtReceiver.Text = "24890.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
      Case 2
         bRxTxLock = False
         txtReceiver.Text = "21000.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
      Case 3
         bRxTxLock = False
         txtReceiver.Text = "18068.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
      Case 4
         bRxTxLock = False
         txtReceiver.Text = "14000.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
      Case 5
         bRxTxLock = False
         txtReceiver.Text = "10100.000"
         btnRxTxLock_Click
         rdo.Mode rmCW
         optEmission(3).Value = True
      Case 6
         bRxTxLock = False
         txtReceiver.Text = "7000.000"
         btnRxTxLock_Click
         rdo.Mode rmLSB
         optEmission(1).Value = True
      Case 7
         bRxTxLock = False
         txtReceiver.Text = "3500.000"
         btnRxTxLock_Click
         rdo.Mode rmLSB
         optEmission(1).Value = True
      Case 8
         bRxTxLock = False
         txtReceiver.Text = "1800.000"
         btnRxTxLock_Click
         rdo.Mode rmLSB
         optEmission(1).Value = True
      Case 9
         bRxTxLock = False
         txtReceiver.Text = "10000.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
      Case 10
         bRxTxLock = False
         txtReceiver.Text = "15000.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
     Case 11
         bRxTxLock = False
         txtReceiver.Text = "7268.000"
         btnRxTxLock_Click
         rdo.Mode rmLSB
         optEmission(1).Value = True
     Case 12
         bRxTxLock = False
         txtReceiver.Text = "14300.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
     Case 13
         bRxTxLock = False
         txtReceiver.Text = "14313.000"
         btnRxTxLock_Click
         rdo.Mode rmUSB
         optEmission(0).Value = True
  End Select
  
End Sub

Private Sub btnRxToTx_Click()
   txtTransmitter.Text = txtReceiver.Text
   rdo.SetTransmitter CDbl(txtTransmitter.Text)
   
End Sub

Private Sub btnRxTxLock_Click()
   If bRxTxLock Then
      bRxTxLock = False
      btnRxToTx.Enabled = True
      btnTxToRx.Enabled = True
      btnRxTxLock.Caption = "Duplex"
   Else
      bRxTxLock = True
      txtTransmitter.Text = txtReceiver.Text
      rdo.SetTransmitter CDbl(txtTransmitter.Text)
      btnRxToTx.Enabled = False
      btnTxToRx.Enabled = False
      btnRxTxLock.Caption = "Simplex"
  End If
  
End Sub

Private Sub btnTxToRx_Click()
   txtReceiver.Text = txtTransmitter.Text
   rdo.SetReceiver CDbl(txtReceiver.Text)
   
End Sub

Private Sub cmdClose_Click()

   If spMain.PortOpen Then spMain.PortOpen = False
   Unload frmMain
   
End Sub

Private Sub cmdRecall_Click(Index As Integer)
   Dim lIndex As Long
   Dim lEntry As Long
   
   lEntry = Index
   
   txtReceiver.Text = GetSetting(App.EXEName, "VALUES", "RX" & CStr(lEntry), "10000.000")
   txtTransmitter.Text = GetSetting(App.EXEName, "VALUES", "TX" & CStr(lEntry), "10000.000")
   bRxTxLock = CBool(GetSetting(App.EXEName, "VALUES", "LOCK" & CStr(lEntry), "TRUE"))
  
  If bRxTxLock Then
      btnRxToTx.Enabled = False
      btnTxToRx.Enabled = False
      btnRxTxLock.Caption = "Simplex"
   Else
      btnRxToTx.Enabled = True
      btnTxToRx.Enabled = True
      btnRxTxLock.Caption = "Duplex"
   End If
   
   optEmission(CInt(GetSetting(App.EXEName, "VALUES", "EMISSION" & CStr(lEntry), "0"))).Value = True
   optEmission_Click (CInt(GetSetting(App.EXEName, "VALUES", "EMISSION" & CStr(lEntry), "0")))
   
   sldAFGain.Value = CInt(GetSetting(App.EXEName, "VALUES", "AFGAIN" & CStr(lEntry), "70"))
   sldRFGain.Value = CInt(GetSetting(App.EXEName, "VALUES", "RFGAIN" & CStr(lEntry), "100"))
   
   bNB = CBool(GetSetting(App.EXEName, "VALUES", "NB" & CStr(lEntry), "FALSE"))
   optNBOn.Value = bNB
   rdo.NoiseBlank bNB
   
   bSQL = CBool(GetSetting(App.EXEName, "VALUES", "SQL" & CStr(lEntry), "FALSE"))
   optSQLOn.Value = bSQL
   rdo.Squelch bSQL
   
   bInitialized = True

End Sub

Private Sub cmdRecall_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   Static bInitialized As Boolean
   Static lIndex As Long
   
   If Not bInitialized Then
      bInitialized = True
      lIndex = -1
   End If
   
   If lIndex <> Index Then
      lIndex = Index
      cmdRecall(Index).ToolTipText = _
         GetSetting(App.EXEName, "VALUES", "RX" & CStr(Index), "10000.000") & "/" & _
         GetSetting(App.EXEName, "VALUES", "TX" & CStr(Index), "10000.000")
   End If
   
End Sub

Private Sub cmdSetSerialPort_Click()
   Dim lIndex As Long
   
   For lIndex = 0 To 7
      If optPort(lIndex).Value = True Then
         SaveSetting App.EXEName, "VALUES", "PORT", lIndex
         lPort = lIndex
         rdo.OpenPort spMain, lPort
         If Not rdo.OpenRadio Then
            bInitialized = False
            MsgBox "No radio found on this port..."
         Else
            lnUpdate.Visible = 0
            Height = 5295
            GetSettings
         End If
      End If
   Next lIndex
   
End Sub

Private Sub cmdSetup_Click()
   optPort(lPort).Value = True
   If Height < 6000 Then
      lnUpdate.Visible = True
      Height = 7590
   Else
      lnUpdate.Visible = False
      Height = 5295
   End If
   
End Sub

Private Sub cmdStore_Click(Index As Integer)
   Dim lIndex As Long
   Dim lEntry As Long
   
   If bInitialized Then
      lEntry = Index
   
      SaveSetting App.EXEName, "VALUES", "RX" & CStr(lEntry), txtReceiver.Text
      SaveSetting App.EXEName, "VALUES", "TX" & CStr(lEntry), txtTransmitter.Text
      SaveSetting App.EXEName, "VALUES", "LOCK" & CStr(lEntry), CStr(bRxTxLock)
      
      For lIndex = 0 To 5
         If optEmission(lIndex).Value = True Then
            SaveSetting App.EXEName, "VALUES", "EMISSION" & CStr(lEntry), lIndex
            Exit For
         End If
      Next lIndex
      
      SaveSetting App.EXEName, "VALUES", "AFGAIN" & CStr(lEntry), sldAFGain.Value
      SaveSetting App.EXEName, "VALUES", "RFGAIN" & CStr(lEntry), sldRFGain.Value
      SaveSetting App.EXEName, "VALUES", "NB" & CStr(lEntry), CStr(bNB)
      SaveSetting App.EXEName, "VALUES", "SQL" & CStr(lEntry), CStr(bSQL)
           
   Else
      MsgBox "Radio must be connected..."
   End If
   
End Sub

Private Sub cmdStore_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   Static bInitialized As Boolean
   Static lIndex As Long
   
   If Not bInitialized Then
      bInitialized = True
      lIndex = -1
   End If
   
   If lIndex <> Index Then
      lIndex = Index
      cmdStore(Index).ToolTipText = _
         GetSetting(App.EXEName, "VALUES", "RX" & CStr(Index), "10000.000") & "/" & _
         GetSetting(App.EXEName, "VALUES", "TX" & CStr(Index), "10000.000")
   End If
   
End Sub

Private Sub Form_Activate()
Static bStarted As Boolean
Dim i As Integer
If Not bStarted Then
    bStarted = True
      
    ' Enable serial ports if they are on this system
    For i = 1 To 8
        If EnumSerPorts(i) = 0 Then
            optPort(i - 1).Enabled = False
        Else
        optPort(i - 1).Enabled = True
        End If
    Next i

    PTC_Select.Value = GetSetting(App.EXEName, "VALUES", "PTCFLAG", "0")
    PTC_Select_Click
    
    PTC_Baud.ListIndex = CInt(GetSetting(App.EXEName, "VALUES", "PTCBAUD", "6"))
    
    lPort = CLng(GetSetting(App.EXEName, "VALUES", "PORT", "0"))
      
    optPort(lPort).Value = True
    rdo.OpenPort spMain, lPort

    RadioType(CLng(GetSetting(App.EXEName, "VALUES", "RADIOTYPE", "0"))).Value = True
    RadioType_Click (CLng(GetSetting(App.EXEName, "VALUES", "RADIOTYPE", "0")))

    hMain = Me.hwnd
    Hook
    lblVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End If
   
End Sub

Private Sub Form_Load()
   If App.PrevInstance Then End
   Width = 10290
   Height = 5295
   lRxIndex = 2
   lTxIndex = 2
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unhook
   If rdo.CloseRadio Then SaveSettings
   If spMain.PortOpen Then spMain.PortOpen = False
   
End Sub

Private Sub optEmission_Click(Index As Integer)
   Select Case Index
      Case 0
         rdo.Mode rmUSB
      Case 1
         rdo.Mode rmLSB
      Case 2
         rdo.Mode rmNBDP
      Case 3
         rdo.Mode rmCW
      Case 4
         rdo.Mode rmAM
   End Select
   
End Sub

Private Sub optNBOff_Click()
   rdo.NoiseBlank False
   bNB = False
   
End Sub

Private Sub optNBOn_Click()
   rdo.NoiseBlank True
   bNB = True
   
End Sub

Private Sub optSQLOff_Click()
   rdo.Squelch False
   bSQL = False
   
End Sub

Private Sub optSQLOn_Click()
   rdo.Squelch True
   bSQL = True
   
End Sub

Private Sub PTC_Baud_click()
    ' Select new Baud rate if PTC II Modem
    rdo.PTCBaud = PTC_Baud.Text
    SaveSetting App.EXEName, "VALUES", "PTCBAUD", PTC_Baud.ListIndex
    
End Sub

Private Sub PTC_Select_Click()

    ' Selects Radio Control via PTC II port
    If PTC_Select.Value = vbChecked Then
    
        PTC_Baud.Enabled = True
        rdo.PTC = True
        SaveSetting App.EXEName, "VALUES", "PTCFLAG", "1"
        
    ElseIf PTC_Select.Value = vbUnchecked Then
    
        PTC_Baud.Enabled = False
        rdo.PTC = False
        SaveSetting App.EXEName, "VALUES", "PTCFLAG", "0"
    End If
        
End Sub

Private Sub radioAddress_keyDown(KeyCode As Integer, Shift As Integer)
        
    Select Case KeyCode
    Case vbKeyReturn
    
        If IsNumeric(radioAddress.Text) Then
            radioAddress.Text = Format$(radioAddress.Text, "00")
            rdo.RadioType = Left$(radioAddress.Text, 2)
            SaveSetting App.EXEName, "VALUES", "RADIOADDRESS", radioAddress.Text
            If rdo.OpenRadio Then
                GetSettings
            Else
                MsgBox "No Radio Found . . ."
            End If
        Else
            Beep ' Error
            
        End If
    End Select
      
End Sub

Private Sub RadioType_Click(Index As Integer)
    ' Selects Radio Type
     Select Case Index
      Case 0
         ' Icom M700Pro
         rdo.RadioType = "02"
         Caption = "Icom M700Pro Control Panel"
         radioAddress.Text = "02"
         radioAddress.Locked = True
      Case 1
         ' Icom M710
         rdo.RadioType = "01"
         Caption = "Icom M710 Control Panel"
         radioAddress.Text = "01"
         radioAddress.Locked = True
      Case 2
         ' Icom M710RT
         rdo.RadioType = "03"
         Caption = "Icom M710RT Control Panel"
         radioAddress.Text = "03"
         radioAddress.Locked = True
       Case Else
         ' Other Radio Address
         radioAddress.Text = GetSetting(App.EXEName, "VALUES", "RADIOADDRESS", "00")
         rdo.RadioType = radioAddress.Text
         Caption = "Icom Control Panel"
         radioAddress.Locked = False
         
         ' Don't make changes until user updates the address with a CR
         Exit Sub
         
       End Select
       SaveSetting App.EXEName, "VALUES", "RADIOTYPE", Format$(Index, "0")
       
      If rdo.OpenRadio Then   ' Opens port
         GetSettings
      Else
         MsgBox "No Radio Found . . ."
      End If
      
End Sub

Private Sub sldAFGain_Change()
   rdo.AudioGain 10 + (sldAFGain.Value / 2)
   btnRxTxLock.SetFocus
   
End Sub

Private Sub sldRFGain_Change()
   rdo.RFGain sldRFGain.Value
   btnRxTxLock.SetFocus
   
End Sub

Private Sub tmrMain_Timer()
   Dim dValue As Double
   
   dValue = CDbl(txtReceiver.Text)
   dValue = dValue + (lStep * 0.01)
   lStep = 0#
   If dValue < 20# Then dValue = 20#
   bRIT = True
   txtReceiver.Text = FormatQRG(dValue)
   bRIT = False
   
End Sub

Private Sub txtReceiver_Change()
   rdo.SetReceiver CDbl(txtReceiver.Text)
   If bRxTxLock And Not bRIT Then
      txtTransmitter.Text = txtReceiver.Text
   End If
   bTxChange = False
   
End Sub

Private Sub txtTransmitter_Change()
   rdo.SetTransmitter CDbl(txtTransmitter.Text)
   If bRxTxLock Then
      txtReceiver.Text = txtTransmitter.Text
   End If
   bRxChange = False
   
End Sub

Private Sub cmdDnFive_Click()
   Dim dValue As Double
   
   dValue = CDbl(txtReceiver.Text)
   dValue = dValue - 5
   
   If dValue < 20# Then dValue = 20#
   txtReceiver.Text = FormatQRG(dValue)
   bStepFive = True
   
End Sub

Private Sub cmdUpFive_Click()
   Dim dValue As Double
   
   dValue = CDbl(txtReceiver.Text)
   dValue = dValue + 5
   If dValue > 29700# Then dValue = 29700#
   txtReceiver.Text = FormatQRG(dValue)
   bStepFive = True
   
End Sub

Private Sub udRx_DownClick(Index As Integer)
   Dim dValue As Double
   
   bStepFive = False
   bTxSelect = False
   lRxIndex = Index
   If bRxTxLock Then bRxChange = True
   dValue = CDbl(txtReceiver.Text)
   Select Case Index
      Case 0
         dValue = dValue - 0.001
      Case 1
         dValue = dValue - 0.01
      Case 2
         dValue = dValue - 0.1
      Case 3
         dValue = dValue - 1
      Case 4
         dValue = dValue - 10
      Case 5
         dValue = dValue - 100
      Case 6
         dValue = dValue - 1000
      Case 7
         dValue = dValue - 10000
   End Select
   
   If dValue < 20# Then dValue = 20#
   txtReceiver.Text = FormatQRG(dValue)
   
End Sub

Private Sub udRx_UpClick(Index As Integer)
   Dim dValue As Double
   
   bStepFive = False
   bTxSelect = False
   lRxIndex = Index
   If bRxTxLock Then bRxChange = True
   dValue = CDbl(txtReceiver.Text)
   Select Case Index
      Case 0
         dValue = dValue + 0.001
      Case 1
         dValue = dValue + 0.01
      Case 2
         dValue = dValue + 0.1
      Case 3
         dValue = dValue + 1
      Case 4
         dValue = dValue + 10
      Case 5
         dValue = dValue + 100
      Case 6
         dValue = dValue + 1000
      Case 7
         dValue = dValue + 10000
   End Select
   
   If dValue > 29700# Then dValue = 29700#
   txtReceiver.Text = FormatQRG(dValue)
   
End Sub

Private Sub udTx_DownClick(Index As Integer)
   Dim dValue As Double
   
   bStepFive = False
   bTxSelect = True
   lTxIndex = Index
   If bRxTxLock Then bTxChange = True
   dValue = CDbl(txtTransmitter.Text)
   Select Case Index
      Case 0
         dValue = dValue - 0.001
      Case 1
         dValue = dValue - 0.01
      Case 2
         dValue = dValue - 0.1
      Case 3
         dValue = dValue - 1
      Case 4
         dValue = dValue - 10
      Case 5
         dValue = dValue - 100
      Case 6
         dValue = dValue - 1000
      Case 7
         dValue = dValue - 10000
   End Select
   
   If dValue < 20# Then dValue = 20#
   txtTransmitter.Text = FormatQRG(dValue)
   
End Sub

Private Sub udTx_UpClick(Index As Integer)
   Dim dValue As Double
   
   bStepFive = False
   bTxSelect = True
   lTxIndex = Index
   If bRxTxLock Then bTxChange = True
   dValue = CDbl(txtTransmitter.Text)
   Select Case Index
      Case 0
         dValue = dValue + 0.001
      Case 1
         dValue = dValue + 0.01
      Case 2
         dValue = dValue + 0.1
      Case 3
         dValue = dValue + 1
      Case 4
         dValue = dValue + 10
      Case 5
         dValue = dValue + 100
      Case 6
         dValue = dValue + 1000
      Case 7
         dValue = dValue + 10000
   End Select
   
   If dValue > 29700# Then dValue = 29700#
   txtTransmitter.Text = FormatQRG(dValue)
   
End Sub

Private Function FormatQRG(dValue As Double) As String
   Dim sResult As String
   
   sResult = Format(dValue, "###00.000")
   
   Select Case Len(sResult)
      Case 6
         FormatQRG = "   " & sResult
      Case 7
         FormatQRG = "  " & sResult
      Case 8
         FormatQRG = " " & sResult
      Case 9
         FormatQRG = sResult
   End Select
   
End Function

Public Sub GetSettings()
   Top = CInt(GetSetting(App.EXEName, "VALUES", "TOP", "0"))
   Left = CInt(GetSetting(App.EXEName, "VALUES", "LEFT", "0"))
   txtReceiver.Text = GetSetting(App.EXEName, "VALUES", "RX", "10000.000")
   'rdo.SetReceiver CDbl(txtReceiver.Text)
   txtTransmitter.Text = GetSetting(App.EXEName, "VALUES", "TX", "10000.000")
   'rdo.SetTransmitter CDbl(txtTransmitter.Text)
   bRxTxLock = CBool(GetSetting(App.EXEName, "VALUES", "LOCK", "TRUE"))
   
   If bRxTxLock Then
      btnRxToTx.Enabled = False
      btnTxToRx.Enabled = False
      btnRxTxLock.Caption = "Simplex"
   Else
      btnRxToTx.Enabled = True
      btnTxToRx.Enabled = True
      btnRxTxLock.Caption = "Duplex"
   End If
   
   optEmission(CInt(GetSetting(App.EXEName, "VALUES", "EMISSION", "0"))).Value = True
   optEmission_Click (CInt(GetSetting(App.EXEName, "VALUES", "EMISSION", "0")))
   
   sldAFGain.Value = CInt(GetSetting(App.EXEName, "VALUES", "AFGAIN", "70"))
   sldRFGain.Value = CInt(GetSetting(App.EXEName, "VALUES", "RFGAIN", "100"))
   
   bNB = CBool(GetSetting(App.EXEName, "VALUES", "NB", "FALSE"))
   optNBOn.Value = bNB
   rdo.NoiseBlank bNB
   
   bSQL = CBool(GetSetting(App.EXEName, "VALUES", "SQL", "FALSE"))
   optSQLOn.Value = bSQL
   rdo.Squelch bSQL
   
   bInitialized = True
   
End Sub

Private Sub SaveSettings()
   Dim lIndex As Long
   
   If bInitialized Then
      SaveSetting App.EXEName, "VALUES", "TOP", Top
      SaveSetting App.EXEName, "VALUES", "LEFT", Left
      SaveSetting App.EXEName, "VALUES", "RX", txtReceiver.Text
      SaveSetting App.EXEName, "VALUES", "TX", txtTransmitter.Text
      SaveSetting App.EXEName, "VALUES", "LOCK", CStr(bRxTxLock)
      
      For lIndex = 0 To 5
         If optEmission(lIndex).Value = True Then
            SaveSetting App.EXEName, "VALUES", "EMISSION", lIndex
            Exit For
         End If
      Next lIndex
      
      SaveSetting App.EXEName, "VALUES", "AFGAIN", sldAFGain.Value
      SaveSetting App.EXEName, "VALUES", "RFGAIN", sldRFGain.Value
      SaveSetting App.EXEName, "VALUES", "NB", CStr(bNB)
      SaveSetting App.EXEName, "VALUES", "SQL", CStr(bSQL)
       
      End If
      
End Sub

Public Function EnumSerPorts(port As Integer) As Long
'returns non-zero value if the port exists
Dim cc As COMMCONFIG, ccsize As Long
    ccsize = LenB(cc)     'gets the size of COMMCONFIG structure
    EnumSerPorts = GetDefaultCommConfig("COM" + Trim(Str(port)) + Chr(0), cc, ccsize)

End Function

