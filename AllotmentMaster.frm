VERSION 5.00
Begin VB.Form RoomNumbers 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Vehicles Passed "
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AllotmentMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7620
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   51
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "V I P Vehicles"
      Height          =   375
      Left            =   4440
      TabIndex        =   55
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Billed for One way"
      Height          =   375
      Left            =   4440
      TabIndex        =   54
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Billed for Two way"
      Height          =   375
      Left            =   4440
      TabIndex        =   53
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehilces Paased on "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   52
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "RoomNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RoomNrVar, RINoVar As Integer
Dim ODateVar, TodayVar, DiffDateVar, MyTimeVar As Date
Private Sub Command1_Click(Index As Integer)
RoomNo = Command1(Index).Caption
If Not Command1(Index).BackColor = &H80FF80 Then
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select altNo from allotmentDb where rNo=" & RoomNo & " and checkout = 'N'", conn
    If Rs.EOF = False Then
    AltNoVar = Rs(0)
    End If
End If
AllotmentForm.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
ColorChange
End Sub

Private Sub Form_Load()
    Me.Height = 8000
    Me.Width = 7740
    Me.Top = 0
    Me.Left = 0
    
End Sub
Public Sub ColorChange()
Dim RIRS1 As New ADODB.Recordset
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select RNo from roomdb order by RNo", conn
    TodayVar = Date
    MyTimeVar = Time
    I = 0
    Do While Not RIRS.EOF
    Command1(I).Caption = RIRS(0)
    Command1(I).Visible = True
    RINoVar = RIRS(0)
    If RIRS1.State = 1 Then RIRS1.Close
    RIRS1.Open "Select checkOut,DateOut from Allotmentdb where RNo=" & Val(RINoVar) & " and checkout='N'", conn
    If RIRS1.EOF = False Then
        Command1(I).BackColor = &HFF& 'RED
        If RIRS1(1) <= TodayVar Then
            Command1(I).BackColor = &H80FFFF 'Yellow
        End If
    Else
    Command1(I).BackColor = &H80FF80 'Green
    End If
    RIRS.MoveNext
    I = I + 1
    Loop
End Sub
