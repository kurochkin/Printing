VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   10230
   ClientTop       =   6255
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9270
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   1680
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GetResult"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim portListener As Variant

Private Sub Command1_Click()
   portListener.StartReading
End Sub

Private Sub Command2_Click()
    portListener.StopReading
End Sub

Private Sub Command3_Click()
    Dim res
        res = portListener.ReadedData
        Text1.Text = res
End Sub

Private Sub Form_Load()
   'Set portListener = CreateObject("SerialPortDataProcessor.PrintLabelProcessor")
   Dim portListener As New PrintLabelProcessor
   portListener.Init
   'portListener.PrintLabel 100
   portListener.PrintLabelV2 100
End Sub
