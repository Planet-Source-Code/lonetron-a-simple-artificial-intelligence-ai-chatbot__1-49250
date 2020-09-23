VERSION 5.00
Begin VB.Form tanuChatbotForm 
   Caption         =   "Transplantable Artificial Neurological Units"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton createbot 
      Caption         =   "Don't have an TANU chatbot? Click here to create online at: http://www.p2bconsortium.com/sss/CreateBot.aspx"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4935
   End
   Begin VB.TextBox txtInitialState 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox txtOutput 
      Height          =   3735
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton cmdForceIntoState 
      Height          =   1695
      Left            =   0
      Picture         =   "tanuChatbotForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtBotPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtBotName 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Initial State"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Bot Password"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Bot Name"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "tanuChatbotForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdForceIntoState_Click()
    If txtBotName.Text = "" Or txtBotPassword.Text = "" Or txtInitialState.Text = "" Then
        MsgBox "You must enter a valid bot name, " & _
            "password and initial state. You can create a " & _
            "bot and set its password and default state at " & _
            "http://www.p2bconsortium.com/sss/createbot.aspx"
    Else
        'Open an HTTP GET to the Transplantable Artificial Neurological Units (TANU) public API
        'to force the chatbot into an initial state
        Dim oTANU As New MSXML2.XMLHTTP
        oTANU.open "GET", _
            "http://web.p2bconsortium.com:81/sss/sss.asmx/resetCurrentState?strBotname=" & _
            txtBotName.Text & "&strStateName=" & txtInitialState.Text & "&strPassword=" & _
            txtBotPassword.Text, False, "", ""
        oTANU.send
        
        'paser the response from TANU to make sure that there was no error
        If InStr(1, oTANU.responseText, "xsi:nil", 1) = 0 Then
            MsgBox "Unable to force bot into the state called '" & _
                txtInitialState.Text & "' Make sure that the state exists. " & _
                "You can create states with the Rapid Bot Trainer at " & _
                "http://www.p2bconsortium.com/sss/rbt.aspx"
        End If
    End If
End Sub

Private Sub cmdSend_Click()
    If txtBotName.Text = "" Or txtBotPassword.Text = "" Or txtInput.Text = "" Then
        MsgBox "You must enter a valid bot name, password and input text. " & _
            "You can create a bot and set its password and default state at " & _
            "http://www.p2bconsortium.com/sss/createbot.aspx"
    Else
        'Display the user's message
        txtOutput.Text = txtOutput.Text & Chr(13) & Chr(10) & "You: " & txtInput.Text
        txtOutput.SelStart = Len(txtOutput.Text)
        
        'Open an HTTP GET to the Transplantable Artificial Neurological Units (TANU) public API
        'to send an event to the bot and get its response action
        Dim oTANU As New MSXML2.XMLHTTP
        oTANU.open "GET", _
            "http://web.p2bconsortium.com:81/sss/sss.asmx/sendEventGetAction?strBotname=" & _
            txtBotName.Text & "&strEvent=" & txtInput.Text & "&strPassword=" & _
            txtBotPassword.Text, False, "", ""
        oTANU.send
        
        'Paser the response from TANU to make sure that there was no error
        If InStr(1, oTANU.responseText, "p2b.vathix.com", 1) = 0 Then
            MsgBox "Unable to send event to bot. Make sure your password " & _
            "is correct and that your event does not include '<', '>', or '§'"
        Else
            'If a transition did occur. (if no transition happened or the new
            'state has no action TANU returns "no transition from this state")
            If "no transition from this state" <> _
            oTANU.responseXML.documentElement.selectSingleNode("//string").Text Then
                'The chatbot did return a message so display it
                txtOutput.Text = txtOutput.Text & Chr(13) & Chr(10) & _
                    "Bot: " & oTANU.responseXML.documentElement.selectSingleNode("//string").Text
                txtOutput.SelStart = Len(txtOutput.Text)
            End If
        End If
    End If
End Sub

Private Sub createbot_Click()
    Call ShellExecute(Me.hWnd, "Open", "http://www.p2bconsortium.com/sss/CreateBot.aspx", vbNullString, vbNullString, vbNormal)
End Sub
