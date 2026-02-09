VERSION 5.00
Begin VB.Form frmMQTTDemo 
   Caption         =   "MQTT Demo"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraClient 
      Caption         =   "MQTT Client"
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   10335
      Begin VB.CommandButton cmdClientDisconnect 
         Caption         =   "Disconnect"
         Height          =   375
         Left            =   8880
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdClientConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   7560
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtClientPort 
         Height          =   375
         Left            =   6360
         TabIndex        =   20
         Text            =   "1883"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtClientHost 
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtSubTopic 
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Text            =   "test/topic"
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton cmdUnsubscribe 
         Caption         =   "Unsubscribe"
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubscribe 
         Caption         =   "Subscribe"
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtPubPayload 
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Text            =   "Hello MQTT!"
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txtPubTopic 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Text            =   "test/topic"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton cmdPublish 
         Caption         =   "Publish"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblClientStatus 
         Caption         =   "Status: Disconnected"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Port:"
         Height          =   255
         Left            =   5640
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Host:"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Topic:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Payload:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Topic:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "MQTT Server"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.ListBox lstClients 
         Height          =   1620
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox txtLog 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2160
         Width           =   9975
      End
      Begin VB.CommandButton cmdStopServer 
         Caption         =   "Stop"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdStartServer 
         Caption         =   "Start"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtServerPort 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Text            =   "1883"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtServerStatus 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Stopped"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblServerStatus 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblStatus 
         Caption         =   "Server Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMQTTDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' frmMQTTDemo - MQTT Server/Client Demo Form
'=========================================================================
Option Explicit

Private WithEvents m_oServer As cMQTTServer
Attribute m_oServer.VB_VarHelpID = -1
Private WithEvents m_oClient As cMQTTClient
Attribute m_oClient.VB_VarHelpID = -1

Private Sub Form_Load()
    Set m_oServer = New cMQTTServer
    Set m_oClient = New cMQTTClient
    
    cmdStopServer.Enabled = False
    cmdClientDisconnect.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oClient = Nothing
    Set m_oServer = Nothing
End Sub

'=========================================================================
' Server Events
'=========================================================================

Private Sub m_oServer_ClientConnected(ClientId As String, Session As cMQTTSession)
    LogMessage "Client connected: " & ClientId
    UpdateClientList
End Sub

Private Sub m_oServer_ClientDisconnected(ClientId As String)
    LogMessage "Client disconnected: " & ClientId
    UpdateClientList
End Sub

Private Sub m_oServer_MessageReceived(ClientId As String, Topic As String, Payload As Variant, QoS As MqttQoS)
    Dim sPayload As String
    If VarType(Payload) = vbByte + vbArray Then
        sPayload = "[Binary data]"
    Else
        sPayload = CStr(Payload)
    End If
    LogMessage "Message from " & ClientId & " on " & Topic & ": " & sPayload
End Sub

Private Sub m_oServer_Subscribe(ClientId As String, TopicFilter As String, QoS As MqttQoS)
    LogMessage "Client " & ClientId & " subscribed to " & TopicFilter & " (QoS " & QoS & ")"
End Sub

Private Sub m_oServer_Unsubscribe(ClientId As String, TopicFilter As String)
    LogMessage "Client " & ClientId & " unsubscribed from " & TopicFilter
End Sub

Private Sub m_oServer_ErrorEvent(Number As Long, Description As String)
    LogMessage "Server Error " & Number & ": " & Description
End Sub

'=========================================================================
' Client Events
'=========================================================================

Private Sub m_oClient_Connected()
    lblClientStatus.Caption = "Status: Connected"
    cmdClientConnect.Enabled = False
    cmdClientDisconnect.Enabled = True
    LogMessage "Client connected to server"
End Sub

Private Sub m_oClient_Disconnected()
    lblClientStatus.Caption = "Status: Disconnected"
    cmdClientConnect.Enabled = True
    cmdClientDisconnect.Enabled = False
    LogMessage "Client disconnected from server"
End Sub

Private Sub m_oClient_MessageArrived(Topic As String, Payload As Variant, QoS As MqttQoS, Retain As Boolean)
    Dim sPayload As String
    If VarType(Payload) = vbByte + vbArray Then
        sPayload = "[Binary data]"
    Else
        sPayload = CStr(Payload)
    End If
    LogMessage "Message arrived on " & Topic & " (QoS " & QoS & ", Retain=" & Retain & "): " & sPayload
End Sub

Private Sub m_oClient_Subscribed(TopicFilter As String, QoS As MqttQoS)
    LogMessage "Subscribed to " & TopicFilter & " (QoS " & QoS & ")"
End Sub

Private Sub m_oClient_Unsubscribed(TopicFilter As String)
    LogMessage "Unsubscribed from " & TopicFilter
End Sub

Private Sub m_oClient_ErrorEvent(Number As Long, Description As String)
    LogMessage "Client Error " & Number & ": " & Description
End Sub

'=========================================================================
' UI Event Handlers
'=========================================================================

Private Sub cmdStartServer_Click()
    Dim lPort As Long
    lPort = CLng(txtServerPort.Text)
    
    If m_oServer.StartServer(lPort) Then
        txtServerStatus.Text = "Running on port " & lPort
        cmdStartServer.Enabled = False
        cmdStopServer.Enabled = True
        LogMessage "MQTT Server started on port " & lPort
    Else
        MsgBox "Failed to start server!", vbExclamation
    End If
End Sub

Private Sub cmdStopServer_Click()
    m_oServer.StopServer
    txtServerStatus.Text = "Stopped"
    cmdStartServer.Enabled = True
    cmdStopServer.Enabled = False
    lstClients.Clear
    LogMessage "MQTT Server stopped"
End Sub

Private Sub cmdClientConnect_Click()
    m_oClient.ClientId = "VBClient_" & Timer
    
    If m_oClient.Connect(txtClientHost.Text, CLng(txtClientPort.Text)) Then
        ' Connected event will fire
    Else
        MsgBox "Failed to connect!", vbExclamation
    End If
End Sub

Private Sub cmdClientDisconnect_Click()
    m_oClient.Disconnect
End Sub

Private Sub cmdSubscribe_Click()
    If Len(txtSubTopic.Text) > 0 Then
        m_oClient.Subscribe txtSubTopic.Text, mqttQoS0
    End If
End Sub

Private Sub cmdUnsubscribe_Click()
    If Len(txtSubTopic.Text) > 0 Then
        m_oClient.Unsubscribe txtSubTopic.Text
    End If
End Sub

Private Sub cmdPublish_Click()
    If Len(txtPubTopic.Text) > 0 Then
        m_oClient.Publish txtPubTopic.Text, txtPubPayload.Text, mqttQoS0, False
    End If
End Sub

'=========================================================================
' Helper Methods
'=========================================================================

Private Sub LogMessage(ByVal sMsg As String)
    txtLog.Text = txtLog.Text & Format(Now, "hh:nn:ss") & " - " & sMsg & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub UpdateClientList()
    ' Note: This is a simplified version. In a real implementation,
    ' you would expose the client list from cMQTTServer
    lstClients.Clear
    lstClients.AddItem "Connected clients: " & m_oServer.ClientCount
End Sub
