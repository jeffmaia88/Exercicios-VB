VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   5880
   ClientTop       =   2535
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtLista 
      BackColor       =   &H8000000F&
      Height          =   4815
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de nomes"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtNome.Text = ""
    
End Sub
Private Sub Form_Activate()
    txtNome.SetFocus
End Sub
Private Sub cmdAdicionar_Click()
    Dim nome As String
   
    nome = Trim(txtNome.Text)
    
    If nome = "" Then
        MsgBox "Você Deve digitar um nome.", vbInformation, "Atenção"
        txtNome.SetFocus
        Exit Sub
    End If
    
    
    txtLista.Text = txtLista.Text & nome & vbCrLf
    MsgBox "Nome incluído com sucesso!", vbInformation, "Sucesso"
    
    txtNome.Text = ""
    txtNome.SetFocus
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim resp As VbMsgBoxResult

    resp = MsgBox("Deseja realmente sair?" & vbCrLf & _
                  "Todos os dados da lista serão perdidos.", _
                  vbYesNo + vbQuestion + vbDefaultButton2, _
                  "Confirmar saída")

    If resp = vbNo Then
        
        Cancel = True
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdAdicionar_Click
    End If
End Sub

Private Sub cmdLimpar_Click()
    txtLista.Text = ""
    txtNome.Text = ""
    txtNome.SetFocus

End Sub

