VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   6225
   ClientTop       =   2790
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6585
   Begin VB.CommandButton Command4 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6375
      Begin VB.CheckBox Check1 
         Caption         =   "Especial"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço completo"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Idade"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   0
      Top             =   2880
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoEdicao As Boolean
Dim idAtual As Long
Dim Conn As ADODB.Connection

Private Sub Form_Load()
    AbrirConexao
    CarregarLista
        
    Frame1.Visible = False
    List1.Visible = True
    Command4.Visible = False
    Command5.Visible = False
    
    Text1.MaxLength = 50
    Text2.MaxLength = 3
    Text3.MaxLength = 250

    
End Sub

Private Sub AbrirConexao()

    On Error GoTo Trata
    
    If Conn Is Nothing Then
        Set Conn = New ADODB.Connection
    End If
    
    If Conn.State = adStateOpen Then
        Exit Sub
    End If
    

    Conn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=DEV_MAIA\PDVNET;" & _
        "Initial Catalog=Exemplos;" & _
        "User ID=sa;" & _
        "Password=SENHADAPDV;"

    Conn.Open
    Exit Sub
    
Trata:
    MsgBox "Erro ao abrir conexão: " & Err.Description, vbCritical, "Erro de conexão"
    
End Sub



Private Sub CarregarLista()
    Dim rs As New ADODB.Recordset
    Dim texto As String
    
    List1.Clear
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT Id,Nome,Idade, Endereco, Especial FROM Cliente ORDER By Nome", Conn
    
     While Not rs.EOF

        texto = rs!Id & " - " & rs!nome & " - " & rs!idade & " - " & rs!endereco

        If rs!especial Then
            texto = texto & " - Especial"
        End If

        List1.AddItem texto
        List1.ItemData(List1.NewIndex) = rs!Id

        rs.MoveNext
    Wend
    
    rs.Close
    
    Set rs = Nothing
   
End Sub

Private Sub Command1_Click()

modoEdicao = False
idAtual = 0

List1.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = True

Frame1.Visible = True

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Check1.Value = vbUnchecked

Text1.SetFocus


End Sub

Private Sub Command2_Click()

    Dim selId As Long
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    If List1.ListIndex = -1 Then
        MsgBox "Selecione um registro na lista para editar.", vbInformation
        Exit Sub
    End If
    
    modoEdicao = True
    
    selId = List1.ItemData(List1.ListIndex)
    idAtual = selId
    
    sql = "SELECT Nome, Idade, Endereco, Especial FROM Cliente Where Id = " & selId
    
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        Text1.Text = rs!nome & ""
        Text2.Text = rs!idade & ""
        Text3.Text = rs!endereco & ""
        
         If rs!especial = True Then
            Check1.Value = vbChecked
        Else
            Check1.Value = vbUnchecked
        End If
    End If

    rs.Close
    
    List1.Visible = False
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    
    Frame1.Visible = True
    Command4.Visible = True
    Command5.Visible = True
    
    Text1.SetFocus

End Sub

Private Sub Command3_Click()
    Dim selId As Long
    Dim resp As VbMsgBoxResult
    Dim sql As String
    
    If List1.ListIndex = -1 Then
        MsgBox "Selecione um registro para excluir.", vbInformation
        Exit Sub
    End If
    
    selId = List1.ItemData(List1.ListIndex)
    resp = MsgBox("Deseja realmente excluir este regsitro?", vbYesNo + vbQuestion, "Confirmar exclusão")
    
    If resp = vbNo Then
        Exit Sub
    End If
    
    sql = "DELETE FROM Cliente Where Id = " & selId
    
    Conn.Execute sql
    
    MsgBox "Registro excluído com Sucesso", vbInformation
    
    CarregarLista

End Sub

Private Sub Command4_Click()

    Dim sql As String
    Dim nome As String
    Dim idade As String
    Dim endereco As String
    Dim especial As Integer

    nome = Trim(Text1.Text)
    idade = Trim(Text2.Text)
    endereco = Trim(Text3.Text)

    If Check1.Value = vbChecked Then
        especial = 1
    Else
        especial = 0
    End If


    If nome = "" Then
        MsgBox "O Campo Nome é Obrigatório", vbInformation
        Text1.SetFocus
        Exit Sub
    End If

    If idade = "" Then
        MsgBox "O Campo Idade é Obrigatório"
        Text2.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(idade) Then
        MsgBox "O Campo idade deve ser numérico.", vbInformation
        Text2.SetFocus
        Exit Sub
    End If
    
    If endereco = "" Then
        MsgBox "O Campo Endereco é obrigatorio.", vbInformation
        Text3.SetFocus
        Exit Sub
    End If
    

    If modoEdicao = True Then
        sql = "UPDATE Cliente SET " & _
              "Nome = '" & Replace(nome, "'", "''") & "', " & _
              "Idade = " & idade & ", " & _
              "Endereco = '" & Replace(endereco, "'", "''") & "', " & _
              "Especial = " & especial & _
              " WHERE Id = " & idAtual

        Conn.Execute sql

        MsgBox "Registro atualizado com sucesso!", vbInformation
    Else

        sql = "INSERT into Cliente(Nome, Idade, Endereco, Especial) VALUES(" & _
            "'" & Replace(nome, "'", "''") & "', " & _
            idade & ", " & _
            "'" & Replace(endereco, "'", "''") & "', " & _
            especial & ")"

        Conn.Execute sql

        MsgBox "Registro incluido com sucesso", vbInformation
    End If
    
    modoEdicao = False
    idAtual = 0

Frame1.Visible = False
Command4.Visible = False
Command5.Visible = False

List1.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True

CarregarLista

End Sub

Private Sub Command5_Click()

modoEdicao = False
idAtual = 0

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Check1.Value = vbUnchecked

Frame1.Visible = False
Command4.Visible = False
Command5.Visible = False

List1.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_LostFocus()
    If Trim(Text2.Text) <> "" Then
        If Val(Text2.Text) > 125 Then
            MsgBox "A idade não pode ser maior que 125 anos.", vbExclamation
            Text2.Text = ""
            Text2.SetFocus
        End If
    End If
End Sub
