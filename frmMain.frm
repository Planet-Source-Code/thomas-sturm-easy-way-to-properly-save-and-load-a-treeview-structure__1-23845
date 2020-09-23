VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Save and Load TreeView"
   ClientHeight    =   3936
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5556
   LinkTopic       =   "Form1"
   ScaleHeight     =   3936
   ScaleWidth      =   5556
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   372
      Left            =   4080
      TabIndex        =   3
      Top             =   3480
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   372
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   1212
   End
   Begin VB.CommandButton cmdAddNode 
      Caption         =   "Add Node"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1572
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3132
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5052
      _ExtentX        =   8911
      _ExtentY        =   5525
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nodX As Node

Private Sub cmdAddNode_Click()
Set nodX = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "Node" & TreeView1.Nodes.Count + 1, "Node" & TreeView1.Nodes.Count + 1)
End Sub

Private Sub cmdLoad_Click()
LoadTVFromFile
End Sub

Private Sub cmdSave_Click()
SaveTVToFile
End Sub

Private Sub Form_Load()

Set nodX = TreeView1.Nodes.Add(, , "Top", "Root")

Set nodX = TreeView1.Nodes.Add("Top", tvwChild, "Node1", "Node1")
Set nodX = TreeView1.Nodes.Add("Top", tvwChild, "Node2", "Node2")

For i% = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i%).Expanded = True
Next i%

End Sub
