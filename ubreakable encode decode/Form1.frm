VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Unbreakable Encode/ Decode v 1.0 alpha"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Encode"
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Decode"
         Height          =   495
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Letter"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Passletter"
         Height          =   615
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   975
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   90
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Result"
         Height          =   615
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   855
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   90
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encode"
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "Encode"
         Height          =   495
         Left            =   2040
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Result"
         Height          =   615
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   855
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   90
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Passletter"
         Height          =   615
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   975
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   90
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Letter"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is probably the crappest encryption going here
'i wrote this cause i saw it on descovery channel
'its an encryption that has been used for ages even used
'in wars
'its supposed to use passwords and encode whole
'sentences but i didn't feel like doing that
'i will soon, mabey this weekend
'anyway what this does it gets the ascii code of the
'uppercase letter in text1 and adds the ascii code of
'the uppercase letter in text2 it then subtracts 64 from
'it and converts this to a new character to test draw a
'table similar to this one:
'A|B|C|D|E|F|G|H|I|J ETC. UNTIL Z.
'B|C|D|E|F|G|H|I|J|K
'C|D|E|F|G|H|I|J|K|L
'D
'E
'ETC. UNTIL Z
'YOU GET THE PICTURE
'Now take the word you want to encode say norm and choose
'password say dog . what you do now is go to the 1st
'letter of the password on the top row and the 1st letter
'of the text in the vertical row and move down/across until
'they meet. this is the encoded character. you keep doing
'this until the text is encoded. if the password is at the
'end start at the beginning again.
'to decode do the same as the encryption but now you do
'the password in the vertical side and the text on the
'horizontal side
'tada
'Themba Kriger : wayne_kerr@hushmail.com send me any
'modifications plz! thanx
Private Sub Command1_Click() 'decoding
On Error Resume Next
Text4 = Chr(Asc(UCase(Text6)) - (Asc(UCase(Text5)) - 64))
End Sub

Private Sub Command2_Click() 'encoding
On Error Resume Next
Text3 = Chr(Asc(UCase(Text1)) + (Asc(UCase(Text2)) - 64))
End Sub
