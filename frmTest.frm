VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Control Resize"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblInstructions 
      AutoSize        =   -1  'True
      Height          =   1635
      Left            =   0
      TabIndex        =   1
      Tag             =   "0030"
      Top             =   0
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Made by Edward Catchpole for Planet Source Code; edward_eddie_theman@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Tag             =   "0330"
      Top             =   2760
      Width           =   11370
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmResize   As New ControlResizer

Private Sub Form_Load()

    lblInstructions.Caption = "Simply change the tag of controls to be resized to the following numbers to make them resize with the form:" _
            & vbCrLf & "0=Make no changes" _
            & vbCrLf & "1=Change in proportion to the form" _
            & vbCrLf & "2=Change statically (add change on)" _
            & vbCrLf & "3=Change statically but limited (same as 2 but value does not go below 0)" _
            & vbCrLf & vbCrLf & "The four numbers go in this order: Left mode, Top mode, Width mode, Height mode"
    
    Call frmResize.InitResizer(Me) 'initiate resize class
    
End Sub

Private Sub Form_Resize()

    Call frmResize.FormResized(Me) 'force control resize
    
End Sub
