VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Treeview To HTML Example"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlTreeview 
      Left            =   3840
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   495
      Left            =   3540
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Populate"
      Height          =   495
      Left            =   3540
      TabIndex        =   1
      Top             =   420
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvTest 
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9022
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Click on a load to load a new child."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------
' Name:           frmMain (frmMain.frm)
' Type:           Form
' Description:    -
'
' Author:         Koen
' Date:           14-1-2002
' E-mail:         deye_99@yahoo.com
' Copyright:      This work is copyrighted © 2001
'
' Comments:       This code has been modified in many ways,
'                 and most of the code is new
'------------------------------------------------------------------
Option Explicit

'   Copyright (c) 2001, Chetan Sarva. All rights reserved.
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions are
'   met:
'
'   -Redistributions of source code must retain the above copyright notice,
'    this list of conditions and the following disclaimer.
'
'   -Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in the
'    documentation and/or other materials provided with the distribution.
'
'   -Neither the name of pixelcop.com nor the names of its contributors may
'    be used to endorse or promote products derived from this software
'    without specific prior written permission.
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
'   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
'   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
'   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
'   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
'   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
'   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Private Sub cmdExport_Click()

    Dim sFile As String ' Filename
    
    'sFile = InputBox("Enter a filename (will be saved to app.path)", , "tree.html")
    If sFile = "" Then sFile = "tree.html"
    
    Dim i As Integer ' Free file
    i = FreeFile()
    
    Open App.Path & "\" & sFile For Output As #i
        Print #i, "" & Tree2HTML(tvTest, "Img/", "Taxons/") & ""
    Close #i
    
    MsgBox "File Saved to '" & App.Path & "\tree.html'.", vbInformation
    
End Sub

Private Sub cmdFill_Click()

    ' Fill our treeview wtih some nodes
    
    Dim i As Long
    Dim ti As Node
    
    Set ti = tvTest.Nodes.Add(, , , "Node 1", getRandomImage())
        ti.Expanded = True
        ti.Bold = True
    tvTest.Nodes.Add ti, tvwChild, , "Node 2", getRandomImage()
    tvTest.Nodes.Add ti, tvwChild, , "Node 3", getRandomImage()
    tvTest.Nodes.Add ti, tvwChild, , "Node 4", getRandomImage()
    Set ti = tvTest.Nodes.Add(ti, tvwChild, , "Node 5", getRandomImage())
        ti.Expanded = True
        ti.Bold = True
    tvTest.Nodes.Add ti, tvwChild, , "Node 6", getRandomImage()
    tvTest.Nodes.Add ti, tvwChild, , "Node 7", getRandomImage()
    
    
End Sub

Private Sub Form_Load()
  'Load example images from a directory
  Call LoadFolderInIml(Me.imlTreeview, App.Path & "\Taxons\")
  tvTest.ImageList = Me.imlTreeview
End Sub

Private Sub tvTest_NodeClick(ByVal Node As MSComctlLib.Node)
  'Add a new child
  tvTest.Nodes.Add Node.Index, tvwChild, , "Node in " & Node.Index, getRandomImage()
End Sub

Private Function getRandomImage() As String
Dim i As Integer
  i = Int(Rnd * (imlTreeview.ListImages.Count - 1))
  getRandomImage = imlTreeview.ListImages(i + 1).Key
End Function
