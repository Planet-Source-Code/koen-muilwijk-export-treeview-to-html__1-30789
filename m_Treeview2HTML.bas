Attribute VB_Name = "modTreeview2HTML"
'------------------------------------------------------------------
' Name:           modTreeview2HTML (m_Treeview2HTML.bas)
' Type:           Module
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
'
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
Private m_sImagePath As String
Private nMaxIndent As Integer
Private a_bIsLastNode() As Boolean
Private m_sNodeImagePath As String
Private m_tv As TreeView

Private Function GetIndent(Indent As Integer, Optional HasChildren As Boolean = False, Optional bIsLastNode As Boolean = False, Optional bIsLastChildNode As Boolean) As String
Dim i As Long
Dim sImgAttributes As String
    
    ' This sub is a help for Node2HTML.
    ' It adds indentation to the node
    ' depending on its type and depth
    ' in the tree.

    If Indent = 0 Then Exit Function
    
    If Indent > nMaxIndent Then
      'Keep an array to remember for each level if the node is the last of it's children.
      ' This information is used later to choose between a simple line '|' or an empty space
      ReDim Preserve a_bIsLastNode(Indent)
      nMaxIndent = Indent
    End If
    
    a_bIsLastNode(Indent) = bIsLastNode
    sImgAttributes = "align=""ABSMIDDLE"" width=""20"" height=""16"" border=""0"""
    
    For i = 1 To Indent
        
      If i < Indent Then
        If a_bIsLastNode(i) Then
          'It seems that there is no node anymore on this level. Display some empty space
          GetIndent = GetIndent & "<img src=""" & m_sImagePath & "tvw_Nothing.gif"" " & sImgAttributes & ">" '"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Else
          GetIndent = GetIndent & "<img src=""" & m_sImagePath & "tvw_Nochildren.gif"" " & sImgAttributes & ">" '"|&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
      Else
        If HasChildren Then
          If bIsLastNode Then
            'This is a special version of Plus (the line 'stops' after the node)
            GetIndent = GetIndent & "<img src=""" & m_sImagePath & "tvw_PlusLast.gif"" " & sImgAttributes & ">" '"+----"
          Else
            GetIndent = GetIndent & "<img src=""" & m_sImagePath & "tvw_Plus.gif"" " & sImgAttributes & ">" '"+----"
          End If
        Else 'Has no children
          If bIsLastChildNode Then
            GetIndent = GetIndent & "<img src=""" & m_sImagePath & "tvw_LastSibling.gif"" " & sImgAttributes & ">" '"L----"
          Else
            GetIndent = GetIndent & "<img src=""" & m_sImagePath & "tvw_Node.gif"" " & sImgAttributes & ">" '"|----"
          End If
        End If
      End If
    Next i

End Function


Private Function Node2HTML(nodeParent As Node, Indent As Integer) As String
    
    'nodeParent.Selected = True
    
    Dim i As Long
    Dim strHTML As String
    Dim nodeChild As Node
    Dim bIsLastChild As Boolean
    
    ' Holds our HTML strings
    Dim sRowStart As String
        'sRowStart = "<TR><TD>"
    Dim sRowEnd As String
       ' sRowEnd = "</TD></TR>"
        
    ' Add this node
    strHTML = getNodeText(nodeParent)
    
    'If nodeParent.Children > 0 Then _
    '    strHTML = "" & strHTML & ""
    
    If Not nodeParent.Parent Is Nothing Then
      'I've taken this check from:
      ' Made by: Manu Bangia
      ' bangiamanu@yahoo.com
      ' http://www.manubangia.com
      bIsLastChild = (nodeParent.Index = nodeParent.Parent.Child.LastSibling.Index)
    End If
    
    Node2HTML = Node2HTML & sRowStart & GetIndent(Indent, True, nodeParent.Next Is Nothing, bIsLastChild) & strHTML & sRowEnd & "<BR>" & vbCrLf
    
    ' Look at the first child node
    Set nodeChild = nodeParent.Child
    
    Do While Not (nodeChild Is Nothing)
        
        If nodeChild.Children > 0 Then
            
            ' Recursion for any nodes with children
            Node2HTML = Node2HTML & Node2HTML(nodeChild, Indent + 1)
            
        Else
            
            ' No children, add it and move on
            strHTML = getNodeText(nodeChild)
            
            If Not nodeChild.Parent Is Nothing Then
              bIsLastChild = (nodeChild.Index = nodeChild.Parent.Child.LastSibling.Index)
            End If
            
            Node2HTML = Node2HTML & sRowStart & GetIndent(Indent + 1, False, False, bIsLastChild) & strHTML & sRowEnd & "<BR>" & vbCrLf
                
        End If
   
        'Get the current child node's next sibling
        Set nodeChild = nodeChild.Next
       
     Loop
    
End Function

Private Function getNodeText(Node As MSComctlLib.Node)
Dim sReturn As String
  
  sReturn = "<span class=""node"">"
  
  If m_sNodeImagePath <> "" And Node.Image <> "" Then
    sReturn = sReturn & "<img src=""" & m_sNodeImagePath & "" & Node.Image & ".gif""/>&nbsp;"
  End If
  
  If Node.Bold = True Then
    sReturn = sReturn & "<B>" & Node.Text & "</B>"
  Else
    sReturn = sReturn & Node.Text
  End If
  sReturn = sReturn & "</SPAN>"
  
  getNodeText = sReturn
End Function

'------------------------------------------------------------------
' Procedure : Tree2HTML
' Date      : 10-1-2002
' Author    : Koen Muilwijk
' Purpose   : Creates a string containing the HTML wich represents the given treeview
' Parameters: tv              = the treeview to export
'             sImagePath      = The relative path where the treeview images can been found
'             sNodeImagePath  = The relative path where the images for the nodes are located.
'                               Leave empty if you don't want the nodes to have images.
'                               The images must have .gif extension and the same name as the image key.
' Output    : String
'------------------------------------------------------------------
Public Function Tree2HTML(tv As TreeView, sImagePath As String, Optional sNodeImagePath As String) As String
    
   ' This is the sub that gets it going.
   ' For each root node in the tree,
   ' we call Node2HTML to dig through
   ' its children and convert it to HTML.
    
   Dim nodeRoot As Node
   
   m_sImagePath = sImagePath
   m_sNodeImagePath = sNodeImagePath
   Set m_tv = tv
   
   Tree2HTML = "<html><head><style>.node { font : 1px MS Sans Serif;position : relative;top : -4px;  left : 2px;}</style><link rel=""STYLESHEET"" type=""text/css"" href=""Treeview.css""><title>Treeview</title></head><body>"
   
   For Each nodeRoot In tv.Nodes
      If (nodeRoot.Parent Is Nothing) Then
         Tree2HTML = Tree2HTML & Node2HTML(nodeRoot, 0)
      End If
   Next
   
   Tree2HTML = Tree2HTML & "</body></html>"
    
End Function


