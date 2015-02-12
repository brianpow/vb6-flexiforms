VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- clsLayout ---

Option Explicit

Public Enum enLayoutStyle
   lsNone = 0     ' fixed position and size
   lsLeft = 1     ' fixed indentation from left of form
   lsRight = 2    ' fixed indentation from right of form
   lsHorizontal = 3 ' fixed left and right indentation (resize horizontally)
   lsTop = 4      ' fixed indentation from top of form
   lsBottom = 8   ' fixed indentation from bottom of form
   lsVertical = 12 'fixed top and bottom indentation (resize vertically)
   lsFull = 15    ' resize horizontally and vertically
   lsXProp = 16   ' position is dependent on form width
   lsYProp = 32   ' position is dependent on form height
   lsWprop = 64   ' width is proportional to form width
   lsDprop = 128  ' width is proportional to form depth (height)
End Enum

Private m_Controls As Collection ' of clsControl objects
Private Function hasScaleMode(Ctrl As Object)
    hasScaleMode = PropertyExist(Ctrl, "ScaleMode", EPType.ReadableProperties)
End Function
Private Function FindContainerWithScaleMode(Ctrl As Object) As Object
    If Not hasScaleMode(Ctrl) Then
        Set FindContainerWithScaleMode = FindContainerWithScaleMode(Ctrl.Container)
    Else
        Set FindContainerWithScaleMode = Ctrl
    End If
End Function
Private Function getLeft(Ctrl As Object)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        getLeft = Ctrl.Left
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        getLeft = Container.ScaleX(Ctrl.Left, Container.ScaleMode, vbTwips)
    End If
End Function

Private Function getTop(Ctrl As Object)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        getTop = Ctrl.Top
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        getTop = Container.ScaleY(Ctrl.Top, Container.ScaleMode, vbTwips)
    End If
End Function

Private Function getWidth(Ctrl As Object)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        getWidth = Ctrl.Width
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        getWidth = Container.ScaleX(Ctrl.Width, Container.ScaleMode, vbTwips)
    End If
End Function
Private Function getHeight(Ctrl As Object)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        getHeight = Ctrl.Height
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        getHeight = Container.ScaleY(Ctrl.Height, Container.ScaleMode, vbTwips)
    End If
End Function
Private Function setTop(Ctrl As Object, Top&)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        Ctrl.Top = Top
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        Ctrl.Top = Container.ScaleY(Top, vbTwips, Container.ScaleMode)
    End If
End Function

Private Function setLeft(Ctrl As Object, Left&)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        Ctrl.Left = Left
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        Ctrl.Left = Container.ScaleX(Left, vbTwips, Container.ScaleMode)
    End If
End Function
Private Function setWidth(Ctrl As Object, Width&)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        Ctrl.Width = Width
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        Ctrl.Width = Container.ScaleX(Width, vbTwips, Container.ScaleMode)
    End If
End Function
Private Function setHeight(Ctrl As Object, Height&)
Dim Container As Object
    If TypeOf Ctrl Is Form Then
        Ctrl.Height = Height
    Else
        Set Container = FindContainerWithScaleMode(Ctrl.Container)
        Ctrl.Height = Container.ScaleY(Height, vbTwips, Container.ScaleMode)
    End If
End Function

Private Function AddControl(Ctrl As Control, Style As enLayoutStyle)
On Local Error GoTo errortrap
   Dim udt As New clsControl
   With udt
        .Style = Style
        Set .Ctrl = Ctrl
        .LeftIndent = getLeft(Ctrl)
        .TopIndent = getTop(Ctrl)
        .RightIndent = getWidth(Ctrl.Container) - (getWidth(Ctrl) + getLeft(Ctrl))
        .BottomIndent = getHeight(Ctrl.Container) - (getHeight(Ctrl) + getTop(Ctrl))
        .Width = getWidth(Ctrl)
        .Height = getHeight(Ctrl)
        .ContainerMinHeight = getHeight(Ctrl.Container)
        .ContainerMinWidth = getWidth(Ctrl.Container)
   End With
   m_Controls.Add udt
   Exit Function
errortrap:
    Debug.Print "AddControl(): Adding failed! " & TypeName(Ctrl), Ctrl.Name
   ' there are some controls that have no visual element in run time (ie timer, image list etc)
End Function

Private Function NotNeg(x As Long)
   NotNeg = IIf(x < 0, 0, x)
End Function

Private Function GetLayoutStyle(s As String) As enLayoutStyle
   Dim ls As enLayoutStyle
   ls = lsNone
   Dim i As Integer
   For i = 1 To Len(s)
      Select Case Mid$(s, i, 1)
      Case "L"
         ls = ls Or lsLeft
      Case "R"
         ls = ls Or lsRight
      Case "T"
         ls = ls Or lsTop
      Case "B"
         ls = ls Or lsBottom
      Case "V"
         ls = ls Or lsVertical
      Case "H"
         ls = ls Or lsHorizontal
      Case "F"
         ls = ls Or lsFull
      Case "X"
         ls = ls Or lsXProp
      Case "Y"
         ls = ls Or lsYProp
      Case "W"
         ls = ls Or lsWprop
      Case "D"
         ls = ls Or lsDprop
      Case "N"
         ls = lsNone
      End Select
   Next
   GetLayoutStyle = ls
End Function

Public Function RedrawLayout(Optional ResizeBelowMinSize As Boolean = False)
   Dim Width As Currency
   Dim Height As Currency
   Dim Ctrl As clsControl
      
   For Each Ctrl In m_Controls
      With Ctrl
         Width = getWidth(.Ctrl.Container)
         If Not ResizeBelowMinSize Then If Width < .ContainerMinWidth Then Width = .ContainerMinWidth
         Height = getHeight(.Ctrl.Container)
         If Not ResizeBelowMinSize Then If Height < .ContainerMinHeight Then Height = .ContainerMinHeight
         
         
         If (.Style And lsLeft) = lsLeft Then
            Call setLeft(.Ctrl, .LeftIndent)
         ElseIf (.Style And lsXProp) = lsXProp Then
            Call setLeft(.Ctrl, (.LeftIndent * Width) / .ContainerMinWidth)
         End If
         If (.Style And lsHorizontal) = lsHorizontal Then
            Call setWidth(.Ctrl, NotNeg((Width - .RightIndent) - .LeftIndent))
         Else
            If (.Style And lsWprop) = lsWprop Then
               Call setWidth(.Ctrl, (.Width * Width) / .ContainerMinWidth)
            End If
            If (.Style And lsRight) = lsRight Then
               Call setLeft(.Ctrl, NotNeg((Width - .RightIndent) - getWidth(.Ctrl)))
            End If
         End If
         
         If (.Style And lsTop) = lsTop Then
            Call setTop(.Ctrl, .TopIndent)
         ElseIf (.Style And lsYProp) = lsYProp Then
            Call setTop(.Ctrl, (.TopIndent * Height) / .ContainerMinHeight)
         End If
         If (.Style And lsVertical) = lsVertical Then
            Call setHeight(.Ctrl, NotNeg((Height - .BottomIndent) - .TopIndent))
         Else
            If (.Style And lsDprop) = lsDprop Then
               Call setHeight(.Ctrl, (.Height * Height) / .ContainerMinHeight)
            End If
            If (.Style And lsBottom) = lsBottom Then
               Call setTop(.Ctrl, NotNeg((Height - .BottomIndent) - getHeight(.Ctrl)))
            End If
         End If
      End With
   Next

End Function

Public Function SetLayout(frm As Form)
      
      Set m_Controls = New Collection ' of clsControl objects
      Dim con As Control
      Dim ls As enLayoutStyle
      For Each con In frm.Controls
         ls = GetLayoutStyle(con.Tag)
         If ls <> lsNone Then
            AddControl con, GetLayoutStyle(con.Tag)
         End If
      Next
End Function




