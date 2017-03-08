Option Explicit

Private Const RED   = TRUE
Private Const BLACK = FALSE

Class RBTree
    Public Root

    Private Sub Class_Initialize
        Set Root = Nothing
    End Sub

    Public Function TreeAssert
        TreeAssert = Root Is Nothing Or TreeA(Root) > 0
    End Function

    Private Function TreeA(n)
        Dim lh, rh
        If n Is Nothing Then
            TreeA = 1
        ElseIf n Is Root And isRed(n) Then
            Wscript.Echo "Root is red "
            TreeA = 0
        ElseIf isLess(n, n.Lchild) Or isLess(n.Rchild, n) Then
            Wscript.Echo "Data doesn't match"
            TreeA = 0
        ElseIf isRed(n) And (isRed(n.Lchild) Or isRed(n.Rchild)) Then
            Wscript.Echo "Two red nodes in a row"
            TreeA = 0
        Else
            lh = TreeA(n.Lchild)
            rh = TreeA(n.Rchild)
            If lh <> 0 And rh <> 0 And lh <> rh Then
                Wscript.Echo "Black height violation"
                TreeA = 0
            ElseIf lh <> 0 and rh <> 0 Then
                If isRed(n) Then
                    TreeA = lh
                Else
                    TreeA = lh+1
                End If
            Else
                TreeA = 0
            End If
        End If
    End Function

    Private Function isRed(n)
        If n Is Nothing Then
            isRed = FALSE
        Else
            isRed = (n.Color = RED)
        End If
    End Function

    Private Function isLess(a, b)
        If a Is Nothing Or b Is Nothing Then
            isLess = FALSE
        Else
            isLess = a.Data < b.Data
        End If
    End Function

    Private Sub ColorFlip(n)
        Wscript.Echo "Color flip ", n.Data
        n.Color = Not n.Color
        n.Lchild.Color = Not n.Lchild.Color
        n.Rchild.Color = Not n.Rchild.Color
    End Sub

    Private Function RotateLeft(n)
        Dim x
        Wscript.Echo "Rotate Left ", n.Data
        Set x = n.Rchild
        Set n.Rchild = x.Lchild
        Set x.Lchild = n
        If Not n.Rchild Is Nothing Then
            Set n.Rchild.Parent = n
        End If
        Set x.Parent = n.Parent
        Set n.Parent = x
        x.Color = n.Color
        n.Color = RED
        If x.Parent Is Nothing Then
            Set Root = x
        ElseIf x.Parent.Lchild Is n Then
            Set x.Parent.Lchild = x
        Else
            Set x.Parent.Rchild = x
        End If
        Set RotateLeft = x
    End Function

    Private Function RotateRight(n)
        Dim x
        Wscript.Echo "Rotate Right ", n.Data
        Set x = n.Lchild
        Set n.Lchild = x.Rchild
        Set x.Rchild = n
        If Not n.Lchild Is Nothing Then
            Set n.Lchild.Parent = n
        End If
        Set x.Parent = n.Parent
        Set n.Parent = x
        x.Color = n.Color
        n.Color = RED
        If x.Parent Is Nothing Then
            Set Root = x
        ElseIf x.Parent.Lchild Is n Then
            Set x.Parent.Lchild = x
        Else
            Set x.Parent.Rchild = x
        End If
        Set RotateRight = x
    End Function

    Public Function NodeInsert(byval v)
        Dim n, p, f
        Set n = Root : Set p = Nothing
        Do Until n Is Nothing
            If v = n.Data Then ' Item already in the tree
                Set NodeInsert = n
                Exit Function
            ElseIf v < n.Data Then
                Set p = n
                Set n = n.Lchild
            Else
                Set p = n
                Set n = n.Rchild
            End If
        Loop
        Set f = (New Node).Init(v)
        Set f.Parent = p
        If p Is Nothing Then ' At the root
            Set Root = f
        Else
            If f.Data < p.Data Then
                Set p.Lchild = f
            Else
                Set p.Rchild = f
            End If
            InsertFixup f
        End If
        Root.Color = BLACK
        Set NodeInsert = f
    End Function

    Private Sub InsertFixup(n)
        Dim p, gp
        Set p = n.Parent
        Do While isRed(p)
            Set gp = p.Parent
            If gp.Lchild Is p Then             ' Left side
                If isRed(gp.Rchild) Then       ' Case 1 - Uncle is red
                    ColorFlip gp
                    Set p = gp.Parent
                Else
                    If isRed(p.Rchild) Then    ' Case 2 - Left/Right is red
                        Set p = RotateLeft(p)
                    End If
                    Set gp = RotateRight(gp)   ' Case 3 - Left/Left is red
                    Exit Do
                End If
            Else                               ' Right side
                If isRed(gp.Lchild) Then       ' Case 1 - Uncle is red
                    ColorFlip gp
                    Set p = gp.Parent
                Else
                    If isRed(p.Lchild) Then    ' Case 2 - Right/Left is red
                        Set p = RotateRight(p)
                    End If
                    Set gp = RotateLeft(gp)    ' Case 3 - Right/Right is red
                    Exit Do
                End If
            End If
        Loop
    End Sub

    Private Function Search(byval v)
        Dim q : Set q = Root
        Do Until q Is Nothing
            If v = q.Data Then
                Exit Do
            ElseIf v < q.Data Then
                Set q = q.Lchild
            Else
                Set q = q.Rchild
            End If
        Loop
        Set Search = q
    End Function

    Private Sub DeleteFixup(n)
        ' Invariant: n is not root, n is (d)black
        Dim db, p, s : Set db = n
        Do
            Wscript.Echo "Delete fixup"
            Set p = db.Parent
            If p.Lchild Is db Then ' db is on the left
                Set s = p.Rchild
                If isRed(s) Then                                              ' Case 1 - Red Sibling Case Reduction
                    Wscript.Echo "Case 1 - Red sibling case reduction"
                    Set p = RotateLeft(p)
                    Set p = db.Parent
                    Set s = p.Rchild
                End If
                If Not isRed(s.Lchild) And Not isRed(s.Rchild) Then           ' Case 2 - Black Sibling and Black Children
                    Wscript.Echo "Case 2 - Black sibling and Black children, move up"
                    s.Color = RED
                    If isRed(p) Or p Is Root Then
                        p.Color = BLACK
                        Exit Do
                    End If
                    Set db = p
                Else
                    If isRed(s.Lchild) Then                                   ' Case 3 - Black sibling and left Red child
                        Wscript.Echo "Case 3 - Black sibling and left Red child"
                        Set p.Rchild = RotateRight(p.Rchild)
                    End If
                    Wscript.Echo "Case 4 - Black sibling and right Red child" ' Case 4 - Black sibling and right Red child
                    Set p = RotateLeft(p)
                    p.Lchild.Color = BLACK
                    p.Rchild.Color = BLACK
                    Exit Do
                End If
            Else 'db is on the right
                Set s = p.Lchild
                If isRed(s) Then                                              ' Case 1 - Red Sibling Case Reduction
                    Wscript.Echo "Case 1 - Red sibling case reduction"
                    Set p = RotateRight(p)
                    Set p = db.Parent
                    Set s = p.Lchild
                End If
                If Not isRed(s.Lchild) And Not isRed(s.Rchild) Then           ' Case 2 - Black Sibling and Black Children
                    Wscript.Echo "Case 2 - Black sibling and Black children, move up"
                    s.Color = RED
                    If isRed(p) Or p Is Root Then
                        p.Color = BLACK
                        Exit Do
                    End If
                    Set db = p
                Else
                    If isRed(s.Rchild) Then                                   ' Case 3 - Black sibling and Right Red child
                        Wscript.Echo "Case 3 - Black sibling and right Red child"
                        Set p.Lchild = RotateLeft(p.Lchild)
                    End If
                    Wscript.Echo "Case 4 - Black sibling and left Red child"  ' Case 4 - Black sibling and Left Red child
                    Set p = RotateRight(p)
                    p.Lchild.Color = BLACK
                    p.Rchild.Color = BLACK
                    Exit Do
                End If
            End If
        Loop Until FALSE
    End Sub

    Private Sub SpliceNode(n, q)
        If Not n Is Root And n.Color = BLACK And Not isRed(q) Then ' Leaf black node or n and q are black
            DeleteFixup n
        End If
        If n.Parent Is Nothing Then
            Set Root = q
        ElseIf n.Parent.Lchild Is n Then
            Set n.Parent.Lchild = q
        Else
            Set n.Parent.Rchild = q
        End If
        If Not q Is Nothing Then
            q.Color = BLACK
            Set q.Parent = n.Parent
        End If
        Set n = Nothing
    End Sub

    Public Sub DeleteNode(byval v)
        Dim n, t : Set n = Root
        Do Until n Is Nothing
            If v < n.Data Then
                Set n = n.Lchild
            ElseIf v > n.Data Then
                Set n = n.Rchild
            Else
                If n.Lchild Is Nothing Then
                    SpliceNode n, n.Rchild
                ElseIf n.Rchild Is Nothing Then
                    SpliceNode n, n.Lchild
                Else ' Two children find inorder successor
                    Set t = n.Rchild
                    Do Until t.Lchild Is Nothing
                        Set t = t.Lchild
                    Loop
                    n.Data = t.Data
                    v = t.Data
                    Set n = n.Rchild
                End If
            End If
        Loop
        If Not Root Is Nothing Then
            Root.Color = BLACK
        End If
    End Sub

    Public Sub InsertRandomData(byval cnt)
        Dim i, rnum
        Randomize timer
        For i = 1 to cnt
            rnum = cint(rnd*cnt)
            If Search(rnum) Is Nothing Then
                If NodeInsert(rnum) Is Nothing Then
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Public Sub PrintTree
        PrintNode(Me.Root)
        Wscript.Echo ""
    End Sub

    Private Sub PrintNode(n)
        If Not n Is Nothing Then
            If n.Lchild Is Nothing Then
                wscript.stdout.write "*,b "
            Else
                wscript.stdout.write n.Lchild.Data & "," & n.Lchild.ColorChar(n.Lchild) & " "
            End If
            wscript.stdout.write n.Data & "," & n.ColorChar(n)
            If n.Parent Is Nothing Then
                wscript.stdout.write ",* "
            Else
                wscript.stdout.write "," & n.Parent.Data & " "
            End If
            If n.Rchild Is Nothing Then
                wscript.stdout.write "*,b "
            Else
                wscript.stdout.write n.Rchild.Data & "," & n.Rchild.ColorChar(n.Rchild) & " "
            End If
            wscript.echo ""
            PrintNode(n.Lchild)
            PrintNode(n.Rchild)
        End If
    End Sub

    Private Sub Class_Terminate
        ' wscript.echo "Tree Term"
        Set Root = Nothing
    End Sub
End Class

Class Node
    Public Data
    Public Color
    Public Lchild
    Public Rchild
    Public Parent

    Private Sub Class_Initialize
        ' Wscript.Echo "Node Init"
        Data  = -1
        Color = RED
        Set Lchild = Nothing
        Set Rchild = Nothing
        Set Parent = Nothing
    End Sub

    Public Function Init(n)
        Data  = n
        Color = RED
        Set Lchild = Nothing
        Set Rchild = Nothing
        Set Parent = Nothing
        Set Init = Me
    End Function

    Public Function ColorChar(n)
        If n Is Nothing Then
            ColorChar = "b"
        ElseIf n.Color = RED Then
            ColorChar = "r"
        Else
            ColorChar = "b"
        End If
    End Function

    Public Function isRed
        isRed = (Me.Color = RED)
    End Function

    Private Sub Class_terminate
        ' Wscript.Echo "Node Term"
        Set Lchild = Nothing
        Set Rchild = Nothing
        Set Parent = Nothing
    End Sub
End Class

Dim T, n, i, S
Set T = New RBTree
T.InsertRandomData 20
T.PrintTree
If Not T.TreeAssert Then
    Wscript.Echo "Tree is not valid."
End If
Do Until T.Root Is Nothing
    T.DeleteNode T.Root.Data
    T.PrintTree
    If Not T.TreeAssert Then
        Wscript.Echo "Tree is not valid."
    End If
Loop
