Class Stack
    Private mSize
    Private mPos
    Private mArr()

    Private Sub Class_Initialize
        mSize = 0
        mPos = -1
    End Sub

    Private Sub Class_Terminate
    End Sub

    Public Function Init(byval n)
        mSize = n
        mPos = -1
        Redim mArr(mSize)
        Set Init = Me
    End Function

    Public Sub Push(byval v)
        if mPos+1 >= mSize Then
            Wscript.Echo "Stack is full."
        Else
            mPos = mPos+1
            mArr(mPos) = v
        End If
    End Sub

    Public Function Pop
        Dim v
        If mPos > -1 Then
            v = mArr(mPos)
            mPos = mPos-1
        Else
            v = -1
            Wscript.echo "Stack is empty."
        End If
        Pop = v
    End Function

    Public Function isEmpty
        isEmpty = (mPos < 0)
    End Function

End Class

Sub Hanoii(n, src, dest, aux)
    Dim i, j, ssrc, sdest, psrc, pdest
    If n mod 2 = 0 Then
        j = dest : dest = aux : aux = j ' Swap aux and dest
    End If
    j = 2^n-1
    sdest ="a" : pdest = "c"
    For i = 1 to 2^n-1
        If i mod 2 = 1 Then ' smallest
            Wscript.echo "Move from ", sdest, " to ", pdest
            ssrc = pdest : sdest = pdest
        Else
            Wscript.echo "Move from ", ssrc, " to ", psrc
            pdest = ssrc : psrc =
    Next
End Sub

Dim s : Set s = (new Stack).Init(3)
Dim j
s.Push(1)
s.Push(2)
s.Push(3)
j = s.Pop
wscript.echo j
j = s.Pop
wscript.echo j
Hanoii 3, "a", "b", "c"
