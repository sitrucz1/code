Option Explicit
On Error Resume Next

Sub Hanoi(Disk, Pegfrom, Pegto, Pegaux)
    If Disk > 0 Then
        Call Hanoi(Disk-1, Pegfrom, Pegaux, Pegto)
        Wscript.Echo "Move Disk " & Disk & " From Peg " & Pegfrom & _
                                           " To Peg " & Pegto
        Call Hanoi(Disk-1, Pegaux, Pegto, Pegfrom)
    End If
End Sub

Function getDisks
    If Wscript.Arguments.Count > 0 Then
        getDisks = Wscript.Arguments.Item(0)
    Else
        getDisks = 4
    End If
End Function

Call Hanoi(getDisks, "A", "C", "B")
