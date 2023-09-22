Public Class Consensus1
    Private _sp() As splits
    Private _nwktree As String
    Sub New()

        _sp = Nothing


    End Sub
   

    Public Property sp() As splits()
        Get
            Return _sp
        End Get
        Set(ByVal val As splits())
            _sp = val
        End Set
    End Property
    Public Function consensussplits(ByVal sp() As splits,mre As Boolean) As splits

        Dim newsp As splits
        Dim newsp1 As splits
        Dim nsplits As Integer = sp(0).nOTUs.GetLength(0) - 1

        ReDim newsp.nOTUs(nsplits + 1, sp(0).nOTUs.GetLength(1))
        ReDim newsp.notus1(nsplits + 1, sp(0).notus1.GetLength(1))
        ReDim newsp1.nOTUs((nsplits + 1) * sp.Length, sp(0).nOTUs.GetLength(1))
        ReDim newsp1.notus1((nsplits + 1) * sp.Length, sp(0).notus1.GetLength(1))

        Dim g As Integer = 1
        Dim g1 As Integer = 1
        For h = 0 To sp.Length - 1
            For i = 1 To nsplits



                If contain(sp(h).nOTUs(i, 2), newsp.nOTUs) = 0 Then

                    Dim c As Double = countsplits(sp(h).nOTUs(i, 2), sp)
                    If c > 50 Then
                        newsp.nOTUs(g, 1) = sp(h).nOTUs(i, 1)
                        newsp.nOTUs(g, 2) = sp(h).nOTUs(i, 2)
                        newsp.notus1(g, 1) = c
                        newsp.notus1(g, 0) = branchL(sp(h).nOTUs(i, 2), sp)
                        g = g + 1

                    ElseIf treeviewer2.splitsize(sp(h).nOTUs(i, 1)) = 1 Then
                        newsp.nOTUs(g, 1) = sp(h).nOTUs(i, 1)
                        newsp.nOTUs(g, 2) = sp(h).nOTUs(i, 2)

                        If mre = True Then
                            newsp.notus1(g, 0) = c
                            newsp.notus1(g, 0) = branchL(sp(h).nOTUs(i, 2), sp)
                        End If


                        g = g + 1

                    Else
                        If contain(sp(h).nOTUs(i, 2), newsp1.nOTUs) = 0 Then
                            newsp1.nOTUs(g1, 1) = sp(h).nOTUs(i, 1)
                            newsp1.nOTUs(g1, 2) = sp(h).nOTUs(i, 2)
                            newsp1.notus1(g1, 1) = c
                            newsp1.notus1(g1, 0) = branchL(sp(h).nOTUs(i, 2), sp)
                            g1 = g1 + 1
                        End If
                    End If
                End If
            Next
        Next

        If mre = True Then
            Do While g < newsp.nOTUs.GetLength(0) - 1
                Dim max As Double = 0
                Dim imax As Integer = 0
                For i = 0 To newsp1.notus1.GetLength(0) - 1
                    If newsp1.notus1(i, 1) > max Then
                        max = newsp1.notus1(i, 1)
                        imax = i

                    End If
                Next

                If max = 0 Then Exit Do
                Dim newa As New newicker
                Dim chek1 As String
                chek1 = newa.rewritesplits(" " & newsp1.nOTUs(imax, 2), sp(0).otumat1.GetLength(0) - 1)
                'If imax = 2 Then Stop
                Dim comp As Boolean = True

                For j = 1 To g - 1
                    If treeviewer2.splitsize(newsp.nOTUs(j, 1)) > 1 Then

                        Dim chek2 As String = newa.rewritesplits(" " & newsp.nOTUs(j, 2), sp(0).otumat1.GetLength(0) - 1)
                        If Form1.checkcomp(chek1, chek2, sp(1).otumat1.GetLength(0) - 1) = False Then
                            comp = False
                            Exit For
                        End If
                    End If

                Next
                'If imax = 24 Then Stop
                If comp = True Then
                    newsp.nOTUs(g, 1) = newsp1.nOTUs(imax, 1)
                    newsp.nOTUs(g, 2) = newsp1.nOTUs(imax, 2)
                    newsp.notus1(g, 1) = newsp1.notus1(imax, 1)
                    newsp.notus1(g, 0) = newsp1.notus1(imax, 0)
                    g = g + 1

                End If
                newsp1.notus1(imax, 1) = -1
            Loop
            g = g - 1
        End If


        newsp.nOTUs = resizearraystring(newsp.nOTUs, g - 1)

        newsp.notus1 = resizearraysingle(newsp.notus1, g - 1)
        Dim arraysize(newsp.nOTUs.GetLength(0) - 1) As Single
        For i = 1 To g - 1
            Dim outg As Integer = Form1.ToolStripComboBox1.SelectedIndex
            If outg = 0 Then outg = 1
            If newsp.nOTUs(i, 1).Contains(" " & outg & " ") Then
                newsp.nOTUs(i, 1) = treeviewer2.revertsplit(newsp.nOTUs(i, 1), sp(0).otumat1.GetLength(0) - 1)
            End If
            Dim si As Integer = treeviewer2.splitsize(newsp.nOTUs(i, 1))

            arraysize(i) = si


        Next
        Dim count As Integer = 1
        Dim newsp2 As splits
        ReDim newsp2.nOTUs(newsp.nOTUs.GetLength(0) - 1, newsp.nOTUs.GetLength(1) - 1)
        ReDim newsp2.notus1(newsp.notus1.GetLength(0) - 1, newsp.notus1.GetLength(1) - 1)
        newsp2.otumat1 = newsp.otumat1
        newsp2.splitest = sp(0).splitest
        For i = 1 To g - 1

            For j = 1 To g - 1
                If arraysize(j) = i Then
                    newsp2.nOTUs(count, 1) = newsp.nOTUs(j, 1)
                    newsp2.nOTUs(count, 2) = newsp.nOTUs(j, 2)
                    newsp2.notus1(count, 0) = newsp.notus1(j, 0)
                    newsp2.notus1(count, 1) = newsp.notus1(j, 1)
                    count = count + 1
                End If
            Next
        Next

      
        newsp.splitest = sp(0).splitest

        newsp2.otumat1 = sp(0).otumat1
        newsp2.DSTarr = sp(0).DSTarr
        newsp2.nOTUs(newsp2.nOTUs.GetLength(0) - 1, 1) = treeviewer2.revertsplit(newsp2.nOTUs(newsp2.nOTUs.GetLength(0) - 2, 1), newsp2.otumat1.GetLength(0) - 1)
        newsp2.nOTUs(newsp.nOTUs.GetLength(0) - 1, 2) = newsp2.nOTUs(newsp2.nOTUs.GetLength(0) - 2, 2)
        newsp2.notus1(newsp.nOTUs.GetLength(0) - 1, 2) = newsp2.notus1(newsp2.nOTUs.GetLength(0) - 2, 2)


        
        

      
        Return newsp2
    End Function
    Function contain(ByVal a As String, ByVal b(,) As String) As Integer
        Dim result As Integer
        Try
            For i = 1 To b.GetLength(0) - 1
                If a = b(i, 2) Then
                    result = i
                    i = 100000
                End If
            Next
        Catch
        End Try
        Return result
    End Function
    Private Function countsplits(ByVal split As String, ByVal sp1() As splits) As Single
        Dim count As Single

        For i = 0 To sp1.Length - 1
            Dim a As Integer = contain(split, sp1(i).nOTUs)
            If a <> 0 Then
                If sp1(i).notus1(a, 0) > 1 / 1000000000000 Then
                    count = count + 1
                End If
            End If
        Next
        count = CInt(Int(count / sp1.Length * 100))


        Return count
    End Function
    Function branchL(ByVal split As String, ByVal sp1() As splits) As Single
        Dim sum As Single
        Dim count As Integer
        Dim index As Integer
        For i = 0 To sp1.Length - 1

            For j = 1 To sp1(i).nOTUs.GetLength(0) - 1
                If split = sp1(i).nOTUs(j, 2) Then
                    index = j

                    j = 100000
                Else
                    index = 0
                End If
            Next
            If index <> 0 Then
                count = count + 1
                sum = sum + sp1(i).notus1(index, 0)
            End If
        Next



        Return (sum / count)
    End Function
    Public Function resizearraystring(ByVal a As String(,), ByVal b As Integer) As String(,)
        Dim newarray(b + 1, a.GetLength(1)) As String
        For i = 1 To b
            For j = 0 To a.GetLength(1) - 1
                newarray(i, j) = a(i, j)
            Next
        Next
        Return newarray
    End Function
    Public Function resizearraysingle(ByVal a As String(,), ByVal b As Integer) As String(,)
        Dim newarray(b + 1, a.GetLength(1)) As String
        For i = 1 To b
            For j = 0 To a.GetLength(1) - 1
                newarray(i, j) = a(i, j)
            Next
        Next
        Return newarray
    End Function
End Class
