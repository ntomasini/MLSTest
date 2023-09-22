Public Class consensusnetwork
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
    Public Function consensusnet(ByVal sp() As splits, ByVal umbral As Integer, ByVal minsup As Single) As splits
        Dim newsp As splits
        Dim nsplits As Integer = sp(0).nOTUs.GetLength(0) - 1

        ReDim newsp.nOTUs((nsplits) * sp.Length - 1, sp(0).nOTUs.GetLength(1))
        ReDim newsp.notus1(nsplits * sp.Length - 1, sp(0).notus1.GetLength(1))
        Dim g As Integer = 1
        For h = 0 To sp.Length - 1
            For i = 1 To nsplits



                If contain(sp(h).nOTUs(i, 2), newsp.nOTUs) = 0 Then
                    Dim c As Double = countsplits(sp(h).nOTUs(i, 2), sp, minsup)
                    If c > umbral Then
                        newsp.nOTUs(g, 1) = sp(h).nOTUs(i, 1)
                        newsp.nOTUs(g, 2) = sp(h).nOTUs(i, 2)
                        newsp.notus1(g, 1) = c
                        newsp.notus1(g, 0) = branchL(sp(h).nOTUs(i, 2), sp)
                        g = g + 1

                    End If
                End If
            Next
        Next
        newsp.nOTUs = resizearraystring(newsp.nOTUs, g - 1)
        newsp.notus1 = resizearraysingle(newsp.notus1, g - 1)
        newsp.splitest = sp(0).splitest
        newsp.otumat1 = sp(0).otumat1
        newsp.DSTarr = sp(0).DSTarr
        Return newsp
    End Function
    Function contain(ByVal a As String, ByVal b(,) As String) As Integer
        Dim result As Integer
        Try
            For i = 1 To b.GetLength(0) - 1
                If a = b(i, 2) Then
                    result = i
                    Exit For
                End If
            Next
        Catch
        End Try
        Return result
    End Function
    Private Function countsplits(ByVal split As String, ByVal sp1() As splits, ByVal minsup As Single) As Single
        Dim count As Single
        If minsup <= 0 Then
            For i = 0 To sp1.Length - 1
                Dim a As Integer = contain(split, sp1(i).nOTUs)
                If a <> 0 Then
                    If sp1(i).notus1(a, 0) > 1 / 1000000000000 Then
                        count = count + 1
                    End If
                End If
            Next
        Else
            For i = 0 To sp1.Length - 1
                Dim a As Integer = contain(split, sp1(i).nOTUs)
                If a <> 0 Then
                    If sp1(i).notus1(a, 0) > 1 / 1000000000000 AndAlso sp1(i).notus1(a, 1) > minsup Then
                        count = count + 1
                    End If
                End If
            Next
        End If




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
    Function resizearraystring(ByVal a As String(,), ByVal b As Integer) As String(,)
        Dim newarray(b + 1, a.GetLength(1)) As String
        For i = 1 To b
            For j = 0 To a.GetLength(1) - 1
                newarray(i, j) = a(i, j)
            Next
        Next
        Return newarray
    End Function
    Function resizearraysingle(ByVal a As String(,), ByVal b As Integer) As String(,)
        Dim newarray(b + 1, a.GetLength(1)) As String
        For i = 1 To b
            For j = 0 To a.GetLength(1) - 1
                newarray(i, j) = a(i, j)
            Next
        Next
        Return newarray
    End Function
End Class
