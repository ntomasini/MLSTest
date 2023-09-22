Public Class newicker
    Private _notus1(,) As Single
    Private _notus(,) As String


    Sub New()
        _notus = Nothing
        _notus1 = Nothing


    End Sub

    Public Function makenwk(ByVal notus1(,) As String, ByVal notus(,) As String, ByVal ntax As Integer, ByVal otumat As String(,), ByVal nodo() As String) As String
        Dim splits(,) As String

        splits = rewriteallsplits(notus, ntax)

        Dim nwktree As String



        nwktree = writesplits(splits, notus1, otumat, nodo)


        Return nwktree

    End Function


    Function rewriteallsplits(ByVal spltarray As String(,), ByVal ntax As Integer) As String(,)
        Dim nsplits As Integer = spltarray.GetLength(0) - 1
        Dim splits(nsplits, 1) As String
        For i = 1 To nsplits

            splits(i, 0) = rewritesplits(spltarray(i, 1), ntax)

            splits(i, 1) = splitsize(splits(i, 0))

        Next

        
        Return splits


    End Function
    Function convert(ByVal split As String) As String
        Dim a As String
        For i = 0 To split.Length - 1
            If split.Substring(i, 1) = "*" Then
                a = a & "."
            Else
                a = a & "*"


            End If
        Next
        Return a
    End Function
    Function splitsize(ByVal a As String) As Integer
        Dim b As Integer
        If a <> Nothing Then
            For i = 0 To a.Length - 1
                If a.Substring(i, 1) = "*" Then
                    b = b + 1

                End If
            Next
        End If
        Return b
    End Function
    Public Function rewritesplits(ByVal split As String, ByVal ntax As Integer) As String
        Dim cerosplit As String = Nothing
        Dim i As Integer
        Dim g As Integer = 0



        For i = 1 To ntax

            If split.Contains(" " & i & " ") = True Then

                cerosplit = cerosplit & "*"
            Else
                cerosplit = cerosplit & "."
            End If

        Next i




        Return cerosplit
    End Function
    Function writesplits(ByVal splits As String(,), ByVal notus1 As String(,), ByVal otunames As String(,), ByVal nodo() As String) As String
        Dim nodearray(splits.GetLength(0) - 1, 2) As String
        Dim tree As String
        Dim minsize As Integer = 1

        Dim j As Integer = 0


        For j = 1 To splits.GetLength(0) - 1
            Dim chara As String


            chara = "*"


            Dim branchl As Single
            Dim branchS As String
            branchl = notus1(j, 0)
            If branchl = 0 Then
                branchS = 0
            Else
                branchS = notus1(j, 1)
            End If

            If splits(j, 1) = 1 Then '

                nodearray(j, 0) = otunames(splits(j, 0).IndexOfAny(chara) + 1, 0) & ":" & branchl



            Else
                Dim n1 As Integer = Array.IndexOf(nodo, nodo(j) & "1")

                nodearray(j, 0) = "(" & nodearray(n1, 0)
                Dim nodo2 As String = nodo(j)
                Dim n2 As Integer = Array.IndexOf(nodo, nodo2 & "2")

                Do
                    n2 = Array.IndexOf(nodo, nodo2 & "2")
                    If n2 = -1 Then
                        nodearray(j, 0) = nodearray(j, 0) & "," & nodearray(Array.IndexOf(nodo, nodo2 & "21"), 0)
                        nodo2 = nodo2 & "2"
                    Else
                        If branchS <> Nothing Then
                            nodearray(j, 0) = nodearray(j, 0) & "," & nodearray(n2, 0) & ")" & branchS & ":" & branchl
                        Else
                            nodearray(j, 0) = nodearray(j, 0) & "," & nodearray(n2, 0) & "):" & branchl
                        End If
                    End If
                Loop While n2 = -1




            End If


        Next
        minsize = minsize + 1
        Dim lastind As Integer = last(splits)
        Dim ind As Integer = compindex(splits(lastind, 0), splits)

        tree = "(" & nodearray(ind, 0) & "," & nodearray(lastind, 0) & ");"





        Return tree
    End Function
    Function last(ByVal splits As String(,)) As Integer
        Dim index As Integer
        For i = 1 To splits.GetLength(0) - 1
            If splits(i, 0) = Nothing Then
                index = i - 1
            End If
        Next
        If index = 0 Then
            index = splits.GetLength(0) - 1
        End If
        Return index
    End Function
    Function compindex(ByVal split As String, ByVal splites(,) As String) As Integer
        Dim index As Integer

        split = convert(split)
        For i = 0 To splites.GetLength(0) - 1
            If splites(i, 0) = split Then
                index = i
            End If
        Next
        Return index
    End Function
    Function codesplit(ByVal split As String, ByVal chara As String) As String
        Dim st As String = " "
        For i = 0 To split.Length - 1
            If split.Substring(i, 1) = chara Then
                st = st & i & " "
            End If
        Next
        Return st
    End Function
    Function tosplit(ByVal tree As String) As splits
        Dim sp As splits

        Dim ntax, nsplits As Integer

        For Each c In tree
            If StrComp(c, ",") = 0 Then
                ntax = ntax + 1
            End If
        Next c
        ntax = ntax + 1
        nsplits = (ntax - 1) * 2
        Dim arr(nsplits) As String
        ReDim sp.nOTUs(nsplits, 2)
        ReDim sp.notus1(nsplits, 1)
        sp.otumat1 = readnames(ntax, tree)
        arr = readsplits(nsplits, tree)
        For i = 1 To ntax
            sp.nOTUs(i, 1) = " " & i & " "
            sp.nOTUs(i, 2) = stdsplit(" " & i & " ", ntax)
            Dim inic As Integer = tree.IndexOf(sp.otumat1(i, 0))
            If tree.Substring(inic + sp.otumat1(i, 0).Length, 1) = ":" Then
                Dim start As Integer = inic + sp.otumat1(i, 0).Length + 1
                Dim nextcoma As Integer = tree.IndexOf(",", start)
                Dim nextpar As Integer = tree.IndexOf(")", start)
                Dim endnu As Integer

                If nextcoma < nextpar And nextcoma <> -1 Then
                    endnu = nextcoma
                Else
                    endnu = nextpar

                End If

                sp.notus1(i, 0) = tree.Substring(start, endnu - start)

            End If

        Next
        For i = ntax + 1 To nsplits
            sp.nOTUs(i, 1) = " "
            For j = 1 To ntax
                If arr(i).Contains(sp.otumat1(j, 0) & ":") Then
                    sp.nOTUs(i, 1) = sp.nOTUs(i, 1) & j & " "
                End If
            Next
            sp.nOTUs(i, 2) = stdsplit(sp.nOTUs(i, 1), ntax)

            Dim inic As Integer = tree.IndexOf(arr(i))
            If tree.Substring(inic + arr(i).Length + 1, 1) = ":" Then
                Dim start As Integer = inic + arr(i).Length + 2
                Dim nextcoma As Integer = tree.IndexOf(",", start)
                Dim nextpar As Integer = tree.IndexOf(")", start)
                Dim endnu As Integer

                If nextcoma < nextpar And nextcoma <> -1 Then
                    endnu = nextcoma
                Else
                    endnu = nextpar

                End If


                sp.notus1(i, 0) = tree.Substring(start, endnu - start)
            Else
                Dim start0 As Integer = inic + arr(i).Length + 1
                Dim nextdospuntos As Integer = tree.IndexOf(":", start0)
                Dim start1 As Integer = nextdospuntos + 1
                Dim nextcoma As Integer = tree.IndexOf(",", start0)
                Dim nextpar As Integer = tree.IndexOf(")", start0)
                Dim endnu As Integer

                If nextcoma < nextpar And nextcoma <> -1 Then
                    endnu = nextcoma
                Else
                    endnu = nextpar

                End If

                sp.notus1(i, 1) = tree.Substring(start0, nextdospuntos - start0)
                sp.notus1(i, 0) = tree.Substring(start1, endnu - start1)
            End If


        Next
        Return sp
    End Function
    Function readnames(ByVal ntax As Integer, ByVal tree As String) As String(,)
        Dim otumatx(ntax, 1) As String
        Dim start As Boolean

        Dim index As Integer = 1
        For i = 0 To tree.Length - 2
            Dim c As Char = tree.Substring(i, 1)
            Dim c1 As Char = tree.Substring(i + 1, 1)
            If start = False Then
                If c = "(" And c1 <> "(" Then
                    start = True
                ElseIf c = "," And c1 <> "(" Then
                    start = True
                Else
                    start = False

                End If
            Else
                If c = ":" Or c = ")" Or c = "," Then
                    index = index + 1
                    start = False
                Else

                    otumatx(index, 0) = otumatx(index, 0) & c
                End If
            End If

        Next

        Return otumatx

    End Function
    Function readsplits(ByVal nsplits As Integer, ByVal tree As String) As String()
        Dim arr(nsplits) As String
        Dim iarr As Integer = nsplits
        For i = 1 To tree.Length - 2
            Dim c As Char = tree.Substring(i, 1)
            Dim count As Integer = 0
            If c = "(" Then
                count = 1

                Dim fin As Integer
                For j = i + 1 To tree.Length
                    Dim cf As Char = tree.Substring(j, 1)
                    If cf = "(" Then
                        count = count + 1
                    ElseIf cf = ")" Then
                        count = count - 1
                    End If
                    If count = 0 Then
                        fin = j
                        Exit For
                    End If


                Next

                arr(iarr) = tree.Substring(i, fin - i)
                iarr = iarr - 1
            End If

        Next


        Return arr


    End Function
End Class
