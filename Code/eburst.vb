Imports System.Math
Public Structure geburst

    Dim matg As String(,)
    Dim initialsize As Integer
    Dim founder As String
    Dim niso As Integer
    Dim gdef As Integer
    Dim matg1 As Single(,)
    Dim centerx, centery As Single
    Dim order As Integer
End Structure
Public Structure reburst
    Dim gs As geburst()
    Dim otunames As String(,)
    Dim perfiles As Integer(,)
    Dim mat As String(,)
    Dim matx As String(,)
End Structure
Module eburst
    Public Function makeburstgraph(ByVal profiles(,) As Integer, ByVal otunames As String(,), ByVal seburst As Boolean, ByVal ndif As Integer, ByVal netburst As Boolean, ByRef gruburst As reburst, Optional ByVal netsize As Integer = 4) As splits
        profiles = reduceDST(profiles)
        Dim textgroup As String
        Dim enter = Environment.NewLine

        Dim group As String

        gruburst.perfiles = profiles
        gruburst.otunames = otunames
        Dim anothergroup As Boolean = True
        Dim i As Integer = 0
        ReDim gruburst.gs(0)
        Dim ini, fin As Integer
        If seburst = True Then
            ini = 1
            fin = profiles.GetLength(1) - 1
        Else
            ini = ndif
            fin = ndif
        End If
        For j = ini To fin

            Dim a(9) As Single

            a(6) = 9999999
            Form1.TextBox6.Text = Form1.TextBox6.Text & "-----------------------------" & enter & "GROUP DEFINITION: Shared Alleles " & profiles.GetLength(1) - 1 - j & "/" & profiles.GetLength(1) - 1 & enter & "-----------------------------" & enter & enter
            gruburst.matx = Nothing
            gruburst.matx = makeslv(profiles, j)
            anothergroup = True
            Do While anothergroup = True

                gruburst.gs(i) = makegroup(gruburst.gs(0), i + 1, otunames, gruburst.matx)
                gruburst.gs(i).gdef = j
                Dim check As Boolean = True
                For k = 0 To gruburst.gs.Length - 2
                    If checke(gruburst.gs(gruburst.gs.Length - 1), gruburst.gs(k)) = False Then
                        check = False
                        k = 1000000
                    End If
                Next
                Dim maxorder As Integer = 0
                For k = 0 To gruburst.gs.Length - 2
                    If testsuperpose(gruburst.gs(gruburst.gs.Length - 1), gruburst.gs(k)) = True Then
                        If gruburst.gs(k).order > maxorder Then
                            maxorder = gruburst.gs(k).order
                        End If
                    End If
                Next
                gruburst.gs(gruburst.gs.Length - 1).order = maxorder + 1
                If seburst = True Then
                    a = statburst(gruburst, a)
                End If

                If check = True Then
                    Dim tab As String = Char.ConvertFromUtf32(Keys.Tab)
                    If seburst = False Then
                        textgroup = "Group " & i + 1 & ": N° of Strains:" & gruburst.gs(i).niso & " / Number of ST:" & gruburst.gs(i).matg.GetLength(0) - 1 & " / Predicted Founder:" & gruburst.gs(i).founder & enter & enter
                        If netburst = False Then
                            textgroup = textgroup & "ST" & tab & "FREQ" & tab & "STRAINS" & enter

                            For m = 1 To gruburst.gs(i).matg.GetLength(0) - 1
                                textgroup = textgroup & gruburst.gs(i).matg(m, 0) & tab & countst(gruburst.gs(i).matg(m, 0), otunames) & tab & gruburst.gs(i).matg(m, 2) & enter

                            Next
                            textgroup = textgroup & enter & enter
                        End If

                        Form1.TextBox6.Text = Form1.TextBox6.Text & textgroup
                        If netburst = True Then

                            testNetBurst(gruburst, i + 1, netsize)
                        End If
                    Else

                    End If

                    gruburst = addone(gruburst)
                    i = i + 1
                End If


                anothergroup = testfin(gruburst.matx)
            Loop
            Dim nisolates As Integer = gruburst.otunames.GetLength(0)
            Dim nsts As Integer = gruburst.perfiles.GetLength(0) - 1
            If seburst = True Then
                a(0) = j
                a(7) = a(2) / a(0)
                a(2) = a(2) / nisolates
                a(4) = (nsts - a(3)) / nsts
                a(3) = a(3) / nsts
                textgroup = Nothing


                textgroup = textgroup & "Number of complexes: " & Math.Round(a(1)) & vbNewLine
                textgroup = textgroup & "Complex size range: (" & Math.Round(a(6)) & "," & Math.Round(a(5)) & ")" & vbNewLine
                textgroup = textgroup & "Complex mean size: " & Math.Round(a(7), 1) & vbNewLine

                textgroup = textgroup & "proportion of Isolates clustering in a complex: " & Math.Round(a(2), 2) & vbNewLine
                textgroup = textgroup & "proportion of STs clustering in a complex: " & Math.Round(a(3), 2) & vbNewLine
                textgroup = textgroup & "proportion of Singleton STs: " & Math.Round(a(4), 2) & vbNewLine
                textgroup = textgroup & vbNewLine & vbNewLine
                Form1.TextBox6.Text = Form1.TextBox6.Text & textgroup

            End If

        Next

        Dim sp As splits
        If seburst = True Then
            ReDim sp.nOTUs(gruburst.gs.Length + profiles.GetLength(0) - 3, 2)
            ReDim sp.notus1(gruburst.gs.Length + profiles.GetLength(0) - 3, 1)
            ReDim sp.otumat1(profiles.GetLength(0) - 1, 1)
            For i = 1 To profiles.GetLength(0) - 1
                sp.nOTUs(i, 1) = " " & i & " "
                sp.nOTUs(i, 2) = Module1.stdsplit(sp.nOTUs(i, 1), profiles.GetLength(0) - 1)
                sp.otumat1(i, 0) = "ST" & i & "{" & writeisolates(i, otunames, "_") & "}"
                sp.notus1(i, 0) = findbl(i, gruburst)
            Next
            For i = 0 To gruburst.gs.Length - 3
                Dim st As String = " "
                For j = 1 To gruburst.gs(i).matg.GetLength(0) - 1
                    st = st & gruburst.gs(i).matg(j, 0) & " "
                Next
                sp.nOTUs(i + profiles.GetLength(0), 1) = st
                sp.nOTUs(i + profiles.GetLength(0), 2) = Module1.stdsplit(st, profiles.GetLength(0) - 1)

                sp.notus1(i + profiles.GetLength(0), 1) = profiles.GetLength(1) - 1 - gruburst.gs(i).gdef
                Dim index As Integer = findindexbigger(gruburst, gruburst.gs(i).matg(1, 0), gruburst.gs(i).gdef)
                If index = 0 Then
                    sp.notus1(i + profiles.GetLength(0), 0) = 0
                Else
                    sp.notus1(i + profiles.GetLength(0), 0) = gruburst.gs(index).gdef - gruburst.gs(i).gdef
                End If
            Next
        Else
            Dim BURSTGRAPH As New treeviewer2 With {.GRUBURST = gruburst, .Text = "BURST Graph"}
            BURSTGRAPH.Show()
        End If

        Return sp

    End Function
    Function statburst(ByVal gruburst As reburst, ByVal stats() As Single) As Single()

        Dim nsts As Integer = gruburst.otunames.GetLength(0) - 1
        Dim nisolates As Integer = gruburst.perfiles.GetLength(0) - 1
       
        stats(1) = stats(1) + 1

        stats(2) = stats(2) + gruburst.gs(gruburst.gs.Length - 1).niso

        stats(3) = stats(3) + gruburst.gs(gruburst.gs.Length - 1).matg.GetLength(0) - 1

        If gruburst.gs(gruburst.gs.Length - 1).niso > stats(5) Then
            stats(5) = gruburst.gs(gruburst.gs.Length - 1).niso
        End If
        If gruburst.gs(gruburst.gs.Length - 1).niso < stats(6) Then
            stats(6) = gruburst.gs(gruburst.gs.Length - 1).niso
        End If

     

        Return stats

    End Function
    Function testNetBurst(ByVal gruburst As reburst, ByVal group As Integer,netsize As Integer)

        Dim matx(,) As String = gruburst.matx.Clone
        Dim nets(0) As String
        Dim treeChain As String = Nothing
        Dim count As Integer = 0
        For l = 1 To gruburst.gs(group - 1).matg.GetLength(0) - 1
            treeChain = "," & gruburst.gs(group - 1).matg(l, 0) & ","
            Dim looping As Boolean = True



            For i = 1 To matx.GetLength(0) - 1
                matx(i, 3) = 1
            Next
            Do While looping = True
                Dim endchain As Boolean = False
                Dim loopfound As Boolean = False
                Dim arr() As String = treeChain.Split(","c)
                Dim currentST As Integer = arr(arr.Length - 2)
                If matx(currentST, 1) = "1" Then
                    endchain = True
                    GoTo l1
                End If
                Dim nextSTs() As String = matx(currentST, 0).Split(" ")
                Dim nextst As Integer
                Dim index As Integer = matx(currentST, 3)
                If index > nextSTs.Length - 2 Then
                    endchain = True
                    GoTo l1
                End If
                nextst = nextSTs(index)

                If arr.Contains(nextst) Then

                    loopfound = True
                Else
                    treeChain = treeChain & nextst & ","
                End If
                If index <= matx(currentST, 1) Then
                    matx(currentST, 3) = matx(currentST, 3) + 1
                End If
                If loopfound = True Then



                    Dim I1 As Integer = Array.IndexOf(arr, nextst.ToString)
                    If arr.Length - 1 - I1 > 3 Then
                        If checkinnertriples(treeChain, nextst, gruburst) = False Then
                            Array.Resize(nets, nets.Length + 1)
                            nets(nets.Length - 1) = ","
                            For h = I1 To arr.Length - 2
                                nets(nets.Length - 1) = nets(nets.Length - 1) & arr(h) & ","
                            Next

                            nets(nets.Length - 1) = nets(nets.Length - 1) & nextst & ","

                            count = count + 1
                        End If
                    End If
                End If
                If arr.Length = netsize + 2 Then
                    treeChain = ","
                    For k = 1 To arr.Length - 2
                        treeChain = treeChain & arr(k) & ","
                    Next
                    For d = 1 To matx.GetLength(0) - 1
                        If treeChain.Contains("," & d.ToString & ",") = False Then
                            matx(d, 3) = "1"
                        End If
                    Next
                    GoTo l2
                End If
l1:
                If endchain = True Then
                    treeChain = ","
                    For k = 1 To arr.Length - 3
                        treeChain = treeChain & arr(k) & ","
                    Next
                    For d = 1 To matx.GetLength(0) - 1
                        If treeChain.Contains("," & d.ToString & ",") = False Then
                            matx(d, 3) = "1"
                        End If
                    Next
                    If treeChain <> "," Then
                        ' matx(arr(arr.Length - 3), 3) = matx(arr(arr.Length - 3), 3) + 1
                    Else
                        looping = False
                    End If


                End If
l2:
                'If treeChain.Contains("203") Or treeChain.Contains("238") Or treeChain.Contains("245") Or treeChain.Contains("248") Then Stop
            Loop
        Next
        For d = 1 To gruburst.gs(group - 1).matg.GetLength(0) - 1
            Dim i As Integer = gruburst.gs(group - 1).matg(d, 0)
            Dim arr1() As String = matx(i, 0).Split(" "c)

            For k = 1 To arr1.Length - 2
                Dim drawit As Boolean = False
                For j = 1 To nets.Length - 1

                    If nets(j).Contains("," & i.ToString & "," & arr1(k).ToString & ",") Or nets(j).Contains("," & arr1(k).ToString & "," & i.ToString & ",") Then
                        drawit = True
                    End If
                Next
                If drawit = False Then
                    arr1(k) = ""
                End If
            Next
            gruburst.matx(i, 0) = " "
            gruburst.matx(i, 2) = " "
            For l = 1 To arr1.Length - 2
                If arr1(l) <> "" Then
                    gruburst.matx(i, 0) = gruburst.matx(i, 0) & arr1(l) & " "
                    gruburst.matx(i, 2) = gruburst.matx(i, 2) & arr1(l) & " "
                End If
            Next
        Next
        Dim countnet As Integer = 0
        Dim impli As String = "Implicated STs: " & vbNewLine

        For i = 1 To gruburst.gs(group - 1).matg.GetLength(0) - 1
            Dim netst As Boolean = False
            For j = 1 To nets.Length - 1
                If nets(j).Contains("," & gruburst.gs(group - 1).matg(i, 0) & ",") = True Then
                    netst = True
                    Exit For
                End If

            Next
            If netst = True Then
                countnet = countnet + 1
                impli = impli & gruburst.gs(group - 1).matg(i, 0) & ", "
            End If


        Next
        If countnet > 0 Then
            Dim x As Integer
            If nets.Length - 1 < 10 Then
                x = nets.Length - 1
            Else
                x = 10
            End If
            Form1.TextBox6.Text = Form1.TextBox6.Text & "network structures detected, " & countnet & " from " & gruburst.gs(group - 1).matg.GetLength(0) - 1 & " STs implicated" & vbNewLine & vbNewLine
            Form1.TextBox6.Text = Form1.TextBox6.Text & impli & vbNewLine & "Examples:" & vbNewLine
            For i = 1 To x
                Form1.TextBox6.Text = Form1.TextBox6.Text & nets(i).Substring(1) & vbNewLine & vbNewLine
            Next
        Else
            Form1.TextBox6.Text = Form1.TextBox6.Text & vbNewLine & "No network structures detected" & vbNewLine & vbNewLine
        End If

    End Function
    Function checkinnertriples(ByVal treechain As String, ByVal lastst As Integer, ByVal gruburst As reburst) As Boolean
        '''''

        '''''
        Dim arr() As String = treechain.Split(","c)
        arr(arr.Length - 1) = lastst

        For i = 1 To arr.Length - 2

            Dim currentst As Integer = arr(i)

            Dim nextst As Integer
            If i = arr.Length - 2 Then
                nextst = arr(2)
            Else
                nextst = arr(i + 2)
            End If

            If gruburst.matx(currentst, 0).Contains(" " & nextst.ToString & " ") Then
                Return True
                Exit Function
            End If
        Next
        Return False

    End Function
    Function founderindex(ByVal array(,) As String, ByVal founder As String) As Integer
        Dim index As Integer = 0
        For i = 0 To array.GetLength(0) - 1
            If array(i, 0) = founder Then
                index = i
                Exit For
            End If
        Next
        Return index
    End Function
    Function findindexbigger(ByVal gruburst As reburst, ByVal st As Integer, ByVal gdef As Integer) As Integer
        Dim index As Integer

        For i = 0 To gruburst.gs.Length - 1
            If gruburst.gs(i).gdef > gdef Then
                For j = 1 To gruburst.gs(i).matg.GetLength(0) - 1

                    If st = gruburst.gs(i).matg(j, 0) Then
                        index = i
                        i = 10000000
                        j = 10000000
                    End If

                Next
            End If
        Next
        Return index
    End Function
    Function findbl(ByVal st As Integer, ByVal gruburst As reburst) As Integer
        Dim gdef As Integer

        For i = 0 To gruburst.gs.Length - 1
            For j = 1 To gruburst.gs(i).matg.GetLength(0) - 1
                If st = gruburst.gs(i).matg(j, 0) Then
                    gdef = gruburst.gs(i).gdef
                    j = 100000000
                    i = 100000000
                End If
            Next
        Next
        Return gdef
    End Function
    Function testsuperpose(ByVal burst1 As geburst, ByVal burst2 As geburst) As Boolean
        Dim contain As Boolean = False
        For i = 1 To burst1.matg.GetLength(0) - 1

            For j = 1 To burst2.matg.GetLength(0) - 1
                If burst1.matg(i, 0) = burst2.matg(j, 0) Then
                    contain = True
                    j = 1000000000
                    i = 1000000000
                End If
            Next

        Next
        Return contain
    End Function
    Function checke(ByVal burst1 As geburst, ByVal burst2 As geburst) As Boolean
        Dim distintas = False
        For i = 1 To burst1.matg.GetLength(0) - 1
            Dim contain As Boolean = False
            For j = 1 To burst2.matg.GetLength(0) - 1
                If burst1.matg(i, 0) = burst2.matg(j, 0) Then
                    contain = True
                    j = 1000000000
                End If
            Next
            If contain = False Then
                i = 1000000000
                distintas = True
            End If
        Next
        Return distintas
    End Function
    Function testfin(ByVal matx As String(,)) As Boolean
        Dim countx, countg As Integer
        For i = 1 To matx.GetLength(0) - 1
            If matx(i, 1) <> "0" Then
                countg = countg + 1
            End If
            If matx(i, 3) = "X" Then
                countx = countx + 1
            End If
        Next
        Dim test As Boolean = True
        If countg <= countx Then test = False
        Return test
    End Function
    Function addone(ByVal groupi As reburst) As reburst
        Dim groupf As reburst
        groupf.otunames = groupi.otunames
        groupf.perfiles = groupi.perfiles
        groupf.matx = groupi.matx

        ReDim groupf.gs(groupi.gs.Length)
        For i = 0 To groupi.gs.Length - 1
            groupf.gs(i) = groupi.gs(i)

        Next
        Return groupf
    End Function
    Function reduceDST(ByVal profiles(,) As Integer) As Integer(,)
        Dim nmaxdst As Integer = 0
        For i = 0 To profiles.GetLength(0) - 1
            If nmaxdst < profiles(i, profiles.GetLength(1) - 1) Then
                nmaxdst = profiles(i, profiles.GetLength(1) - 1)
            End If
        Next

        Dim prof(nmaxdst, profiles.GetLength(1) - 1) As Integer
        For i = 1 To nmaxdst
            For j = 0 To profiles.GetLength(1) - 1
                prof(i, j) = profiles(firstDST(profiles, i), j)

            Next
        Next
        Return prof
    End Function
    Function firstDST(ByVal profiles(,) As Integer, ByVal dst As Integer) As Integer
        Dim index As Integer
        For i = 0 To profiles.GetLength(0) - 1
            If profiles(i, profiles.GetLength(1) - 1) = dst Then
                index = i
                i = 1000000
            End If
        Next
        Return index
    End Function
    Function makeslv(ByVal profiles(,) As Integer, ByVal nmaxd As Integer) As String(,)
        Dim len As Integer = profiles.GetLength(0) - 1
        Dim matslv(len, 4) As String
        For i = 1 To len
            matslv(i, 2) = " "
            matslv(i, 0) = " "
            Dim count As Integer = 0
            Dim count1 As Integer = 0
            For j = 1 To len

                Dim ndif As Integer = 0
                For k = 0 To profiles.GetLength(1) - 2
                    If profiles(i, k) <> profiles(j, k) Then
                        ndif = ndif + 1
                        If ndif > nmaxd Then
                            k = 1000
                        End If

                    End If

                Next

                If ndif = 1 Then
                    matslv(i, 2) = matslv(i, 2) & j & " "
                    count1 = count1 + 1
                End If
                If ndif <= nmaxd And ndif <> 0 Then
                    matslv(i, 0) = matslv(i, 0) & j & " "

                    count = count + 1

                End If
            Next
            matslv(i, 1) = count
            matslv(i, 4) = count1
        Next
        Return matslv
    End Function

    Function makegroup(ByVal ebgroup As geburst, ByVal group As Integer, ByVal otunames As String(,), ByVal matx As String(,)) As geburst
        Dim count As Integer = 0
        Dim index As Integer = 0
        For i = 1 To matx.GetLength(0) - 1
            If matx(i, 3) = Nothing Then
                If matx(i, 1) > count Then
                    count = matx(i, 1)
                    index = i
                End If
            End If
        Next
        Dim groupsize As Integer = count + 1


        Dim resultp(groupsize) As String
        For i = 1 To groupsize
            If i = 1 Then
                resultp(i) = index
                matx(resultp(i), 3) = "X"
            Else

                resultp(i) = extractST(matx(index, 0), i - 1)
                matx(resultp(i), 3) = "X"
            End If
        Next
        Dim cont As Boolean = True

        Dim h As Integer = 1

        For i = 1 To matx.GetLength(0) - 1
            If resultp.Contains(i) = False Then

                If contains(resultp, matx(i, 0)) = True Then

                    Array.Resize(resultp, resultp.Length + 1)
                    resultp(resultp.Length - 1) = i
                    matx(i, 3) = "X"
                    i = 0
                End If
            End If
        Next




        Dim niso As Integer
        For i = 1 To resultp.Length - 1
            niso = niso + countst(resultp(i), otunames)

        Next
        Dim aburst As geburst
        aburst.founder = findfounder(resultp, matx)
        aburst.niso = niso
        aburst.initialsize = groupsize



        ReDim aburst.matg(resultp.Length - 1, 2)
        For i = 1 To resultp.Length - 1

            aburst.matg(i, 0) = resultp(i)
            aburst.matg(i, 1) = countst(resultp(i), otunames)
            aburst.matg(i, 2) = writeisolates(resultp(i), otunames, ", ")
        Next

        Return aburst
    End Function
    Function extractST(ByVal a As String, ByVal b As Integer) As String
        Dim count As Integer
        Dim read As Boolean
        Dim text As String

        For i = 0 To a.Length - 1
            If count = b And a.Substring(i, 1) <> " " Then
                text = text & a.Substring(i, 1)

            End If
            If a.Substring(i, 1) = " " Then
                count = count + 1


            End If
        Next
        Return text
    End Function
    Function contains(ByVal resultp() As String, ByVal a As String) As Boolean
        Dim res As Boolean = False
        For i = 1 To resultp.Length - 1
            Try
                If a.Contains(" " & resultp(i) & " ") = True Then
                    res = True
                    i = 100000000
                End If
            Catch ex As Exception

            End Try

        Next
        Return res
    End Function
    Function findfounder(ByVal resultp() As String, ByVal mat(,) As String) As String
        Dim t As String
        Dim max As Integer
        For i = 1 To resultp.Length - 1
            If mat(resultp(i), 4) > max Then
                t = resultp(i)
                max = mat(resultp(i), 4)
            ElseIf mat(resultp(i), 4) = max Then
                t = "None"
            End If
        Next
        Return t
    End Function
    Function countst(ByVal st As String, ByVal otunames As String(,)) As Integer
        Dim count As Integer
        For i = 0 To otunames.GetLength(0) - 1
            If otunames(i, 1) = st Then
                count = count + 1

            End If
        Next
        Return count
    End Function
    Function writeisolates(ByVal st As String, ByVal otunames As String(,), ByVal separator As String) As String
        Dim t As String
        For i = 0 To otunames.GetLength(0) - 1
            If otunames(i, 1) = st Then
                t = t & otunames(i, 0) & separator
            End If
        Next
        Return t
    End Function
End Module
