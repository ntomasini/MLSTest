Imports System.IO

Module Module2
    Private stopp As Boolean
    Public Property _stopp() As Integer
        Get
            Return stopp
        End Get
        Set(ByVal a As Integer)
            stopp = a
        End Set
    End Property
    Function UPGMAperfal(ByVal perfiles(,) As String, ByVal archivo As String, ByVal nrep As Integer) As splits
        arch = archivo

        nseq = perfiles.GetLength(0) - 2

        lseq = perfiles.GetLength(1) - 2



        Dim seqs(1) As String
        Dim nseq1 As Integer = nseq + 1
        Dim distancias(nseq1, nseq1) As Single
        Dim i, j As Integer
        Dim bootstrapn() As Integer

        Dim DSTarr(10) As Integer

        Dim splitest(10, 1) As String

        For i = 1 To nseq + 1

            splitest(10, 0) = splitest(10, 0) & i & " "


        Next

        For i = 0 To 10
            splitest(i, 1) = stdsplit(" " & splitest(i, 0) & " ", nseq + 1)

        Next

        Dim dist As Integer



        For i = 0 To nseq1

            For j = i + 1 To nseq1
                If i <> 0 Then
                    If i <> j Then

                        For z = 1 To lseq
                            If perfiles(i, z) <> perfiles(j, z) Then
                                dist = dist + 1
                            End If
                        Next
                        distancias(i, j) = dist
                        distancias(j, i) = distancias(i, j)
                        dist = 0
                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i
        For i = 0 To 10
            If splitest(i, 0) <> "" Then
                DSTarr(i) = comparedst(splitest(i, 0), distancias)
            End If
        Next

        distancias1 = distancias
        Dim sp As splits
        sp = UPGMAproc(nseq1, distancias, splitest, DSTarr, perfiles, True, nrep)
        Return sp

    End Function

    Function UPGMA(ByVal otumat1(,) As String, ByVal archivo As String, ByVal nrep As Integer) As splits

        arch = archivo

        nseq = otumat1.GetLength(0) - 2

        lseq = otumat1(1, 1).ToString.Length

        otumat1 = reductor(otumat1)
        If Form1.hethandpp = 1 Then
            otumat1 = duplicate(otumat1, otumat1(1, 1).Length, nseq)
            lseq = otumat1(1, 1).ToString.Length
        End If
        lseqred = otumat1.GetValue(1, 1).ToString.Length
        Dim sp As splits
        sp = UPGMAgo(otumat1, arch, nrep)
        Return sp
    End Function



    Function min(ByVal a As Integer, ByVal arr(,) As Single) As Integer()
        Dim i As Integer
        Dim j As Integer
        Dim minarr(1) As Integer
        Dim m As Double = 10000
        For i = 1 To a
            For j = i + 1 To a
                If i <> j Then
                    If arr.GetValue(i, j) < m Then
                        m = arr(i, j)
                        minarr(0) = i
                        minarr(1) = j
                    End If
                End If
            Next j
        Next i
        Return minarr
    End Function
    Function delcero(ByVal a As Integer, ByVal arr(,) As Single, ByVal arr1() As Integer, ByVal totus() As Integer) As Single(,)
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer = 1
        Dim l As Integer = 0
        Dim array1(a - 1, a - 1) As Single

        For i = 1 To a
            l = 1
            For j = 1 To a

                If j = arr1(1) Or j = arr1(0) Then
                Else

                    array1(k, l) = arr(i, j)
                    array1(l, k) = arr(i, j)
                    If i = a Then
                        Dim dij As Single
                        Dim din As Single
                        Dim djn As Single
                        Dim Ti As Integer = totus(0)
                        Dim Tj As Integer = totus(1)


                        dij = arr(arr1(0), arr1(1))
                        din = arr(arr1(0), j)
                        djn = arr(arr1(1), j)
                        array1(a - 1, l) = (din * Ti + djn * Tj) / (Ti + Tj)
                        array1(l, a - 1) = (din * Ti + djn * Tj) / (Ti + Tj)
                        array1(0, a - 1) = arr(0, a) + 1
                        array1(a - 1, 0) = arr(0, a) + 1
                    End If
                    l = l + 1

                End If

            Next j
            If i = arr1(0) Or i = arr1(1) Then
            Else
                array1(k, 0) = arr(i, 0)
                array1(0, k) = arr(i, 0)

                k = k + 1
            End If
        Next i

        Return array1
    End Function

    Function delcero1(ByVal a As Integer, ByVal arr(,) As Single, ByVal arr1() As Integer) As Single(,)
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer = 1
        Dim l As Integer = 1
        Dim array1(a - 1, a - 1) As Single

        For i = 1 To a
            l = k + 1
            For j = i + 1 To a

                If j = arr1.GetValue(1) Then
                Else

                    array1.SetValue(arr.GetValue(i, j), k, l)
                    array1.SetValue(arr.GetValue(i, j), l, k)

                    l = l + 1

                End If

            Next j
            If i <> arr1.GetValue(0) Then

                array1.SetValue(arr.GetValue(i, 0), k, 0)
                array1.SetValue(arr.GetValue(i, 0), 0, k)
                k = k + 1
            End If

        Next i

        Return array1
    End Function
    Function otumat(ByVal file As String, ByVal nseq As Integer) As String(,)
        Dim lector As TextReader
        lector = New StreamReader(file)
        Dim otu(nseq + 1, 1) As String

        Dim lin As String
        Dim i As Integer = 0
        lin = "1"


        Do While lin <> Nothing
            lin = lector.ReadLine()
            If lin <> Nothing Then
                If lin.Substring(0, 1) = ">" Then

                    otu(i + 1, 0) = lin.Substring(1, lin.Length - 1)
                    i = i + 1
                Else
                    otu(i, 1) = lin.Substring(0, lin.Length - 1)
                End If
            End If
        Loop
        lector.Close()

        Return otu
    End Function
    Function Udistonode(ByVal mini() As Integer, ByVal distancias(,) As Single, ByVal dacum As Single) As Single
        Dim dist As Single
        dist = (distancias.GetValue(mini) / 2) - dacum
        Return dist
    End Function
    Function UPGMAgo(ByVal otumat1(,) As String, ByVal arch As String, ByVal nrep As Integer) As splits


        Dim seqs(1) As String
        Dim nseq1 As Integer = nseq + 1
        Dim distancias(nseq1, nseq1) As Single
        Dim i, j As Integer
        Dim bootstrapn() As Integer
        Dim dist As medirdistancia
        Dim DSTarr(10) As Integer

        Dim splitest(,) As String = Module1.splitestmaker


        For i = 1 To nseq + 1

            splitest(splitest.GetLength(0) - 1, 0) = splitest(splitest.GetLength(0) - 1, 0) & i & " "


        Next


        Dim sequ1, sequ2 As String

        dist = New medirdistancia


        For i = 0 To nseq1

            For j = i + 1 To nseq1
                If i <> 0 Then
                    If i <> j Then
                        seqs(0) = otumat1(i, 1)
                        seqs(1) = otumat1(j, 1)

                        sequ1 = seqs(0)
                        sequ2 = seqs(1)

                        distancias(i, j) = dist.distance(sequ1, sequ2, lseq, False, 0, 1, bootstrapn)
                        distancias(j, i) = distancias(i, j)

                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i

        For i = 0 To splitest.GetLength(0) - 1
            If splitest(i, 0) <> "" Then
                DSTarr(i) = comparedst(splitest(i, 0), distancias)

            End If
        Next

        distancias1 = distancias
        Dim sp As splits
        sp = UPGMAproc(nseq1, distancias, splitest, DSTarr, otumat1, False, nrep)
        Return sp

    End Function
    Function UPGMAproc(ByVal nseq1 As Integer, ByVal distancias(,) As Single, ByVal splitest(,) As String, ByVal DSTarr() As Integer, ByVal otumat1(,) As String, ByVal perfil As Boolean, ByVal nrep As Integer) As splits
        Dim mina1, mina2 As Integer
        Dim i, j As Integer
        Dim c As Integer = nseq1


        Dim sp1 As splits

        Dim texto As String
        Dim nseqx2 As Integer = nseq * 2
        ReDim sp1.nOTUs((nseqx2), 3)
        ReDim sp1.notus1(nseqx2, 2)


        Dim splits(nseq, 2) As String
        Dim h As Integer = 1

        Dim notu1final(2) As Single

        Dim nseqr As Integer
        nseqr = c
        h = 1


        Do While c > 1




            Dim min1(1) As Integer



            min1 = min(c, distancias)

            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))

            Dim nsplitcl1, nsplitcl2 As String
            Dim Totus(1) As Integer

            If mina1 <= nseq1 Then


                sp1.nOTUs(h, 1) = (" " & mina1 & " ")
                sp1.nOTUs(h, 0) = (mina1 & " ")
                sp1.nOTUs(h, 2) = stdsplit(sp1.nOTUs(h, 1), nseq1)
                Totus(0) = 1
                nsplitcl1 = " " & mina1 & " "

                sp1.notus1(h, 0) = Udistonode(min1, distancias, 0)
                sp1.notus1(h, 2) = sp1.notus1(h, 0)
                h = h + 1
            Else
                Dim pos As Integer = 0
                i = 1
                Do While pos = 0

                    If sp1.nOTUs(i, 0) = mina1 & " " Then
                        pos = i
                    End If
                    i = i + 1
                Loop
                nsplitcl1 = sp1.nOTUs(pos, 1)
                Totus(0) = countU(sp1.nOTUs(pos, 2), sp1.nOTUs(pos, 1), nseq1)
                sp1.notus1(pos, 0) = Udistonode(min1, distancias, sp1.notus1(pos, 2))
            End If


            If mina2 <= nseq1 Then


                sp1.nOTUs(h, 1) = " " & mina2 & " "
                sp1.nOTUs(h, 0) = mina2 & " "
                sp1.nOTUs(h, 2) = stdsplit(sp1.nOTUs(h, 1), nseq1)
                nsplitcl2 = " " & mina2 & " "
                sp1.notus1(h, 0) = Udistonode(min1, distancias, 0)
                Totus(1) = 1

                h = h + 1
            Else

                Dim pos As Integer = 0
                i = 1
                Do While pos = 0

                    If sp1.nOTUs(i, 0) = mina2 & " " Then
                        pos = i
                    End If
                    i = i + 1
                Loop


                nsplitcl2 = sp1.nOTUs(pos, 1)
                Totus(1) = countU(sp1.nOTUs(pos, 2), sp1.nOTUs(pos, 1), nseq1)
                sp1.notus1(pos, 0) = Udistonode(min1, distancias, sp1.notus1(pos, 2))
            End If
            If c > 2 Then

                sp1.nOTUs(h, 1) = nsplitcl1.Substring(0, nsplitcl1.Length - 1) & nsplitcl2
                sp1.nOTUs(h, 0) = nseq + 1 + (nseq + 2 - c) & " "
                sp1.nOTUs(h, 2) = stdsplit(sp1.nOTUs.GetValue(h, 1), nseq1)

                'sp1.nOTUs1(h, 0) = Udistonode(min1, c, distancias)
                sp1.notus1(h, 2) = distancias(min1(0), min1(1)) / 2

                h = h + 1


                distancias = delcero(c, distancias, min1, Totus)
            End If

            c = c - 1

        Loop


        Dim seed As Integer = nrep
        Dim seed1 As Double = Rnd(-seed)
        Form1.ProgressBar1.Maximum = nrep
        Form1.ProgressBar1.Value = 0
        stopp = False
        Dim bootstrapbyte() As Integer
        If nrep <> 0 Then
            For i = 1 To nrep
                If perfil = False Then
                    bootstrapbyte = ubootstraping(sp1.nOTUs, Rnd(), otumat1, nseq)
                Else
                    bootstrapbyte = ubootsperfal(sp1.nOTUs, Rnd(), otumat1, nseq)
                End If
                For j = 1 To nseq * 2
                    sp1.notus1(j, 1) = bootstrapbyte(j) + sp1.notus1(j, 1)
                Next j
                Form1.ProgressBar1.Increment(1)
                Application.DoEvents()
                If stopp = True Then
                    sp1.otumat1 = otumat1

                    sp1.splitest = splitest
                    sp1.DSTarr = DSTarr
                    Return sp1
                    Exit Function
                End If
            Next i


        End If
        sp1.otumat1 = otumat1

        sp1.splitest = splitest
        sp1.DSTarr = DSTarr

        Return sp1

    End Function

    Function countU(ByVal a As String, ByVal b As String, ByVal seqs As Integer) As Integer
        Dim i As Integer = 0
        Dim n As Integer = 0
        If b.Contains(" 1 ") Then
            Do While i + 1 < a.Length - 1

                If a.Substring(i + 1, 1) = " " Then
                    n = n + 1
                    i = i + 1
                Else

                    i = i + 1
                End If
            Loop
        Else
            n = 0
            Do While i + 1 < a.Length - 1

                If a.Substring(i + 1, 1) = " " Then
                    n = n + 1
                    i = i + 1
                Else

                    i = i + 1
                End If


            Loop
            n = seqs - n
        End If
        Return n
    End Function

    Function ubootstraping(ByVal bnotus(,) As String, ByVal seed As Double, ByVal otumat1(,) As String, ByVal nseq As Integer) As Integer()
        Dim mina1, mina2 As Integer
        Dim seqs(1) As String
        Dim bootstraprnd() As Integer
        Dim nseq1 As Integer = nseq + 1
        Dim nseqx2 As Integer = nseq * 2
        Dim distancias(nseq1, nseq1) As Single
        Dim i, j As Integer

        Dim dist As medirdistancia

        Dim sequ1, sequ2 As String
        j = 1
        dist = New medirdistancia
        Dim cadena As String
        Dim g As Integer
        Dim k As Integer = 0

        For i = 1 To lseq
            g = CInt(Int(Rnd() * (lseq - 1)))
            If g < lseqred Then
                Array.Resize(bootstraprnd, k + 1)
                bootstraprnd(k) = g
                k = k + 1
            End If
        Next

        For i = 0 To nseq1

            For j = i + 1 To nseq1
                If i <> 0 Then
                    If i <> j Then
                        seqs(0) = otumat1(i, 1)
                        seqs(1) = otumat1(j, 1)

                        sequ1 = seqs(0)
                        sequ2 = seqs(1)
                        If distancias1(i, j) <> 0 Then
                            distancias(i, j) = dist.distance(sequ1, sequ2, lseq, True, 0, distancias1(i, j), bootstraprnd)
                            distancias(j, i) = distancias(i, j)
                        End If
                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i

        Dim c As Integer = nseq1

        '  Do While distancias.GetValue(min(c, distancias)) = 0
        'distancias = delcero1(c, distancias, min(c, distancias))

        ' c = c - 1
        'Loop



        Dim nOTUS((nseqx2), 2) As String
        Dim nOtus1(nseqx2, 2) As Single
        Dim splits(nseq, 2) As String
        Dim h As Integer = 1
        Dim nseqr As Integer = c

        Do While c > 1




            Dim min1(1) As Integer



            min1 = min(c, distancias)

            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))

            Dim nsplitcl1, nsplitcl2 As String
            Dim Totus(1) As Integer

            If mina1 <= nseq1 Then


                nOTUS(h, 1) = (" " & mina1 & " ")
                nOTUS(h, 0) = (mina1 & " ")
                nOTUS(h, 2) = stdsplit(nOTUS(h, 1), nseq1)
                Totus(0) = 1
                nsplitcl1 = " " & mina1 & " "

                nOtus1(h, 0) = Udistonode(min1, distancias, 0)
                nOtus1(h, 2) = nOtus1(h, 0)
                h = h + 1
            Else
                Dim pos As Integer = 0
                i = 1
                Do While pos = 0

                    If nOTUS(i, 0) = mina1 & " " Then
                        pos = i
                    End If
                    i = i + 1
                Loop
                nsplitcl1 = nOTUS(pos, 1)
                Totus(0) = countU(nOTUS(pos, 2), nOTUS(pos, 1), nseq1)
                nOtus1(pos, 0) = Udistonode(min1, distancias, nOtus1(pos, 2))
            End If


            If mina2 <= nseq1 Then


                nOTUS(h, 1) = " " & mina2 & " "
                nOTUS(h, 0) = mina2 & " "
                nOTUS(h, 2) = stdsplit(nOTUS(h, 1), nseq1)
                nsplitcl2 = " " & mina2 & " "
                nOtus1(h, 0) = Udistonode(min1, distancias, 0)
                Totus(1) = 1

                h = h + 1
            Else

                Dim pos As Integer = 0
                i = 1
                Do While pos = 0

                    If nOTUS(i, 0) = mina2 & " " Then
                        pos = i
                    End If
                    i = i + 1
                Loop


                nsplitcl2 = nOTUS(pos, 1)
                Totus(1) = countU(nOTUS(pos, 2), nOTUS(pos, 1), nseq1)
                nOtus1(pos, 0) = Udistonode(min1, distancias, nOtus1(pos, 2))
            End If
            If c > 2 Then

                nOTUS(h, 1) = nsplitcl1.Substring(0, nsplitcl1.Length - 1) & nsplitcl2
                nOTUS(h, 0) = nseq + 1 + (nseq + 2 - c) & " "
                nOTUS(h, 2) = stdsplit(nOTUS.GetValue(h, 1), nseq1)

                'nOTUs1(h, 0) = Udistonode(min1, c, distancias)
                nOtus1(h, 2) = distancias(min1(0), min1(1)) / 2

                h = h + 1


                distancias = delcero(c, distancias, min1, Totus)
            End If

            c = c - 1

        Loop
        Dim bootstrap(nseqx2) As Integer
        i = 1
        j = 1
        For i = 1 To (nseqx2)
            For j = 1 To nseqx2
                If bnotus(i, 2) = nOTUS(j, 2) And nOtus1(j, 0) <> 0 Then
                    bootstrap(i) = 1
                    j = (nseqx2)
                End If
            Next
        Next
        Return bootstrap
    End Function
    Function ubootsperfal(ByVal bnotus(,) As String, ByVal seed As Double, ByVal perfiles(,) As String, ByVal nseq As Integer) As Integer()
        Dim mina1, mina2 As Integer
        Dim seqs(1) As String
        Dim bootstraprnd() As Integer
        Dim nseq1 As Integer = nseq + 1
        Dim nseqx2 As Integer = nseq * 2
        Dim distancias(nseq1, nseq1) As Single
        Dim i, j As Integer


        Dim sequ1, sequ2 As String
        j = 1

        Dim cadena As String
        Dim g As Integer
        Dim k As Integer = 0

        For i = 1 To lseq
            g = CInt(Int((Rnd() * (lseq - 1)) + 1))

            Array.Resize(bootstraprnd, k + 1)
            bootstraprnd(k) = g
            k = k + 1

        Next
        Dim dist As Integer

        For i = 0 To nseq1

            For j = i + 1 To nseq1
                If i <> 0 Then
                    If i <> j Then


                        For z = 1 To lseq
                            If perfiles(i, bootstraprnd(z - 1)) <> perfiles(j, bootstraprnd(z - 1)) Then
                                dist = dist + 1
                            End If
                        Next
                        distancias(i, j) = dist
                        distancias(j, i) = dist
                        dist = 0
                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i

        Dim c As Integer = nseq1

        '  Do While distancias.GetValue(min(c, distancias)) = 0
        'distancias = delcero1(c, distancias, min(c, distancias))

        ' c = c - 1
        'Loop



        Dim nOTUS((nseqx2), 2) As String
        Dim nOtus1(nseqx2, 2) As Single
        Dim splits(nseq, 2) As String
        Dim h As Integer = 1
        Dim nseqr As Integer = c

        Do While c > 1




            Dim min1(1) As Integer



            min1 = min(c, distancias)

            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))

            Dim nsplitcl1, nsplitcl2 As String
            Dim Totus(1) As Integer

            If mina1 <= nseq1 Then


                nOTUS(h, 1) = (" " & mina1 & " ")
                nOTUS(h, 0) = (mina1 & " ")
                nOTUS(h, 2) = stdsplit(nOTUS(h, 1), nseq1)
                Totus(0) = 1
                nsplitcl1 = " " & mina1 & " "

                nOtus1(h, 0) = Udistonode(min1, distancias, 0)
                nOtus1(h, 2) = nOtus1(h, 0)
                h = h + 1
            Else
                Dim pos As Integer = 0
                i = 1
                Do While pos = 0

                    If nOTUS(i, 0) = mina1 & " " Then
                        pos = i
                    End If
                    i = i + 1
                Loop
                nsplitcl1 = nOTUS(pos, 1)
                Totus(0) = countU(nOTUS(pos, 2), nOTUS(pos, 1), nseq1)
                nOtus1(pos, 0) = Udistonode(min1, distancias, nOtus1(pos, 2))
            End If


            If mina2 <= nseq1 Then


                nOTUS(h, 1) = " " & mina2 & " "
                nOTUS(h, 0) = mina2 & " "
                nOTUS(h, 2) = stdsplit(nOTUS(h, 1), nseq1)
                nsplitcl2 = " " & mina2 & " "
                nOtus1(h, 0) = Udistonode(min1, distancias, 0)
                Totus(1) = 1

                h = h + 1
            Else

                Dim pos As Integer = 0
                i = 1
                Do While pos = 0

                    If nOTUS(i, 0) = mina2 & " " Then
                        pos = i
                    End If
                    i = i + 1
                Loop


                nsplitcl2 = nOTUS(pos, 1)
                Totus(1) = countU(nOTUS(pos, 2), nOTUS(pos, 1), nseq1)
                nOtus1(pos, 0) = Udistonode(min1, distancias, nOtus1(pos, 2))
            End If
            If c > 2 Then

                nOTUS(h, 1) = nsplitcl1.Substring(0, nsplitcl1.Length - 1) & nsplitcl2
                nOTUS(h, 0) = nseq + 1 + (nseq + 2 - c) & " "
                nOTUS(h, 2) = stdsplit(nOTUS.GetValue(h, 1), nseq1)

                'nOTUs1(h, 0) = Udistonode(min1, c, distancias)
                nOtus1(h, 2) = distancias(min1(0), min1(1)) / 2

                h = h + 1


                distancias = delcero(c, distancias, min1, Totus)
            End If

            c = c - 1

        Loop
        Dim bootstrap(nseqx2) As Integer
        i = 1
        j = 1
        For i = 1 To (nseqx2)
            For j = 1 To nseqx2
                If bnotus(i, 2) = nOTUS(j, 2) And nOtus1(j, 0) <> 0 Then
                    bootstrap(i) = 1
                    j = (nseqx2)
                End If
            Next
        Next
        Return bootstrap
    End Function
End Module
