Imports System.IO

Public Structure splits
    Dim nOTUs(,) As String
    Dim notus1(,) As String
    Dim otumat1(,) As String
    Dim DSTarr() As Integer
    Dim splitest(,) As String
    Dim nsites As Integer
    Dim lseq As Integer
End Structure 'represent a tree in splits format
Public Structure dsx
    Dim distancias(,) As Single
    Dim Vij(,) As Single
    Dim DST() As Integer
End Structure 'represent a Distance matrix with the covariance matrix


Public Structure otumatxx
    Dim otumat1(,) As String
    Dim dista(,) As Single
    Dim treelength As Single
End Structure 'represent an alignment
Public Structure vectomat
    Dim matv() As Single
End Structure

Module Module1
    Public otum() As otumatxx
    Public filename As String
    Public nseq As Integer
    Public otumat1(,) As String
    Public lseq As Integer
    Public lseqred As Integer
    Public distancias1(,) As Single
    Public arch As String
    Private stopp As Boolean
    Private REPS As Integer
    Public Property _stopp() As Boolean
        Get
            Return stopp
        End Get
        Set(ByVal a As Boolean)
            stopp = a
        End Set
    End Property


    'NJ methods
    Function NJ(ByRef otumat1(,) As String, ByRef splitest(,) As String, ByVal nrep As Integer, ByVal bionj As Boolean, ByVal support As Boolean, ByVal suppforsplitest As Boolean) As splits


        nseq = otumat1.GetLength(0) - 2
        Dim flseq As Integer
        lseq = otumat1(1, 1).ToString.Length
        Dim sp4 As splits
        flseq = lseq
        Dim a(1) As Integer
        otumat1 = reductor(otumat1) 'delete constant sites
        If Form1.hethandpp = 1 Then
            otumat1 = duplicate(otumat1, otumat1(1, 1).Length, nseq)
            lseq = otumat1(1, 1).ToString.Length
        End If
        lseqred = otumat1(1, 1).ToString.Length
        Dim sp3 As dsx
        sp3 = NJgo(otumat1, bionj, False) 'calculates the distance matrix



        sp4 = NJproc(nseq + 1, sp3.distancias, otumat1, False, nrep, sp3.Vij, bionj) 'Makes the splits
        sp4.lseq = flseq
        '-----------------------------------------
        'Calculates clade significance if support is true
        If support = True Then
            Dim bar As Boolean

            If nrep > 0 Then
                bar = True
            End If

            For n = nseq + 2 To sp4.nOTUs.GetLength(0) - 2
                If suppforsplitest = True Then
                    If contiene(splitest, 1, sp4.nOTUs(n, 2)) = True Then
                        sp4.notus1(n, 1) = Math.Round(cladesupport(sp4.nOTUs(n, 2), sp4, sp3, otumat1) * 1000) / 1000
                    End If
                Else
                    Dim prv As String
                    If bar = True Then
                        prv = sp4.notus1(n, 1) & "/"
                    End If
                    sp4.notus1(n, 1) = prv & Math.Round(cladesupport(sp4.nOTUs(n, 2), sp4, sp3, otumat1) * 1000) / 1000
                End If
            Next
        End If

        sp4.DSTarr = DSTs(splitest, sp3.distancias)
        sp4.otumat1 = otumat1
        sp4.splitest = splitest
        Return sp4
    End Function

    Function NJproc(ByVal nseq1 As Integer, ByVal distancias(,) As Single, ByVal otumat1(,) As String, ByVal perfil As Boolean, ByVal nrep As Integer, ByVal Vij(,) As Single, ByVal bionj As Boolean) As splits
        Dim mina1, mina2 As Integer
        Dim i, j As Integer
        Dim c As Integer = nseq1



        distancias1 = distancias
        Dim sp1 As splits

        Dim nseqx2 As Integer = (nseq1 - 1) * 2
        ReDim sp1.nOTUs((nseqx2), 3)
        ReDim sp1.notus1(nseqx2, 1)


        Dim h As Integer = 1


        Dim nseqr As Integer
        nseqr = c
        h = 1
        Dim min1(1) As Integer
        Dim NOTUSMATX(nseqr) As String
        For V = 1 To NOTUSMATX.Length - 1
            NOTUSMATX(V) = V
            sp1.nOTUs(V, 1) = " " & V & " "
            sp1.nOTUs(V, 2) = stdsplit(sp1.nOTUs(V, 1), nseqr)
        Next
        Dim rmatx(c) As Single
        Do While c > 2

            Dim nmat(,) As Single
            ReDim nmat(c, c)
            Dim minrev(1) As Integer
            Dim diu, dju As Single
            If c = nseqr Then
                rmatx = rmat(c, distancias)
            End If
            nmat = Mmat(c, distancias, rmatx)
            min1 = min(c, nmat)
            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))
            minrev(0) = min1(1)
            minrev(1) = min1(0)



            '''
            diu = distonode(rmatx, min1, c, distancias)
            dju = distonode(rmatx, minrev, c, distancias)



            Dim lamb As Single = 0.5
            If bionj = True Then

                Dim suma As Single
                If Vij(min1(0), min1(1)) <> 0 Then
                    For w = 1 To c
                        If min1(0) <> w And min1(1) <> w And min1(0) <> min1(1) Then
                            suma = suma + (Vij(min1(0), w) - Vij(min1(1), w))
                        End If
                    Next
                    lamb = 0.5 + suma / (2 * (c - 2) * Vij(min1(0), min1(1)))
                End If
                If lamb > 1 Then lamb = 1
                If lamb < 0 Then lamb = 0
                Vij = biodelcero(c, Vij, min1, lamb)
            End If




            NOTUSMATX = Reducenotusmat(c, NOTUSMATX, min1)

            distancias = delcero(c, distancias, min1, lamb, diu, dju, rmatx)
            Dim posit As Integer = distancias(c - 1, 0)
            sp1.nOTUs(posit, 1) = " " & NOTUSMATX(NOTUSMATX.Length - 1) & " "

            sp1.nOTUs(posit, 2) = stdsplit(" " & NOTUSMATX(NOTUSMATX.Length - 1) & " ", nseqr)
            sp1.notus1(mina1, 0) = diu
            sp1.notus1(mina2, 0) = dju


            c = c - 1




        Loop
        Dim aa As Integer = sp1.notus1.GetLength(0) - 1
        If sp1.nOTUs(aa, 1) = Nothing Then
            For s = 1 To 2
                If distancias(s, 0) <= nseq1 Then
                    sp1.nOTUs(aa, 0) = distancias(s, 0) & " "
                    sp1.nOTUs(aa, 1) = " " & distancias(s, 0) & " "
                    sp1.nOTUs(aa, 2) = stdsplit(sp1.nOTUs(aa, 1), nseq1)
                End If
            Next
        End If
        sp1.notus1(distancias(2, 0), 0) = distancias(1, 2) / 2
        sp1.notus1(distancias(1, 0), 0) = distancias(1, 2) / 2




        If nrep <> 0 Then
            Dim seed As Integer = nrep
            Dim seed1 As Single = Rnd(-seed)
            Form1.ProgressBar1.Maximum = nrep
            Form1.ProgressBar1.Value = 0

            Dim bootstrapbyte() As Integer

            stopp = False
            For i = 1 To nrep
                If perfil = False Then
                    bootstrapbyte = bootstraping(sp1.nOTUs, Rnd(), otumat1, nseq1 - 1, bionj)

                Else
                    bootstrapbyte = bootsperfal(sp1.nOTUs, Rnd(), otumat1, nseq1 - 1)
                End If
                For j = nseq1 + 1 To (nseq1 - 1) * 2
                    If sp1.notus1(j, 0) > 10 ^ -9 Or sp1.notus1(j, 0) < -10 ^ -9 Then
                        'If j = 50 Then Stop
                        sp1.notus1(j, 1) = bootstrapbyte(j) + sp1.notus1(j, 1)
                    End If
                Next j
                Form1.ProgressBar1.Increment(1)
                Application.DoEvents()
                If stopp = True Then




                    For h = nseq1 + 1 To sp1.notus1.GetLength(0) - 1
                        If sp1.notus1(h, 1) <> Nothing Then
                            sp1.notus1(h, 1) = sp1.notus1(h, 1) / nrep * 100
                        End If
                    Next
                    sp1.otumat1 = otumat1

                    Return sp1
                    Exit Function
                End If
                'REPS = REPS - INV 
            Next i
        End If




        If nrep <> 0 Then
            For h = nseq + 1 To sp1.notus1.GetLength(0) - 1

                If sp1.notus1(h, 1) <> Nothing Then
                    sp1.notus1(h, 1) = sp1.notus1(h, 1) / nrep * 100
                End If

            Next
        End If
        sp1.otumat1 = otumat1
        Return sp1

    End Function
    Function NJperfal(ByVal perfiles(,) As String, ByVal otumat1(,) As String, ByVal archivo As String, ByVal nrep As Integer) As splits
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
                        distancias(i, j) = dist / 10
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
        sp = NJproc(nseq1, distancias, otumat1, True, nrep, Nothing, False)



        Return sp
    End Function 'NJ Method for Allelic Profiles, it return a tree in a split structure

    Function NJgo(ByRef otumat1(,) As String, ByVal bionj As Boolean, ByVal bootstrap As Boolean) As dsx
        Dim a As dsx
        Dim seqs(1) As String
        Dim nseq1 As Integer = otumat1.GetLength(0) - 1
        Dim distancias(nseq1, nseq1) As Single
        Dim Vij(nseq1, nseq1) As Single
        Dim i, j As Integer
        Dim bootstrapn() As Integer
        Dim dist As medirdistancia
        If bootstrap Then
            Dim g As Integer = 0
            Dim k As Integer = 0
            For i = 1 To lseq
                g = CInt(Int(Rnd() * (lseq - 1)))
                'If g < lseq Then
                Array.Resize(bootstrapn, k + 1)
                bootstrapn(k) = g
                k = k + 1
                ' End If
            Next
        End If

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

                        distancias(i, j) = dist.distance(sequ1, sequ2, lseq, bootstrap, 0, 1, bootstrapn)
                        distancias(j, i) = distancias(i, j)

                        If bionj = True Then
                            Vij(i, j) = distancias(i, j)
                            Vij(j, i) = Vij(i, j)
                        End If

                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i
        a.Vij = Vij

        a.distancias = distancias
        Return a


    End Function 'it calculates a distance matrix from an alignment



    'NJ algorithm Functions--------------------------------------
    Function rmat(ByVal a As Integer, ByVal arr(,) As Single) As Single()
        Dim i As Integer
        Dim array1(a) As Single
        Dim sum As Single
        Dim j As Integer
        Dim k As Integer
        For i = 1 To a
            sum = 0
            For j = 1 To a

                If i <> j Then

                    sum = sum + arr(i, j)

                    'sum = sum + arr(j, i)
                End If

            Next j

            array1(i) = sum
        Next i
        Return array1
    End Function 'Calculates the sum of the distance for each taxa
    Function Mmat(ByVal a As Integer, ByVal arr(,) As Single, ByVal arr2() As Single) As Single(,)
        Dim i As Integer
        Dim j As Integer
        Dim dij As Single = 0
        Dim Mij As Single = 0
        Dim ri As Single
        Dim rj As Single
        Dim array1(,) As Single

        ReDim array1(a, a)

        For i = 1 To a - 1

            For j = i + 1 To a

                dij = arr(i, j)
                ri = arr2(i)
                rj = arr2(j)
                Mij = dij - ((ri + rj) / (a - 2))

                array1(i, j) = Math.Truncate(Mij * 1000000) / 10000000
                'array1(j, i) = array1(i, j)


            Next j

        Next i

        Return array1
    End Function 'Calculates Q matrix
    Function min(ByVal a As Integer, ByVal arr(,) As Single) As Integer()
        Dim i As Integer
        Dim j As Integer
        Dim minarr(1) As Integer
        Dim m As Single = 10000000000
        For i = 1 To a
            For j = i + 1 To a
                If i <> j Then
                    If arr(i, j) < m Then
                        m = arr(i, j)
                        minarr(0) = i
                        minarr(1) = j
                    End If
                End If
            Next j
        Next i
        Return minarr
    End Function 'it return the pair that has the minimum value in the Q matrix
    Function delcero(ByVal a As Integer, ByVal arr(,) As Single, ByVal arr1() As Integer, ByVal lam As Single, ByVal diu As Single, ByVal dju As Single, ByVal rmatx() As Single) As Single(,)
        Dim rmatmod(rmatx.Length - 2) As Single


        Dim i As Integer
        Dim j As Integer
        Dim k As Integer = 1
        Dim l As Integer = 0
        Dim array1(,) As Single
        ReDim array1(a - 1, a - 1)

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

                        dij = arr(arr1(0), arr1(1))
                        din = arr(arr1(0), j)
                        djn = arr(arr1(1), j)
                        array1(a - 1, l) = lam * din + (1 - lam) * djn - lam * diu - (1 - lam) * dju
                        array1(l, a - 1) = array1(a - 1, l)
                        array1(0, a - 1) = arr(0, a) + 1
                        array1(a - 1, 0) = arr(0, a) + 1
                        rmatmod(l) = rmatx(j) - (arr(j, arr1(0)) + arr(j, arr1(1)) + arr(arr1(0), arr1(1))) / 2

                        rmatmod(a - 1) = rmatmod(a - 1) + array1(a - 1, l)

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

        For i = 1 To rmatmod.Length - 1
            rmatx(i) = rmatmod(i)
        Next

        Return array1
        ReDim array1(1, 1)
    End Function 'Reduces distance matrix in the each step of the NJ algorithm
    Function biodelcero(ByVal a As Integer, ByVal Vij(,) As Single, ByVal arr1() As Integer, ByVal lam As Single) As Single(,)



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


                    If i = a Then

                        If Vij IsNot Nothing Then
                            Dim v1i As Single = Vij(arr1(0), j)
                            Dim v2i As Single = Vij(arr1(1), j)
                            Dim v12 As Single = Vij(arr1(0), arr1(1))
                            array1(a - 1, l) = lam * v1i + (1 - lam) * v2i - lam * (1 - lam) * v12
                            array1(l, a - 1) = Vij(a - 1, l)
                        End If

                    End If
                    l = l + 1

                End If

            Next j

        Next i

        Return array1
    End Function 'Reduces the covariance matrix in BIONJ
    Function distonode(ByVal rmat1() As Single, ByVal mini() As Integer, ByVal notus As Integer, ByVal distancias(,) As Single) As Single
        Dim dist As Single
        dist = (distancias(mini(0), mini(1)) / 2) + ((rmat1(mini(0)) - rmat1(mini(1))) / (2 * (notus - 2)))
        Return dist
    End Function 'It calculates branch length
    Function Reducenotusmat(ByVal a As Integer, ByVal notusmatx() As String, ByVal min1() As Integer) As String()





        Dim l As Integer = 1

        Dim notusmat(a - 1) As String
        For i = 1 To a
            If i = a Then
                notusmat(a - 1) = notusmatx(min1(0)) & " " & notusmatx(min1(1))


            End If
            If i <> min1(0) And i <> min1(1) Then





                notusmat(l) = notusmatx(i)
                l = l + 1
            End If





        Next i


        Return notusmat
    End Function 'Makes splits in each step of the NJ/BIONJ algorithm
    '--------------------------------------------------------------------


    'Bootstrappig functions----------------------------------------
    Function bootstraping(ByVal bnotus(,) As String, ByVal seed As Double, ByRef otumat1(,) As String, ByVal nseq As Integer, ByVal bionj As Boolean) As Integer()
        Dim timer As New Stopwatch
        timer.Start()

        Dim mina1, mina2 As Integer
        Dim seqs(1) As String
        Dim bootstraprnd() As Integer
        Dim nseq1 As Integer = nseq + 1
        Dim nseqx2 As Integer = nseq * 2
        Dim distancias(nseq1, nseq1) As Single
        Dim i, j As Integer
        Dim Vij(nseq1, nseq1) As Single
        Dim dist As medirdistancia

        Dim sequ1, sequ2 As String
        j = 1
        dist = New medirdistancia
        Dim cadena As String
        Dim g As Integer = 0
        Dim k As Integer = 0

        'lseqred = REPS

        'Array.Resize(bootstraprnd, REPS)
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
                            If bionj = True Then
                                Vij(i, j) = distancias(i, j)
                                Vij(j, i) = Vij(i, j)
                            End If
                        End If
                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i






        Dim c As Integer = nseq1





        Dim nOTUS((nseqx2), 2) As String

        Dim nOtus1(nseqx2, 1) As Single
        Dim notus2(nseq - 1, 1) As String
        Dim matnodes(nseq + 1) As String
        Dim splits(nseq, 2) As String
        Dim h As Integer = 1
        Dim nseqr As Integer = c
        Dim rmatx(c) As Single
        Dim nmat(c, c) As Single
        Dim min1(1) As Integer
        Dim minrev(1) As Integer
        Dim NOTUSMATX(nseqr) As String
        For V = 1 To NOTUSMATX.Length - 1
            NOTUSMATX(V) = V
        Next



        Dim bootstrap(nseqx2) As Integer
        Do While c > 2
            If c = nseqr Then
                rmatx = rmat(c, distancias)
            End If
            nmat = Mmat(c, distancias, rmatx)




            min1 = min(c, nmat)
            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))
            minrev(0) = min1(1)
            minrev(1) = min1(0)
            Dim nsplitcl1, nsplitcl2 As String
            Dim diu, dju As Single
            diu = distonode(rmatx, min1, c, distancias)
            dju = distonode(rmatx, minrev, c, distancias)

            Dim lamb As Single = 0.5
            If bionj = True Then

                Dim suma As Single
                If Vij(min1(0), min1(1)) <> 0 Then
                    For w = 1 To c
                        If min1(0) <> w And min1(1) <> w And min1(0) <> min1(1) Then
                            suma = suma + (Vij(min1(0), w) - Vij(min1(1), w))
                        End If
                    Next
                    lamb = 0.5 + suma / (2 * (c - 2) * Vij(min1(0), min1(1)))
                End If
                If lamb > 1 Then lamb = 1
                If lamb < 0 Then lamb = 0
                Vij = biodelcero(c, Vij, min1, lamb)
            End If
            distancias = delcero(c, distancias, min1, lamb, diu, dju, rmatx)
            Dim posit As Integer = distancias(c - 1, 0)

            NOTUSMATX = Reducenotusmat(c, NOTUSMATX, min1)
            nOTUS(posit, 1) = " " & NOTUSMATX(NOTUSMATX.Length - 1) & " "
            nOTUS(posit, 2) = stdsplit(" " & NOTUSMATX(NOTUSMATX.Length - 1) & " ", nseqr)
            nOtus1(mina1, 0) = diu
            nOtus1(mina2, 0) = dju

            c = c - 1
        Loop

        nOtus1(distancias(2, 0), 0) = distancias(1, 2)
        nOtus1(distancias(1, 0), 0) = distancias(1, 2)

        ' Dim aa As Integer = nOtus1.GetLength(0) - 1
        'If nOTUS(aa, 0) = Nothing Then
        'For s = 1 To 2
        'If distancias(s, 0) <= nseq1 Then
        'nOTUS(aa, 0) = distancias(s, 0) & " "
        'nOTUS(aa, 1) = " " & distancias(s, 0) & " "
        'nOTUS(aa, 2) = stdsplit(nOTUS(aa, 1), nseq1)
        'End If
        'Next
        'End If
        'nOtus1(aa, 0) = distancias(1, 2) / 2





        'i = 1
        'j = 1

        For i = 1 To (nseqx2 - 1)
            For j = 1 To nseqx2 - 1

                If bnotus(i, 2) = nOTUS(j, 2) And nOtus1(j, 0) > 0 Then

                    bootstrap(i) = 1

                    Exit For
                End If
            Next

        Next

        Return bootstrap
    End Function 'bOOTSTRAPPING FOR SEQUENCES
    Function bootsperfal(ByVal bnotus(,) As String, ByVal seed As Double, ByVal perfiles(,) As String, ByVal nseq As Integer) As Integer()
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





        Dim nOTUS((nseqx2), 2) As String
        Dim nOtus1(nseqx2, 2) As Single
        Dim splits(nseq, 2) As String
        Dim h As Integer = 1
        Dim nseqr As Integer = c

        Do While c > 2


            Dim rmatx(c) As Single
            Dim nmat(c, c) As Single
            Dim min1(1) As Integer
            Dim minrev(1) As Integer
            rmatx = rmat(c, distancias)
            nmat = Mmat(c, distancias, rmatx)
            min1 = min(c, nmat)
            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))
            minrev(0) = min1(1)
            minrev(1) = min1(0)
            Dim nsplitcl1, nsplitcl2 As String


            If mina1 < nseq + 2 Then


                nOTUS(h, 1) = " " & mina1 & " "
                nOTUS(h, 0) = mina1 & " "
                nOTUS(h, 2) = stdsplit(nOTUS(h, 1), nseq1)
                nsplitcl1 = mina1
                nOtus1(h, 0) = distonode(rmatx, min1, c, distancias)

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
                nOtus1(pos, 0) = distonode(rmatx, min1, c, distancias)
            End If


            If mina2 < nseq + 2 Then


                nOTUS(h, 1) = " " & mina2 & " "
                nOTUS(h, 0) = mina2 & " "
                nOTUS(h, 2) = stdsplit(nOTUS(h, 1), nseq1)
                nsplitcl2 = mina2
                nOtus1(h, 0) = distonode(rmatx, minrev, c, distancias)
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
                nOtus1(pos, 0) = distonode(rmatx, minrev, c, distancias)
            End If

            nOTUS(h, 1) = " " & nsplitcl1 & " " & nsplitcl2 & " "
            nOTUS(h, 0) = nseq + 1 + (nseq + 2 - c) & " "
            nOTUS(h, 2) = stdsplit(nOTUS.GetValue(h, 1), nseq + 1)
            h = h + 1

            Dim diju As Single = distancias(min1(0), min1(1)) / 2
            distancias = delcero(c, distancias, min1, 1 / 2, diju, diju, rmatx)
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
    End Function 'bOOTSTRAPPING FOR ALLELIC PROFILES

  
    'Modified NJ algorithm to force a branch to exist or not
    Function NJproc1(ByVal nseq1 As Integer, ByVal distancias(,) As Single, ByVal Vij(,) As Single, ByVal bionj As Boolean, ByVal split As String, ByVal force As Boolean) As splits
        Dim mina1, mina2 As Integer
        Dim i, j As Integer
        Dim c As Integer = nseq1
        Dim newa As New newicker


        distancias1 = distancias
        Dim sp1 As splits

        Dim nseqx2 As Integer = (nseq1 - 1) * 2
        ReDim sp1.nOTUs((nseqx2), 3)
        ReDim sp1.notus1(nseqx2, 1)


        Dim h As Integer = 1


        Dim nseqr As Integer
        nseqr = c

        Dim min1(1) As Integer
        Dim NOTUSMATX(nseqr) As String
        For V = 1 To NOTUSMATX.Length - 1
            NOTUSMATX(V) = V

        Next
        Dim rmatx(c) As Single
        Do While c > 2
            Dim nmat(c, c) As Single
            Dim minrev(1) As Integer
            Dim diu, dju As Single

            If c = nseqr Then
                rmatx = rmat(c, distancias)
            End If
            nmat = Mmat(c, distancias, rmatx)

fmin:
            min1 = min(c, nmat)
            mina1 = distancias(0, min1(0))
            mina2 = distancias(0, min1(1))
            minrev(0) = min1(1)
            minrev(1) = min1(0)

            nmat(min1(0), min1(1)) = 1000000
            nmat(min1(1), min1(0)) = 1000000


            diu = distonode(rmatx, min1, c, distancias)
            dju = distonode(rmatx, minrev, c, distancias)



            Dim lamb As Single = 0.5

            If bionj = True Then

                Dim suma As Single
                If Vij(min1(0), min1(1)) <> 0 Then
                    For w = 1 To c
                        If min1(0) <> w And min1(1) <> w And min1(0) <> min1(1) Then
                            suma = suma + (Vij(min1(0), w) - Vij(min1(1), w))
                        End If
                    Next
                    lamb = 0.5 + suma / (2 * (c - 2) * Vij(min1(0), min1(1)))
                End If
                If lamb > 1 Then lamb = 1
                If lamb < 0 Then lamb = 0
                Vij = biodelcero(c, Vij, min1, lamb)
            End If
            Dim checkedsplit As String = stdsplit(" " & NOTUSMATX(min1(0)) & " " & NOTUSMATX(min1(1)) & " ", nseqr)
            If force = True Then
                checkedsplit = newa.rewritesplits(" " & checkedsplit, nseq1)

                If Form1.checkcomp(checkedsplit, split, nseq1) = False Then

                    GoTo fmin

                End If
            Else


                If checkedsplit = split Then

                    GoTo fmin
                End If
            End If


            NOTUSMATX = Reducenotusmat(c, NOTUSMATX, min1)

            distancias = delcero(c, distancias, min1, lamb, diu, dju, rmatx)
            Dim posit As Integer = distancias(c - 1, 0)
            sp1.nOTUs(posit, 1) = " " & NOTUSMATX(NOTUSMATX.Length - 1) & " "
            sp1.nOTUs(posit, 2) = stdsplit(" " & NOTUSMATX(NOTUSMATX.Length - 1) & " ", nseqr)
            sp1.notus1(mina1, 0) = Math.Abs(diu)
            sp1.notus1(mina2, 0) = Math.Abs(dju)

            c = c - 1




        Loop



        Dim aa As Integer = sp1.notus1.GetLength(0) - 1
        If sp1.nOTUs(aa, 1) = Nothing Then
            For s = 1 To 2
                If distancias(s, 0) <= nseq1 Then
                    sp1.nOTUs(aa, 0) = distancias(s, 0) & " "
                    sp1.nOTUs(aa, 1) = " " & distancias(s, 0) & " "
                    sp1.nOTUs(aa, 2) = stdsplit(sp1.nOTUs(aa, 1), nseq1)
                End If
            Next
        End If
        sp1.notus1(distancias(2, 0), 0) = distancias(1, 2) / 2
        sp1.notus1(distancias(1, 0), 0) = distancias(1, 2) / 2

        Return sp1

    End Function 'algorithm
    Function minalt(ByVal arr(,) As Single, ByVal min1() As Integer) As Integer()

        Dim minarr(1) As Integer
        Dim m As Single = 10000000000
        For i = 1 To arr.GetLength(0) - 1
            If i <> min1(0) And i <> min1(1) Then

                If arr(i, min1(0)) < m Then
                    m = arr(i, min1(0))
                    minarr(0) = min1(0)
                    minarr(1) = i
                End If
                If arr(i, min1(1)) < m Then
                    m = arr(i, min1(1))
                    minarr(0) = i
                    minarr(1) = min1(1)
                End If
            End If

        Next i
        Return minarr
    End Function
    Function checkxtrasteps(ByVal bnotus(,) As String, ByVal seed As Double, ByVal nseq As Integer, ByVal bionj As Boolean, ByVal extralengths() As Single, ByVal treelengths() As Single, ByVal splits(,) As String, ByVal testsel As Boolean) As Single()
        Dim anotus(bnotus.GetLength(0) - 1) As Single
        Dim splitests(,) As String = splitestmaker()
        Dim selsplits As Boolean

        For bn = nseq + 2 To bnotus.GetLength(0) - 2
            Dim test As Boolean = False
            If extralengths Is Nothing Then
                test = True
            Else
                If extralengths(bn) > 0 Then
                    test = True
                End If
            End If
            If testsel = True And splits(1, 1) <> Nothing And test = True Then
                test = False
                For i = 1 To splits.GetLength(0) - 1
                    If bnotus(bn, 2) = splits(i, 1) Then
                        test = True
                        Exit For
                    End If
                Next


            End If

            If test = True Then


                '''

                Dim seqs(1) As String
                Dim bootstraprnd() As Integer
                Dim nseq1 As Integer = nseq + 1
                Dim nseqx2 As Integer = nseq * 2

                Dim i, j As Integer

                Dim dist As medirdistancia

                Dim sequ1, sequ2 As String
                j = 1
                dist = New medirdistancia
                Dim cadena As String
                Dim g As Integer
                Dim k As Integer = 0


                For s = 0 To otum.Length - 1

                    Dim distancias(nseq1, nseq1) As Single
                    Dim vij(nseq1, nseq1) As Single
                    'Dim a As Single = Rnd()

                    distancias = otum(s).dista

                    Dim c As Integer = nseq1



                    Dim sp As splits
                    If otum(s).otumat1(1, 1) = Nothing Then
                        GoTo 22
                    End If
                    sp = NJproc1(distancias.GetLength(0) - 1, otum(s).dista, vij, False, bnotus(bn, 0), True)



                    Dim asf As Single = 0
                    For v = 1 To sp.notus1.GetLength(0) - 1
                        asf = asf + sp.notus1(v, 0)

                    Next

                    asf = asf - otum(s).treelength


                    If asf < 0 Then ' otum(s).treelength Then
                        'anotus(bn) = anotus(bn) + otum(s).treelength
                        'Stop
                    Else


                        anotus(bn) = anotus(bn) + asf

                    End If
                    ''''
                    'Exit For
                    ''''
22:             Next

            End If

        Next

        Return anotus
    End Function







    'Clade significance-----------------------------

    Function cladesupport(ByVal split As String, ByVal spx As splits, ByVal dista As dsx, ByVal alignm(,) As String) As Single
        Dim spalt As splits
        Dim newa As New newicker
        Dim N As Integer = dista.distancias.GetLength(0) - 1

        spalt = NJproc1(N, dista.distancias, Nothing, False, split, False)

        For i = N + 1 To spx.nOTUs.GetLength(0) - 1
            spx.nOTUs(i, 0) = newa.rewritesplits(" " & spx.nOTUs(i, 2), alignm.GetLength(0) - 1)
            spalt.nOTUs(i, 0) = newa.rewritesplits(" " & spalt.nOTUs(i, 2), alignm.GetLength(0) - 1)
        Next
        For i = 1 To N
            spalt.nOTUs(i, 1) = spx.nOTUs(i, 1)
            spalt.nOTUs(i, 2) = spx.nOTUs(i, 2)
            spalt.notus1(i, 0) = spx.notus1(i, 0)
        Next
        spalt.otumat1 = spx.otumat1
        'Dim nwkform As New treeviewer2 With {.sptree = spx, .Text = "Tree Viewerx"} '
        'nwkform.Show()
        'Dim nwkform1 As New treeviewer2 With {.sptree = spalt, .Text = "Tree Vieweralt"} '
        'nwkform1.Show()

        Dim matBx(N, N), matBalt(N, N) As Integer
        Dim maxbx, maxbalt As Integer
        maxbx = 0
        maxbalt = 0
        For i = 1 To N

            For j = i + 1 To N

                matBx(i, j) = calcB(spx.nOTUs, i, j)
                If maxbx < matBx(i, j) Then maxbx = matBx(i, j)



                matBalt(i, j) = calcB(spalt.nOTUs, i, j)
                If maxbalt < matBalt(i, j) Then maxbalt = matBalt(i, j)


            Next

        Next
        Dim x(1) As Single
        Dim x1(alignm(1, 1).Length - 1) As Single
        Dim count, count1 As Integer
        For i = 0 To alignm(1, 1).Length - 1
            Dim L As Double = 0
            Dim L1 As Double = 0
            For j = 1 To N
                For k = j + 1 To N
                    Dim dist As New medirdistancia

                    Dim distaz As Double
                    If alignm(j, 1).Chars(i) <> alignm(k, 1).Chars(i) Then
                        distaz = dist.pesos(alignm(j, 1).Chars(i), alignm(k, 1).Chars(i))
                    Else : distaz = 0
                    End If

                    Dim topo1 As Integer = 2 ^ (maxbx - matBx(j, k)) '(matBx(j, k) - 1)
                    Dim topo2 As Integer = 2 ^ (maxbalt - matBalt(j, k)) '(matBalt(j, k) - 1)
                    '''

                    'L = L + dista.distancias(j, k) / topo1
                    'L1 = L1 + dista.distancias(j, k) / topo2
                    '''
                    L = L + distaz * topo1
                    L1 = L1 + distaz * topo2


                Next

            Next
            L = L / (2 ^ (maxbx - 1))
            L1 = L1 / (2 ^ (maxbalt - 1))
            If L1 - L > 0 Then
                count = count + 1

            End If

            If Math.Abs(L - L1) > 0 Then
                count1 = count1 + 1
                Array.Resize(x, count1)
                x(count1 - 1) = L1 - L
            End If


        Next


        Dim RFD, c, f As Single

        For i = 0 To x.Length - 1
            If x(i) > 0 Then
                f = f + x(i)
            Else
                c = c + Math.Abs(x(i))
            End If
        Next

        Dim a As Single = 1 - wilcoxon(x)
        'If c < 0 Then Stop

        RFD = (f - c) / f

        Return a 'RFD 'count / count1 ' 

    End Function
    Function cladesupport1(ByVal spx As splits, ByVal spalt As splits, ByVal alignm(,) As String) As Integer()


        Dim newa As New newicker
        Dim N As Integer = alignm.GetLength(0) - 1

        For i = N + 1 To spx.nOTUs.GetLength(0) - 1
            spx.nOTUs(i, 0) = newa.rewritesplits(spx.nOTUs(i, 1), alignm.GetLength(0) - 1)
            spalt.nOTUs(i, 0) = newa.rewritesplits(spalt.nOTUs(i, 1), alignm.GetLength(0) - 1)
        Next


        Dim matBx(N, N), matBalt(N, N) As Integer
        For i = 1 To N

            For j = i + 1 To N

                matBx(i, j) = calcB(spx.nOTUs, i, j)



                matBalt(i, j) = calcB(spalt.nOTUs, i, j)



            Next

        Next
        Dim counts(1) As Integer
        Dim count, count1 As Integer
        For i = 0 To alignm(1, 1).Length - 1
            Dim L As Double = 0
            Dim L1 As Double = 0
            For j = 1 To N
                For k = j + 1 To N
                    Dim dist As New medirdistancia

                    Dim distaz As Single
                    If alignm(j, 1).Chars(i) <> alignm(k, 1).Chars(i) Then
                        distaz = dist.pesos(alignm(j, 1).Chars(i), alignm(k, 1).Chars(i))
                    Else : distaz = 0
                    End If
                    Dim topo1 = 2 ^ (matBx(j, k) - 1)
                    Dim topo2 = 2 ^ (matBalt(j, k) - 1)

                    L = L + distaz / topo1 '((c - 2) * 2 ^ (calcB(notus, splitA, sa(j - 1), k) - 2))
                    L1 = L1 + distaz / topo2 '((c - 2) * 2 ^ (calcB(notus, splitB, sa(j - 1), k) - 2))


                Next

            Next

            If L1 - L > 10 ^ -9 Then
                count = count + 1

            End If

            If Math.Abs(L - L1) > 10 ^ -6 Then
                count1 = count1 + 1

            End If


        Next
        counts(0) = count
        counts(1) = count1

        Return counts
    End Function
    Function wilcoxon(ByVal x() As Single) As Single
        Dim N As Integer = x.Length
        Dim x1(x.Length - 1, 3) As Single
        Dim min As Single = 1000
        Dim idmin As Integer
        For j = 0 To N - 1
            For i = 0 To N - 1
                If Math.Abs(x(i)) <= min Then
                    idmin = i
                    min = Math.Abs(x(i))
                End If
            Next

            x1(j, 0) = x(idmin)
            x1(j, 1) = Math.Abs(x(idmin))
            x1(j, 3) = j + 1
            x1(j, 2) = Math.Sign(x(idmin))
            x(idmin) = 1000
            min = 1000
        Next
        For i = 0 To N - 2
            If x1(i, 1) = x1(i + 1, 1) Then
                Dim ni As Integer = 2
                Dim r As Single = 2 * (i + 1) + 1
                For j = i + 2 To N - 1
                    If x1(j, 1) = x1(i, 1) Then
                        ni = ni + 1
                        r = r + j + 1
                    Else : Exit For
                    End If
                Next
                r = r / ni
                For f = i To i + ni - 1
                    x1(f, 3) = r
                Next
                i = i + ni - 1
            End If






        Next
        Dim W As Single
        For i = 0 To N - 1
            If x1(i, 2) = 1 Then
                W = W + x1(i, 3)
            End If
        Next
        Dim pval As Single
        If N <= 20 Then
            Dim Wt As Integer = Math.Round(W)
            Dim Wc() As Integer = wilcoD(N)
            Dim t As Integer = 0
            For I = Wt To Wc.Length - 1
                t = t + Wc(I)

            Next
            pval = t / (2 ^ N)
        Else

            Dim z As Double
            Dim a As Double
            Dim ov As Double
            ov = N

            z = (W - 0.5 - (ov * (ov + 1) / 4)) / Math.Sqrt(ov * (ov + 1) * (2 * ov + 1) / 24)
            pval = 1 - Form2.ztest(z)
            If pval < 0 Then pval = 0

        End If
        Return pval
    End Function
    Function wilcoD(ByVal N As Integer) As Integer()
        Dim arr(N, N * (N + 1) / 2) As Integer
        arr(1, 0) = 1
        arr(1, 1) = 1
        For i = 2 To N
            For j = 0 To i * (i + 1) / 2
                If j < i Then
                    arr(i, j) = arr(i - 1, j)
                ElseIf j >= i And j <= i * (i - 1) / 2 Then
                    arr(i, j) = arr(i - 1, j) + arr(i - 1, j - i)
                Else
                    arr(i, j) = arr(i - 1, j - i)
                End If


            Next
        Next
        Dim W(N * (N + 1) / 2) As Integer
        For i = 0 To N * (N + 1) / 2
            W(i) = arr(N, i)
        Next
        Return W
    End Function
    Function lengthtree(ByVal notus(,) As String, ByVal dista(,) As Single) As Single
        Dim B As Integer
        Dim L As Single
        For i = 1 To dista.GetLength(0) - 1
            For j = i + 1 To dista.GetLength(0) - 1
                'B = calcB(notus, i, j)
                L = L + (dista(i, j) * (2 ^ (-B)))

            Next
        Next
        Return L
    End Function
    Function calcB(ByVal notus(,) As String, ByVal i As Integer, ByVal j As Integer) As Integer
        Dim count As Integer
        For z = 1 To notus.GetLength(0) - 2
            If notus(z, 0) <> Nothing Then

                If notus(z, 0).Chars(i - 1) = "*"c Then
                    If notus(z, 0).Chars(j - 1) = "."c Then
                        count = count + 1
                    End If
                Else

                    If notus(z, 0).Chars(j - 1) = "*"c Then
                        count = count + 1
                    End If
                End If
            Else
                If z = i Or z = j Then
                    count = count + 1
                End If
            End If

        Next


        Return count
    End Function 'it calculates the number of branches between two taxa
    '-----------------------------------------------


    'Other-------------------------------------------
    Function localS(ByVal split1 As String, ByVal split2 As String, ByVal splitc As String, ByVal notus(,) As String, ByVal alignm(,) As String, ByVal c As Integer, ByVal topo() As Integer) As Single
        Dim sitenmat1(alignm(1, 1).Length - 1) As Double
        Dim sitenmat2(alignm(1, 1).Length - 1) As Single
        Dim N As Integer = alignm.GetLength(0) - 1
        Dim s1size As Integer = treeviewer2.splitsize(" " & split1 & " ")
        Dim s2size As Integer = treeviewer2.splitsize(" " & split2 & " ")
        Dim splitA As String = " " & split1 & " " & splitc & " "
        Dim splitB As String = " " & split2 & " " & splitc & " "
        Dim s As String = split1 & " " & split2
        split1 = " " & split1 & " "
        split2 = " " & split2 & " "
        splitc = " " & splitc & " "
        Dim sa() As String = s.Split(" "c)
        Dim count, count1 As Integer
        For i = 0 To sitenmat1.Length - 1
            Dim L As Double = 0
            Dim L1 As Double = 0
            For j = 1 To s1size + s2size
                Dim dist As New medirdistancia
                Dim dista As Single
                For k = 1 To N
                    If split1.Contains(" " & k & " ") And split1.Contains(" " & sa(j - 1) & " ") Or (split2.Contains(" " & k & " ") And split2.Contains(" " & sa(j - 1) & " ")) Then
                        'Stop
                    Else
                        If alignm(sa(j - 1), 1).Substring(i, 1) <> alignm(k, 1).Substring(i, 1) Then
                            dista = dist.pesos(alignm(sa(j - 1), 1).Substring(i, 1), alignm(k, 1).Substring(i, 1))
                        Else : dista = 0
                        End If
                        If sa.Contains(k) Then
                            dista = dista / 2
                        End If
                        If dista > 0 Then
                            Dim topo1, topo2 As Integer
                            Dim b As Integer = topo(sa(j - 1)) + topo(k) + 1

                            If (" " & splitc & " ").Contains(" " & k & " ") Then
                                If split1.Contains(" " & sa(j - 1) & " ") Then
                                    topo1 = 2 ^ b
                                    topo2 = (c - 2) * 2 ^ b
                                Else
                                    topo1 = (c - 2) * 2 ^ b
                                    topo2 = 2 ^ b
                                End If
                            Else
                                If split1.Contains(" " & sa(j - 1) & " ") Then
                                    topo1 = (c - 2) * 2 ^ b
                                    topo2 = (c - 2) * 2 ^ (b - 1)
                                Else
                                    topo1 = (c - 2) * 2 ^ (b - 1)
                                    topo2 = (c - 2) * 2 ^ b
                                End If

                            End If


                            L = L + dista / topo1 '((c - 2) * 2 ^ (calcB(notus, splitA, sa(j - 1), k) - 2))
                            L1 = L1 + dista / topo2 '((c - 2) * 2 ^ (calcB(notus, splitB, sa(j - 1), k) - 2))
                        End If
                    End If
                    Application.DoEvents()
                Next

            Next

            If L1 - L > 10 ^ -9 Then
                count = count + 1

            End If

            If Math.Abs(L - L1) > 10 ^ -6 Then
                count1 = count1 + 1

            End If


            'sitenmat2(i) = 
        Next

        Return 1 - combinatF1(0.5, count1, count)
    End Function 'not used
    Function combinatF1(ByVal x As Double, ByVal N As Integer, ByVal m As Integer) As Double
        Dim p As Double = 0
        If N > 0 Then

            If N <= 100 Then
                For i = m To N
                    p = p + ((factorial(N) * (x ^ i) * ((1 - x) ^ (N - i))) / (factorial(i) * factorial(N - i)))

                Next
            Else
                Dim DESVIO As Single = Math.Sqrt(m * (N - m) / N)
                Dim z As Single
                z = ((N / 2) - 0.5 - m) / DESVIO
                If z >= 0 Then
                    p = Form2.ztest(z)
                Else
                    p = 1 - Form2.ztest(Math.Abs(z))
                End If
            End If
        Else
            p = 1
        End If

        Return p
    End Function 'Binomial distribution
    Function combinatF2(ByVal x As Double, ByVal N As Integer, ByVal m As Integer) As Double
        Dim p As Double = 0

        For i = m To m
            p = ((producto(N - m + 1, N) * (x ^ i) * ((1 - x) ^ (N - i))) / (factorial(i)))

        Next


        Return p
    End Function 'Binomial distribution
    Function producto(ByVal n1 As Integer, ByVal n2 As Integer) As Double
        producto = 1
        For a = n1 To n2
            producto = producto * a
        Next

    End Function
    Function factorial(ByVal val) As Double
        Dim a As Double = val
        For i = 1 To a - 1
            a = a * i
        Next
        If val = 0 Then
            a = 1
        End If
        Return a
    End Function
    Function letter(ByVal x As Integer) As Char


        Select Case x
            Case 0
                letter = "A"c
            Case 1
                letter = "T"c
            Case 2
                letter = "C"c
            Case 3
                letter = "G"c
        End Select


    End Function
    Function count(ByVal a As String) As Integer
        Dim i As Integer = 0
        Dim n As Integer = 1
        Do While i + 1 < a.Length - 1

            If a.Substring(i + 1, 1) = " " Then
                n = n + 1
                i = i + 1
            Else

                i = i + 1
            End If
        Loop
        Return n
    End Function
    Function comparedst(ByVal a As String, ByVal distancias(,) As Single) As Integer
        Dim size As Integer = count(a) - 1
        Dim dst(size) As Integer
        Dim n, j As Integer
        j = 0
        Dim b As String = Nothing
        Dim start, fin As Integer
        Dim startbo As Boolean
        Dim indexa As Integer = 0
        For j = 0 To a.Length - 1
            If a.Substring(j, 1) = " " Then
                If startbo = False Then
                    startbo = True
                    start = j

                Else
                    startbo = False
                    b = b.Substring(0, b.Length - 1)
                    dst(indexa) = b
                    If j < a.Length - 1 Then
                        j = j - 1
                    End If
                    indexa = indexa + 1
                    b = Nothing
                End If

            End If
            If startbo = True Then
                b = b & a.Substring(j + 1, 1)
            End If
        Next

        n = size + 1
        For i = 0 To size - 1
            For j = i + 1 To size
                If distancias(dst(i), dst(j)) = 0 Then
                    n = n - 1
                    j = size


                End If
            Next
        Next
        Return n
    End Function
    Function reductor(ByRef otu1 As String(,)) As String(,)
        Dim i As Long = 0
        Dim otu2(otu1.GetLength(0) - 1) As System.Text.StringBuilder
        For o = 1 To otu2.GetLength(0) - 1
            otu2(o) = New System.Text.StringBuilder()
        Next
        Dim consta As Boolean = True
        Dim seq1, seq2 As String
        Dim nseq1 As Integer = otu1.GetLength(0) - 1
        Dim lseq1 As Integer = otu1(1, 1).Length

        Do While i < lseq1
            consta = True

            For j = 1 To nseq1 - 1


                If otu1(j, 1)(i) <> otu1(j + 1, 1)(i) Then
                    consta = False
                    Exit For
                End If

            Next
            If consta = False Then

                For t = 1 To nseq1

                    'otu1(t, 1) = otu1(t, 1).Remove(i, 1) 'otu1(t, 1).ToString.Substring(0, i) & otu1(t, 1).ToString.Substring(i + 1, otu1(t, 1).ToString.Length - (i + 1))

                    otu2(t) = otu2(t).Append(otu1(t, 1)(i))

                Next
                'i = i - 1
                'lseq1 = lseq1 - 1




            End If
            i = i + 1
        Loop
        For j = 1 To nseq1


            otu1(j, 1) = otu2(j).ToString()


        Next
        Return otu1
    End Function 'Delete constant sites in a alignment
    Function splitestmaker() As String(,)
        Dim splitest(Form1.DataGridView6.RowCount + 1, 1) As String
        For i = 1 To Form1.DataGridView6.RowCount
            splitest(i, 1) = Form1.DataGridView6.Item(2, i - 1).Value
            splitest(i, 0) = Form1.DataGridView6.Item(3, i - 1).Value

        Next
        Return splitest
    End Function 'makes splits to test
    Function splitestmakerEX() As String(,)
        Dim splitest(Form1.DataGridView8.RowCount + 1, 1) As String
        For i = 1 To Form1.DataGridView8.RowCount
            splitest(i, 1) = Form1.DataGridView8.Item(2, i - 1).Value
            splitest(i, 0) = Form1.DataGridView8.Item(3, i - 1).Value

        Next
        Return splitest
    End Function
    Function dstcalc(ByVal distancias(,) As Single) As Integer



        Dim Dstarr As Integer
        Dim splitest As String
        splitest = " "
        For i = 1 To distancias.GetLength(0) - 1

            splitest = splitest & i & " "


        Next


        Dstarr = comparedst(splitest, distancias)



        Return Dstarr

    End Function
    Function DSTs(ByVal splitest(,) As String, ByVal distancias(,) As Single) As Integer()
        Dim Dstarr(splitest.GetLength(0) - 1) As Integer
        splitest(splitest.GetLength(0) - 1, 0) = " "

        For i = 1 To distancias.GetLength(0) - 1

            splitest(splitest.GetLength(0) - 1, 0) = splitest(splitest.GetLength(0) - 1, 0) & i & " "


        Next
        For i = 0 To splitest.GetLength(0) - 1
            If splitest(i, 0) <> "" Then
                Dstarr(i) = comparedst(splitest(i, 0), distancias)
            End If
        Next

        Return Dstarr

    End Function 'calculate DSTs
    Function stdsplit(ByVal split As String, ByVal ntax As Integer) As String
        Dim cerosplit As String
        Dim i As Integer
        Dim g As Integer = 0
        For i = 1 To ntax
            If split.Contains(" " & 1 & " ") = True Then
                If split.Contains(" " & i & " ") = True Then

                    cerosplit = cerosplit & i & " "
                End If
            Else
                If split.Contains(" " & i & " ") = False Then

                    cerosplit = cerosplit & i & " "
                End If
            End If
        Next i
        cerosplit = cerosplit & ","


        Return cerosplit
    End Function 'makes a split in a standard format
    Public Function nseqx(ByVal file As String) As Integer

        Dim lector As TextReader
        lector = New StreamReader(file)
        Dim i As Integer
        Dim lin As String = "1"
        Dim A As Integer = 0
        Dim len As Integer
        Dim f As Boolean = False
        Do While A < 2
            lin = lector.ReadLine()
            'If lin = Nothing Then Stop
            If lin <> Nothing Then
                If lin.Substring(0, 1) = ">" Then
                    i = i + 1
                    A = 0
                Else
                    If f = False Then
                        f = True
                        len = lin.Length
                    Else
                        If lin.Length <> len Then
                            'Return -1
                            'Exit Function
                        End If
                    End If
                End If
            Else
                A = A + 1
            End If
        Loop
        lector.Close()
        Dim nseq As Integer
        nseq = i - 1
        Return nseq
    End Function 'retrieves the number of sequences into an alignment
    Function contiene(ByVal array(,) As String, ByVal i As Integer, ByVal value As String) As Boolean
        For j = 0 To array.GetLength(0) - 1
            If array(j, i) = value Then
                Return True
                Exit Function
            End If
        Next
    End Function



End Module
