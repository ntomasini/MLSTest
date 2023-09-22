Imports System.Math
Public Structure EIGEN
    Dim MatEiVa As Double(,)
    Dim MatEiVe As Double(,)
End Structure
Module MDS1
    Function MDS1(ByVal distas As Double(,)) As Double(,)
        Dim b, b1 As Double(,)
        Dim P2(,) As Double = cuadradomatriz(distas)
        Dim J As Double(,) = makeJ(distas.GetLength(0) - 1)
        b1 = multiplymatrix(medioJneg(J), P2)
        b = multiplymatrix(b1, J)
        Dim eigen1 As EIGEN
        eigen1 = MatEigenvalue_QR(b)
        Dim matEival(,) As Double = makeraizEvalmat(eigen1.MatEiVa)
        Dim sum As Double = sumabs(eigen1.MatEiVa)
        Dim explain(1) As Single
        explain(0) = eigen1.MatEiVa(eigen1.MatEiVa.GetLength(0) - 1, 1) / sum
        explain(1) = eigen1.MatEiVa(eigen1.MatEiVa.GetLength(0) - 2, 1) / sum
        Dim vector(,) As Double
        vector = (menosvector(makevector(eigen1.MatEiVe)))
        Dim result(,) As Double = multiplymatrix(vector, matEival)
        result(0, 0) = explain(0)
        result(0, 1) = explain(1)
        Return result

    End Function
    Function sumabs(ByVal mat1 As Double(,)) As Double
        For i = 0 To mat1.GetLength(0) - 1
            sumabs = sumabs + Math.Abs(mat1(i, 1))

        Next
    End Function
    Private Function makevector(ByVal mateivect As Double(,)) As Double(,)
        Dim arr(mateivect.GetLength(0) - 1, 2) As Double

        For i = 1 To arr.GetLength(0) - 1
            arr(i, 1) = mateivect(i, mateivect.GetLength(1) - 1)
            arr(i, 2) = mateivect(i, mateivect.GetLength(1) - 2)
        Next
        Return arr
    End Function
    Private Function menosvector(ByVal mat1 As Double(,)) As Double(,)
        Dim arr(mat1.GetLength(0) - 1, mat1.GetLength(1) - 1) As Double
        For i = 1 To mat1.GetLength(0) - 1
            For j = 1 To mat1.GetLength(1) - 1
                arr(i, j) = -mat1(i, j)
            Next
        Next
        Return arr
    End Function
    Private Function makeraizEvalmat(ByVal mat1 As Double(,)) As Double(,)
        Dim matrix(2, 2) As Double
        matrix(1, 1) = Sqrt(mat1(mat1.GetLength(0) - 1, 1))
        matrix(1, 0) = 0
        matrix(2, 2) = Sqrt(mat1(mat1.GetLength(0) - 2, 1))
        matrix(2, 1) = 0
        Return matrix

    End Function
    Private Function cuadradomatriz(ByVal dista As Double(,)) As Double(,)
        For i = 1 To dista.GetLength(0) - 1
            For j = 1 To dista.GetLength(0) - 1
                dista(i, j) = dista(i, j) * dista(i, j)

            Next
        Next
        Return dista
    End Function
    Private Function makeJ(ByVal nseqs As Integer) As Double(,)
        Dim arr(nseqs, nseqs) As Double
        For i = 1 To nseqs
            For j = 1 To nseqs
                If i = j Then
                    arr(i, j) = 1 - (1 / nseqs)
                Else
                    arr(i, j) = -1 / nseqs
                End If

            Next
        Next
        Return arr

    End Function
    Private Function medioJneg(ByVal arra As Double(,)) As Double(,)
        Dim arr1(arra.GetLength(0) - 1, arra.GetLength(0) - 1) As Double
        For i = 1 To arr1.GetLength(0) - 1
            For j = 1 To arr1.GetLength(0) - 1
                arr1(i, j) = -arra(i, j) / 2

            Next
        Next
        Return arr1
    End Function
    Private Function multiplymatrix(ByVal mat1 As Double(,), ByVal mat2 As Double(,))
        Dim result(mat1.GetLength(0) - 1, mat2.GetLength(1) - 1) As Double

        For i = 1 To result.GetLength(0) - 1
            For j = 1 To result.GetLength(1) - 1
                For r = 1 To mat1.GetLength(1) - 1
                    result(i, j) = result(i, j) + (mat1(i, r) * mat2(r, j))
                Next
            Next
        Next
        Return result
    End Function
    Private Function MatEigenvalue_QR(ByVal Mat(,) As Double) As EIGEN
        'Find real and complex eigenvalues with the iterative QR method
        Dim wr#(), wi#(), tiny#, nm#
        Dim A(Mat.GetLength(0) - 1, Mat.GetLength(1) - 1) As Double
        Dim b(,) As Double
        tiny = 2 * 10 ^ -15
        Array.Copy(Mat, A, Mat.Length)

        Dim N = UBound(A, 1)

        ReDim wr(N), wi(N)

        ELMHES0(N, A)

        HQR(nm, N, 1, N, A, wr, wi, 0)

        ReDim b(N, 2)

        For i = 1 To N
            If i > 0 Then
                b(i, 1) = wr(i)
                b(i, 2) = wi(i)
            Else
                b(i, 1) = "?"
                b(i, 2) = "?"
            End If
        Next
        b = MatMopUp(b, tiny)
        b = MatrixSort(b, "A")
        Dim c(1, 2) As Double
        c(0, 1) = b(b.GetLength(0) - 2, 1)
        c(1, 1) = b(b.GetLength(0) - 1, 1)

        Dim asa(,) As Double
        asa = MatEigenvector(Mat, b, 0)
        Dim eigen1 As EIGEN
        eigen1.MatEiVa = b
        eigen1.MatEiVe = asa
        Return eigen1
    End Function
    Private Sub ELMHES0(ByVal N, ByVal Mat)
        '  sources by Martin, R. S. and Wilkinson, J. H., see [MART70].   *
        '
        Dim k As Long, i As Long, j As Long, x As Double, y As Double

        For k = 2 To N - 1
            i = k
            x = 0
            For j = k To N
                If (Abs(Mat(j, k - 1)) > Abs(x)) Then
                    x = Mat(j, k - 1)
                    i = j
                End If
            Next j
            If (i <> k) Then
                '           SWAP0 rows and columns of MAT
                For j = k - 1 To N
                    Dim temp As Double = Mat(i, j)
                    Mat(i, j) = Mat(k, j)
                    Mat(k, j) = temp
                Next j
                For j = 1 To N
                    Dim temp As Double = Mat(j, i)
                    Mat(j, i) = Mat(j, k)

                    Mat(j, k) = temp
                Next j
            End If
            If (x <> 0) Then
                For i = k + 1 To N
                    y = Mat(i, k - 1)
                    If (y <> 0) Then
                        y = y / x
                        Mat(i, k - 1) = y
                        For j = k To N
                            Mat(i, j) = Mat(i, j) - (y * Mat(k, j))
                        Next j
                        For j = 1 To N
                            Mat(j, k) = Mat(j, k) + (y * Mat(j, i))
                        Next j
                    End If
                Next i
            End If
        Next k
    End Sub
    Private Sub SWAP0(ByVal x, ByVal y)
        Dim temp As Double
        temp = x
        x = y
        y = temp
    End Sub

    Private Sub HQR(ByVal nm, ByVal N, ByVal low, ByVal igh, ByVal h, ByVal wr, ByVal wi, ByVal Ierr)
        '
        '     THIS SUBROUTINE IS A TRANSLATION OF THE ALGOL PROCEDURE HQR,
        '     NUM. MATH. 14, 219-231(1970) BY MARTIN, PETERS, AND WILKINSON.
        '     HANDBOOK FOR AUTO. COMP., VOL.II-LINEAR ALGEBRA, 359-371(1971).
        '
        '     THIS SUBROUTINE FINDS THE EIGENVALUES OF A REAL
        '     UPPER HESSENBERG MATRIX BY THE QR METHOD.
        '
        '     ON INPUT
        '
        '        NM MUST BE SET TO THE ROW DIMENSION OF TWO-DIMENSIONAL
        '          ARRAY PARAMETERS AS DECLARED IN THE CALLING PROGRAM
        '          DIMENSION STATEMENT.
        '
        '        N IS THE ORDER OF THE MATRIX.
        '
        '        LOW AND IGH ARE INTEGERS DETERMINED BY THE BALANCING
        '          SUBROUTINE  BALANC.  IF  BALANC  HAS NOT BEEN USED,
        '          SET LOW=1, IGH=N.
        '
        '        H CONTAINS THE UPPER HESSENBERG MATRIX.  INFORMATION ABOUT
        '          THE TRANSFORMATIONS USED IN THE REDUCTION TO HESSENBERG
        '          FORM BY  ELMHES  OR  ORTHES, IF PERFORMED, IS STORED
        '          IN THE REMAINING TRIANGLE UNDER THE HESSENBERG MATRIX.
        '
        '     ON OUTPUT
        '
        '        H HAS BEEN DESTROYED.  THEREFORE, IT MUST BE SAVED
        '          BEFORE CALLING  HQR  IF SUBSEQUENT CALCULATION AND
        '          BACK TRANSFORMATION OF EIGENVECTORS IS TO BE PERFORMED.
        '
        '        WR AND WI CONTAIN THE REAL AND IMAGINARY PARTS,
        '          RESPECTIVELY, OF THE EIGENVALUES.  THE EIGENVALUES
        '          ARE UNORDERED EXCEPT THAT COMPLEX CONJUGATE PAIRS
        '          OF VALUES APPEAR CONSECUTIVELY WITH THE EIGENVALUE
        '          HAVING THE POSITIVE IMAGINARY PART FIRST.  IF AN
        '          ERROR EXIT IS MADE, THE EIGENVALUES SHOULD BE CORRECT
        '          FOR INDICES IERR+1,...,N.
        '
        '        IERR IS SET TO
        '          ZERO       FOR NORMAL RETURN,
        '          J          IF THE LIMIT OF 30*N ITERATIONS IS EXHAUSTED
        '                     WHILE THE J-TH EIGENVALUE IS BEING SOUGHT.
        '
        '     ------------------------------------------------------------------

        Dim i&, j&, k&, L&, m&, en&, ll&, mm&, na&, itn&, its&, MP2&, ENM2&
        Dim p#, q#, r#, s#, t#, w#, x#, y#, ZZ#, tst1#, tst2#
        Dim NOTLAS As Boolean
        '
        Ierr = 0
        k = 1

Lab50:
        '
        en = igh
        t = 0.0#
        itn = 30 * N
        '     .......... SEARCH FOR NEXT EIGENVALUES ..........
Lab60:
        If (en < low) Then GoTo Lab1001
        its = 0
        na = en - 1
        ENM2 = na - 1
        '     .......... LOOK FOR DoubleSMALL SUB-DIAGONAL ELEMENT
        '                FOR L=EN STEP -1 UNTIL LOW DO -- ..........
Lab70:
        For ll = low To en
            L = en + low - ll
            If (L = low) Then GoTo Lab100
            s = Abs(h(L - 1, L - 1)) + Abs(h(L, L))
            If (s = 0) Then s = 1 's = norm  ' fix bug 2.11.05 VL
            tst1 = s
            tst2 = tst1 + Abs(h(L, L - 1))
            If (tst2 = tst1) And Abs(h(L, L - 1)) < 1 Then GoTo Lab100
        Next ll
        '     .......... FORM SHIFT ..........
Lab100:
        x = h(en, en)
        If (L = en) Then GoTo Lab270
        y = h(na, na)
        w = h(en, na) * h(na, en)
        If (L = na) Then GoTo Lab280
        If (itn = 0) Then GoTo Lab1000
        If ((its <> 10) And (its <> 20)) Then GoTo Lab130
        '     .......... FORM EXCEPTIONAL SHIFT ..........
        t = t + x
        '
        For i = low To en
            h(i, i) = h(i, i) - x
        Next i
        '
        s = Abs(h(en, na)) + Abs(h(na, ENM2))
        x = 0.75 * s
        y = x
        w = -0.4375 * s * s
Lab130:
        its = its + 1
        itn = itn - 1
        '     .......... LOOK FOR TWO CONSECUTIVE SMALL
        '                SUB-DIAGONAL ELEMENTS.
        '                FOR M=EN-2 STEP -1 UNTIL L DO -- ..........
        For mm = L To ENM2
            m = ENM2 + L - mm
            ZZ = h(m, m)
            r = x - ZZ
            s = y - ZZ
            p = (r * s - w) / h(m + 1, m) + h(m, m + 1)
            q = h(m + 1, m + 1) - ZZ - r - s
            r = h(m + 2, m + 1)
            s = Abs(p) + Abs(q) + Abs(r)
            p = p / s
            q = q / s
            r = r / s
            If (m = L) Then GoTo Lab150
            tst1 = Abs(p) * (Abs(h(m - 1, m - 1)) + Abs(ZZ) + Abs(h(m + 1, m + 1)))
            tst2 = tst1 + Abs(h(m, m - 1)) * (Abs(q) + Abs(r))
            If (tst2 = tst1) Then GoTo Lab150
        Next mm
        '
Lab150:
        MP2 = m + 2
        '
        For i = MP2 To en
            h(i, i - 2) = 0.0#
            If (i <> MP2) Then h(i, i - 3) = 0.0#
        Next i
        '     .......... DOUBLE QR STEP INVOLVING ROWS L TO EN AND
        '                COLUMNS M TO EN ..........
        For k = m To na
            NOTLAS = k <> na
            If (k = m) Then GoTo Lab170
            p = h(k, k - 1)
            q = h(k + 1, k - 1)
            r = 0.0#
            If (NOTLAS) Then r = h(k + 2, k - 1)
            x = Abs(p) + Abs(q) + Abs(r)
            If (x = 0.0#) Then GoTo Lab260
            p = p / x
            q = q / x
            r = r / x
Lab170:
            s = dsign(Math.Sqrt(p * p + q * q + r * r), p)
            If (k = m) Then GoTo Lab180
            h(k, k - 1) = -s * x
            GoTo Lab190
Lab180:
            If (L <> m) Then h(k, k - 1) = -h(k, k - 1)
Lab190:
            p = p + s
            x = p / s
            y = q / s
            ZZ = r / s
            q = q / p
            r = r / p
            If (NOTLAS) Then GoTo Lab225
            '     .......... ROW MODIFICATION ..........
            For j = k To N
                p = h(k, j) + q * h(k + 1, j)
                h(k, j) = h(k, j) - p * x
                h(k + 1, j) = h(k + 1, j) - p * y
            Next j
            '
            j = Math.Min(en, k + 3)
            '     .......... COLUMN MODIFICATION ..........
            For i = 1 To j
                p = x * h(i, k) + y * h(i, k + 1)
                h(i, k) = h(i, k) - p
                h(i, k + 1) = h(i, k + 1) - p * q
            Next i
            GoTo Lab255
Lab225:
            '     .......... ROW MODIFICATION ..........
            For j = k To N
                p = h(k, j) + q * h(k + 1, j) + r * h(k + 2, j)
                h(k, j) = h(k, j) - p * x
                h(k + 1, j) = h(k + 1, j) - p * y
                h(k + 2, j) = h(k + 2, j) - p * ZZ
            Next j
            '
            j = Math.Min(en, k + 3)
            '     .......... COLUMN MODIFICATION ..........
            For i = 1 To j
                p = x * h(i, k) + y * h(i, k + 1) + ZZ * h(i, k + 2)
                h(i, k) = h(i, k) - p
                h(i, k + 1) = h(i, k + 1) - p * q
                h(i, k + 2) = h(i, k + 2) - p * r
            Next i
Lab255:
            '
        Next k
Lab260:
        '
        GoTo Lab70
        '     .......... ONE ROOT FOUND ..........
Lab270:
        wr(en) = x + t
        wi(en) = 0.0#
        en = na
        GoTo Lab60
        '     .......... TWO ROOTS FOUND ..........
Lab280:
        p = (y - x) / 2.0#
        q = p * p + w
        ZZ = Math.Sqrt(Abs(q))
        x = x + t
        If (q < 0.0#) Then GoTo Lab320
        '     .......... REAL PAIR ..........
        ZZ = p + dsign(ZZ, p)
        wr(na) = x + ZZ
        wr(en) = wr(na)
        If (ZZ <> 0.0#) Then wr(en) = x - w / ZZ
        wi(na) = 0.0#
        wi(en) = 0.0#
        GoTo Lab330
        '     .......... COMPLEX PAIR ..........
Lab320:
        wr(na) = x + p
        wr(en) = x + p
        wi(na) = ZZ
        wi(en) = -ZZ
Lab330:
        en = ENM2
        GoTo Lab60
        '     .......... SET ERROR -- ALL EIGENVALUES HAVE NOT
        '                CONVERGED AFTER 30*N ITERATIONS ..........
Lab1000:
        Ierr = en
Lab1001:
    End Sub
    Private Function dsign(ByVal x, ByVal y)
        If y >= 0 Then
            dsign = Abs(x)
        Else
            dsign = -Abs(x)
        End If
    End Function
    Private Function MatMopUp(ByVal Mat, ByVal ErrMin)
        'eliminates values too small
        Dim A
        If ErrMin <> Nothing Then ErrMin = 10 ^ -14
        A = Mat

        For i = 1 To UBound(A, 1)
            For j = 1 To UBound(A, 2)
                If IsNumeric(A(i, j)) Then
                    If Abs(A(i, j)) < ErrMin Then A(i, j) = 0
                End If
            Next j
        Next i
        MatMopUp = A
        Return MatMopUp
    End Function
    Private Function MatrixSort(ByVal A, ByVal Order)
        '
        'Sorting Routine with swapping algorithm
        'A() may be matrix (N x M) or vector (N)
        'Sort is always based on the first column
        'Order = A (Ascending), D (Descending)
        'Note: it's simple but slow. Use only in non critical part
        '
        Dim flag_exchanged As Boolean
        Dim i_min&, i_max&, j_min&, j_max&, i&, k&, j&
        Dim c As Double

        i_min = 1
        i_max = UBound(A, 1)
        j_min = 1
        j_max = UBound(A, 2)

        'Sorting algortithm begin
        Do
            flag_exchanged = False
            For i = i_min To i_max Step 2
                k = i + 1
                If k > i_max Then Exit For
                If (A(i, j_min) > A(k, j_min) And Order = "A") Or _
                   (A(i, j_min) < A(k, j_min) And Order = "D") Then
                    'swap rows
                    For j = j_min To j_max
                        c = A(k, j)
                        A(k, j) = A(i, j)
                        A(i, j) = c
                    Next j
                    flag_exchanged = True
                End If
            Next
            If i_min = LBound(A, 1) Then
                i_min = LBound(A, 1) + 1
            Else
                i_min = LBound(A, 1)
            End If
        Loop Until flag_exchanged = False And i_min = LBound(A, 1)
        Return A
    End Function


    '-------------------------

    Private Function MatEigenvector(ByVal Mat(,) As Double, ByVal Eigenvalues(,) As Double, ByVal MaxErr As Double)
        'returns the eigenvector associate to a given eigenvalue
        'Eigenvalues may be also a vector of eigenvalues
        Dim L(,), u(,), w(,), k_max As Double
        Dim A(Mat.GetLength(0) - 1, Mat.GetLength(1) - 1) As Double
        MaxErr = 10 ^ -10
        Array.Copy(Mat, A, Mat.Length)


        L = Eigenvalues
        If IsArray(L) Then
            k_max = UBound(L)
        Else
            k_max = 1
        End If
        Dim N = UBound(A, 1)
        ReDim w(N, N)
        Dim Lk As Double

        Dim k As Integer = k_max - 2
        Do Until k > k_max

            Lk = L(k, 1)
            Array.Copy(Mat, A, Mat.Length)
            If k > 1 Then

            End If

            For i = 1 To N
                A(i, i) = A(i, i) - Lk
            Next
            'solve singolar system

            u = SysLinSing(A, MaxErr)
            Form1.ProgressBar1.Increment(1)
            Application.DoEvents()
            'inspection of U
            Dim m As Integer = 0
            For i = 1 To N
                'note VL 24-11-02
                '      m=0 => non singular matrix
                '      m=1 => eigenvalue simple
                '      m>1 => eigenvalue multiple
                If u(i, i) = 1 Then m = m + 1
            Next
            If m = 0 Then  'matrix not singular
                MatEigenvector = u
                Exit Function
            End If
            'copy not null eigenvectors from U to W
            For j = 1 To N
                If u(j, j) = 1 Then
                    For i = 1 To N
                        w(i, k) = u(i, j)
                    Next i
                    k = k + 1
                    Exit For 'fix bug: multiplicity eigenvalues 24.11.02 VL
                End If
            Next j
        Loop
        'mod. 8.1.06
        Dim tol As Double
        Normalize_Eigenvector_Sign(w, tol)
        NormalizeMatrix(w, 2, tol)

        Return w
    End Function

    Private Function SysLinSing(ByVal Mat(,) As Double, ByVal MaxErr As Double)
        'Solve a singular system
        'This version accept system where n°equations <= n°variables
        'risolve il sistema lineare singolare [Mat]x=V
        'Input parameters
        'Mat = rectangular matrix n x m
        'V = vettore n x 1
        'MaxErr= relative error for approximate zero element
        'Output: returns a matrix C (nxn) or C+d (nx(n+1)) if V is present
        'that is the linear tranformation matrix  y=[C]x+d
        'rev. 27-7-2006  .
        Dim a1(,), A(,) As Double, m, N, Det, tol#
        Dim b, elem_max, count1%, count2%

        a1 = Mat
        Dim na As Integer = UBound(a1, 1)
        Dim ma As Integer = UBound(a1, 2)
        Dim mb As Double
        Dim nb As Double = na : mb = 1

        If na <> nb Or mb <> 1 Then
            SysLinSing = "?" : Exit Function
        End If

        'load full matrix
        N = Math.Max(na, ma) : m = 1
        elem_max = 0
        ReDim A(N, ma + m)
        For i = 1 To na
            For j = 1 To ma
                A(i, j) = a1(i, j)
                If Abs(A(i, j)) > elem_max Then elem_max = Abs(A(i, j))
            Next j
            For j = 1 To m
                A(i, j + ma) = 0
                'If Not IsMissing(v) Then A(i, j + ma) = b(i, j)
            Next j
        Next i
        '
        tol = MaxErr
        Dim MaxErrRel As Double
        If elem_max > 1 Then MaxErrRel = tol * elem_max Else MaxErrRel = tol
        If MaxErrRel > 10 ^ -6 Then MaxErrRel = 10 ^ -6

        Call GaussJordan(A, N, N + m, Det, "D", MaxErrRel)

        'analizza soluzioni: infinite o nessuna
        'reorg matrix
        m = N + m
        'mopup matrix -------  VL 2-11-02
        For i = 1 To N
            For j = 1 To m
                If Abs(A(i, j)) < MaxErrRel Then A(i, j) = 0
            Next j
        Next i
        '-----------------
        For i = 1 To N
            'search for row null except one element (if exist)
            Dim Count As Integer = 0
            Dim i1 As Integer = 0
            For j = 1 To N
                If A(i, j) <> 0 Then
                    Count = Count + 1
                    i1 = j
                End If
            Next j
            If Count = 1 And i1 <> i Then
                'swap rows i and i1
                SwapRow(A, i, i1)
            End If
            If Count = 0 Then
                'check if the problem is impossible
                For j = N + 1 To m
                    If A(i, j) <> 0 Then GoTo Error_Handler
                Next j
            End If
        Next i

        For k = 1 To N
            If A(k, k) <> 0 Then
                'cerca un altro elemento non zero sopra
                For i = k - 1 To 1 Step -1
                    If A(i, k) <> 0 And i <> k Then
                        Dim PIa, pk As Double
                        PIa = -A(i, k)
                        pk = A(k, k)
                        For j = 1 To m
                            'linear combination between i1 and i rows
                            A(i, j) = pk * A(i, j) + PIa * A(k, j)
                            '                    If Abs(a(i, j)) < tol Then a(i, j) = 0
                        Next j
                    End If
                Next i
            End If
        Next k

        'normalize
        For i = 1 To N
            If A(i, i) <> 0 And A(i, i) <> 1 Then
                Dim PIa As Double
                PIa = A(i, i)
                For j = 1 To m
                    A(i, j) = A(i, j) / PIa
                Next j
            End If
        Next i
        'check rank of system matrix and augmented matrix
        For i = 1 To N
            count1 = 0 : count2 = 0
            For j = 1 To m
                If Abs(A(i, j)) > tol Then
                    If j <= N Then count1 = count1 + 1 Else count2 = count2 + 1
                End If
            Next j
            If count1 = 0 And count2 > 0 Then GoTo Error_Handler
        Next i

        For j = 1 To N
            If A(j, j) = 0 Then
                For i = 1 To N
                    A(i, j) = -A(i, j)
                Next i
                A(j, j) = 1
            Else
                A(j, j) = 0
            End If
        Next j
        Dim tmp(,) As Double
        tmp = MatMopUp(A, tol)
        Return tmp

        Exit Function
Error_Handler:
        SysLinSing = "?"
    End Function
    Private Sub GaussJordan(ByVal A, ByVal N, ByVal m, ByVal Det, ByVal f, ByVal dTiny)
        '==============================================================
        'Gauss-Jordan algorithm for triangle-diagonal matrix reduction
        'with partial pivot method
        'A = Matrix (n x m), m >= n
        'det =determinant of A (n x n)
        'f = type of reduction : T triangle, D diagonal
        'this version apply the check for too small elements: |aij|<Tiny
        'rev. version of 23-6-2002
        '==============================================================
        Dim i As Integer, j As Integer, k As Integer
        Dim w As Integer
        Det = 1
        For k = 1 To N
            If f = "T" Then
                w = k + 1 ' Triangolarizza
            ElseIf f = "D" Then
                w = 1     ' Diagonalizza
            Else
                Exit Sub
            End If
            'search max pivot in column k
            Dim ipivot As Integer
            ipivot = k
            Dim PivotMax As Double = Abs(A(k, k))
            For i = k + 1 To N
                If Abs(A(i, k)) > PivotMax Then
                    ipivot = i : PivotMax = Abs(A(i, k))
                End If
            Next i
            ' swap row
            If ipivot > k Then
                SwapRow(A, k, ipivot)
                Det = -Det
            End If

            ' check pivot 0
            If Abs(A(k, k)) <= dTiny Then
                A(k, k) = 0
                Det = 0
                Exit Sub
            End If

            'normalization
            Dim pk As Double
            pk = A(k, k)
            Det = Det * pk
            For j = 1 To m
                A(k, j) = A(k, j) / pk
            Next j

            'linear reduction
            For i = w To N
                If i <> k And A(i, k) <> 0 Then
                    pk = A(i, k)
                    For j = 1 To m
                        A(i, j) = A(i, j) - pk * A(k, j)
                    Next j
                End If
            Next i
        Next k
    End Sub
    Sub SwapRow(ByVal A, ByVal k, ByVal i)
        'Swaps rows k and i
        Dim j&, temp
        For j = LBound(A, 2) To UBound(A, 2)
            temp = A(i, j)
            A(i, j) = A(k, j)
            A(k, j) = temp
        Next
    End Sub
    Private Sub Normalize_Eigenvector_Sign(ByVal A, ByVal tol)
        'normalize the sign of each eigenvector making positive the first
        'non zero element  |aij| > tol
        Dim N&, m&, i&, j&, sg
        tol = 2 * 10 ^ -15
        N = UBound(A, 1)
        m = UBound(A, 2)
        For j = 1 To m
            sg = 0
            For i = 1 To N
                If Abs(A(i, j)) > 1000 * tol Then
                    If sg = 0 Then sg = Sign(A(i, j))
                    If sg < 0 Then
                        A(i, j) = -A(i, j)
                    Else
                        Exit For 'exit inner for
                    End If
                End If
            Next i
        Next j
    End Sub
    Private Sub NormalizeMatrix(ByVal A, ByVal NormType, ByVal tiny)
        'NormType = 1 (scaled to absolute max)
        '2 (module=1)
        '3 (scaled to absolute min)
        '4 (scaled to absolute mean)
        '5 (normalized mean = 0 and stdev = 1)
        'mod. 11-1-2007
        tiny = 2 * 10 ^ -14
        Dim N As Integer = UBound(A, 1)
        Dim m As Integer = UBound(A, 2)
        Dim n_min As Integer = LBound(A, 1)
        Dim m_min As Integer = LBound(A, 2)
        'mop-up
        For j = m_min To m
            For i = n_min To N
                If Abs(A(i, j)) < tiny Then A(i, j) = 0 '20-12-03
        Next i, j
        'normalize
        Dim s1 As Double
        For j = m_min To m


            s1 = 0 '
            For i = n_min To N
                s1 = s1 + A(i, j) ^ 2
            Next i
            s1 = Sqrt(s1)

            If Abs(s1) > tiny Then  'fix bug for null vectors . 6-6-05 VL
                For i = n_min To N
                    A(i, j) = A(i, j) / s1
                Next i
            End If
        Next j
    End Sub
End Module
