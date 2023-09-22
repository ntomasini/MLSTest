Imports System.IO
Module Perfal
    Private Structure matdist
        Dim distancias As Single(,)
    End Structure

    Function DPo() As String(,)

        Dim lenR As Integer = Form1.DataGridView1.RowCount - 4
        Dim lenC As Integer = Form1.DataGridView1.ColumnCount - 1

        Dim perfiles(lenR - 1, lenC - 1) As String
        For i = 0 To lenR - 1
            For j = 0 To lenC - 1
                perfiles(i, j) = Form1.DataGridView1.Item(j + 1, i).Value
            Next
        Next
        Dim npol(lenC - 1) As Integer






        Dim out(1) As Integer
        Form1.DataGridView1.Item(0, lenR + 3).Value = "DP (95% Confidence Interval)"
        For i = 1 To lenC
            Dim psi(lenR - 1) As Single
            Dim suma As Integer = 0


            Dim dp As Single = DPc(perfiles, i, False, 0)
            For p = 1 To lenR
                psi(p - 1) = lenR * dp - (lenR - 1) * DPc(perfiles, i, True, p)
            Next
            Dim meani As Single = mean(psi)
            Dim var As Single = variance(psi, meani)
            Dim ciup As Single = meani + 2 * Math.Sqrt(var / lenR)
            If ciup > 1 Then ciup = 1
            Dim cilo As Single = meani - 2 * Math.Sqrt(var / lenR)
            If cilo < 0 Then cilo = 0
            Dim a As String = Math.Round(dp, 3) & "(" & Math.Round(cilo, 3) & "-" & Math.Round(ciup, 3) & ")"
            Form1.DataGridView1.Item(i, lenR + 3).Value = a
        Next






    End Function
    Function mean(ByVal psi() As Single) As Single
        Dim sum As Single
        For i = 0 To psi.length - 1
            sum = sum + psi(i)
        Next
        Return sum / psi.Length
    End Function
    Function variance(ByVal psi() As Single, ByVal mean As Single) As Single
        Dim sum As Single
        For i = 0 To psi.Length - 1
            sum = sum + (psi(i) - mean) ^ 2
        Next
        Return sum / (psi.Length - 1)
    End Function
    Function DPc(ByVal perfiles(,) As String, ByVal locus As Integer, ByVal jn As Boolean, ByVal psi As Integer) As Single
        Dim perfclone(,) As String = perfiles.Clone
        If jn = True Then
            perfclone(psi - 1, locus - 1) = 0
        End If
        Dim out(1) As Integer
        Dim maxi As Integer = max(perfclone, locus - 1, out)
        Dim suma As Single

        For p = 1 To maxi
            Dim count As Integer = 0
            For q = 0 To perfiles.GetLength(0) - 1

                If perfclone(q, locus - 1) = p Then
                    count = count + 1
                End If

            Next
            count = count * (count - 1)
            suma = suma + count
        Next
        Dim dp As Single
        If jn = False Then
            dp = 1 - (suma / ((perfiles.GetLength(0)) * (perfiles.GetLength(0) - 1)))
        Else
            dp = 1 - (suma / ((perfiles.GetLength(0) - 1) * (perfiles.GetLength(0) - 2)))
        End If

        Return dp
    End Function
    Function perf(ByVal a() As String, ByVal nsq As Integer, ByVal locus() As String, ByVal result As Boolean) As String(,)
        nseq = nsq - 1
        Dim seqmat(,) As String
        Dim perfiles(nsq, a.Length) As String
        Dim otumatred(,) As String
        Dim npol(a.Length - 1, 1) As Integer
        Dim out(1) As Integer

        For i = 1 To a.Length - 1
            seqmat = lectorperfal(a(i), nsq)
            otumatred = lectorperfal(a(i), nsq)
            lseq = seqmat(1, 1).Length

            otumatred = Module1.reductor(otumatred)
            npol(i, 0) = otumatred.GetValue(1, 1).ToString.Length

            perfiles = distanciasc(seqmat, nsq, perfiles, i)
            npol(i, 1) = max(perfiles, i, out)
        Next



        If result = False Then
            For i = 1 To a.Length
                If i <> 1 Then
                    Form1.DataGridView1.Columns.Add(i, locus(i - 2))
                Else

                    Form1.DataGridView1.Columns.Add(i, "Name")
                End If

            Next

            Form1.DataGridView1.Rows.Add(nsq + 4)

            For i = 0 To nsq - 1
                For j = 0 To a.Length - 1
                    If j = 0 Then
                        Form1.DataGridView1.Item(j, i).Value = seqmat(i + 1, 0).ToString

                    Else

                        Form1.DataGridView1.Item(j, i).Value = perfiles(i + 1, j).ToString
                    End If

                Next
            Next
            Form1.DataGridView1.Item(0, nsq).Value = "Number of Alleles"
            Form1.DataGridView1.Item(0, nsq + 1).Value = "Number of Polymorphisms"
            Form1.DataGridView1.Item(0, nsq + 2).Value = "Typing Efficiency"
            Form1.DataGridView1.Item(0, nsq + 3).Value = "Discriminatory Power"
            Dim nsq1 As Integer = nsq
            Dim dp As Single = 0

            For i = 1 To a.Length - 1

                Dim eff As Single = (npol(i, 1) / npol(i, 0))
              
                dp = DPc(perfiles, i + 1, False, 0)


                Form1.DataGridView1.Item(i, nsq).Value = npol(i, 1)
                Form1.DataGridView1.Item(i, nsq + 1).Value = npol(i, 0)
                Form1.DataGridView1.Item(i, nsq + 2).Value = Math.Round(eff, 3)
                Form1.DataGridView1.Item(i, nsq + 3).Value = Math.Round(dp, 3)

                Form1.DataGridView1.Item(i, nsq).Style.BackColor = Color.Aquamarine
                Form1.DataGridView1.Item(i, nsq + 1).Style.BackColor = Color.Aquamarine
                Form1.DataGridView1.Item(i, nsq + 2).Style.BackColor = Color.Aquamarine
                Form1.DataGridView1.Item(i, nsq + 3).Style.BackColor = Color.Aquamarine
            Next
            Form1.DataGridView1.Columns.Add("ST", "ST")
            Dim totalpol, totalDST As Integer
            For i = 0 To npol.GetLength(0) - 1
                totalpol = totalpol + npol(i, 0)
                totalDST = totalDST + npol(i, 1)
            Next
            Dim suma As Integer = 0

           

            Form1.DataGridView1.Item(a.Length, nsq + 1).Value = totalpol
            Form1.DataGridView1.Item(a.Length, nsq + 2).Value = Math.Round(totalDST / totalpol, 3)

        Else
            Return perfiles
        End If
    End Function
    Function distanciasc(ByVal otumat1(,) As String, ByVal nsq As Integer, ByVal perfiles(,) As String, ByVal n As Integer) As String(,)

        Dim seqs(1) As String

        Dim distancias(nsq, nsq) As Single
        Dim i, j As Integer
        Dim bootstrapn() As Integer
        Dim dist As medirdistancia

        lseq = otumat1(1, 1).ToString.Length


        Dim sequ1, sequ2 As String

        dist = New medirdistancia


        For i = 0 To nsq

            For j = i + 1 To nsq
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
        Dim dst As Integer = 1


        For i = 0 To nsq - 1
            If i = 0 Then
                For j = 1 To nsq
                    perfiles(j, 0) = otumat1(j, 0)
                Next

            Else


                If perfiles(i, n) = 0 Then
                    perfiles(i, n) = dst
                    For j = i + 1 To nsq
                        If distancias(i, j) = 0 Then
                            perfiles(j, n) = dst

                        End If

                    Next
                    dst = dst + 1
                Else

                End If
                If i = nsq - 1 And perfiles(i + 1, n) = 0 Then
                    perfiles(i + 1, n) = dst
                End If
            End If

        Next

        Return perfiles
    End Function
    Function max(ByVal perfiles(,) As String, ByVal nc As Integer, ByVal out() As Integer) As Integer
        Dim nmax As Integer
        For i = 0 To perfiles.GetLength(0) - 1
            If i <> out(0) And i <> out(1) Then
                If perfiles(i, nc) > nmax Then
                    nmax = perfiles(i, nc)
                End If
            End If
        Next
        Return nmax
    End Function
    Function lectorperfal(ByVal a As String, ByVal nseq As Integer) As String(,)
        Dim lector As TextReader
        Dim i, j, k, l, m As Integer
        Dim concat(nseq, 1) As String



        Dim n As Integer = 1
        Dim lin As String = 1
        Dim length As Integer






        lector = New StreamReader(a)
        n = 0
        lin = 1
        Dim index As Integer = 0
        Dim ax As Integer = 0
        Do While ax < 2

            lin = lector.ReadLine()
            If lin <> Nothing Then
                If lin.Substring(0, 1) = ">" Then
                    If Form1.DataGridView7.Item(1, index).Value = True Then
                        concat(n + 1, 0) = Form1.DataGridView7.Item(0, index).Value
                        n = n + 1

                    End If
                    index = index + 1
                    ax = 0
                Else
                    If Form1.DataGridView7.Item(1, index - 1).Value = True Then
                        concat(n, 1) = concat(n, 1) & lin.ToUpper
                    End If

                End If
            Else
                ax = ax + 1
            End If

        Loop



        lector.Close()






        Return concat
    End Function
    Function perfx(ByVal a() As String, ByVal nsq As Integer, ByVal locus() As String) As Single(,)
        nseq = nsq - 1
        Dim matdist(a.Length - 1) As matdist

        Dim seqmat(,) As String
        Dim perfiles(nsq, a.Length) As String
        Dim meandistances(nsq, nsq) As Single
        Dim distfinal(nsq, nsq) As Single
        Dim out(1) As Integer

        For i = 1 To a.Length - 1

            seqmat = lectorperfal(a(i), nsq)

            lseq = seqmat(1, 1).Length




            matdist(i).distancias = distanciasx(seqmat, nsq)

        Next

        For j = 1 To nseq
            For k = j + 1 To nseq + 1
                Dim mean As Single
                For i = 1 To a.Length - 1
                    mean = mean + matdist(i).distancias(j, k)
                Next
                mean = mean / (a.Length - 1)
                'mean = 0
                Dim count As Integer = 0
                'If j = 23 And k = 24 Then Stop
                For i = 1 To a.Length - 1
                    If matdist(i).distancias(j, k) <= mean - 0.1 Or matdist(i).distancias(j, k) >= mean + 0.1 Then

                        count = count + 1
                    End If
                Next
                'If count = 1 Then Stop
                meandistances(j, k) = count / (a.Length - 1)
                meandistances(k, j) = meandistances(j, k)
            Next

        Next


        Return meandistances
    End Function
    Function distanciasx(ByVal otumat1(,) As String, ByVal nsq As Integer) As Single(,)

        Dim seqs(1) As String

        Dim distancias(nsq, nsq) As Single
        Dim i, j As Integer
        Dim bootstrapn() As Integer
        Dim dist As medirdistancia

        lseq = otumat1(1, 1).ToString.Length


        Dim sequ1, sequ2 As String

        dist = New medirdistancia


        For i = 0 To nsq

            For j = i + 1 To nsq
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
        Return distancias
    End Function

End Module
