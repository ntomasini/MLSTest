
Imports System.IO
Module SNPmod








    Public Function duplicate(ByVal a(,) As String, ByVal lseq As Integer, ByVal nseq As Integer) As String(,)
        Dim otu As String(,)
        Dim i As Integer
        Dim base As String
        otu = a

        Do While i < lseq
            For j = 1 To nseq + 1
                base = otu(j, 1).Substring(i, 1)
                base = basecod(base)
                If i <> 0 Then
                    otu(j, 1) = otu(j, 1).Substring(0, i) & base & otu(j, 1).Substring(i + 1, lseq - (i + 1))
                Else
                    otu(j, 1) = base & otu(j, 1).Substring(1, lseq - 1)


                End If



            Next
            lseq = lseq + 1
            i = i + 2
        Loop
        Return otu
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

                    otu.SetValue(lin.Substring(1, lin.Length - 1), i + 1, 0)
                    i = i + 1
                Else
                    otu.SetValue(otu(i, 1) & lin.Substring(0, lin.Length - 1), i, 1)
                End If
            End If
        Loop
        lector.Close()

        Return otu
    End Function

    Function basecod(ByVal base As String) As String
        Select Case base
            Case "A", "T", "C", "G", "-", "N"
                base = base & base
            Case "R"
                base = "AG"
            Case "Y"
                base = "TC"
            Case "M"
                base = "AC"
            Case "S"
                base = "CG"
            Case "W"
                base = "AT"
            Case "K"
                base = "TG"

        End Select
        Return base
    End Function

    Public Function duplicateSNPonly(ByVal a(,) As String, ByVal lseq As Integer, ByVal nseq As Integer, ByVal pos As Integer) As String(,)
        Dim otu As String(,)

        Dim base As String
        otu = a


        For j = 1 To nseq + 1
            base = otu.GetValue(j, 1).ToString.Substring(pos, 1)
            base = basecod(base)
            If pos <> 0 Then
                otu.SetValue(otu.GetValue(j, 1).ToString.Substring(0, pos) & base & otu.GetValue(j, 1).ToString.Substring(pos + 1, lseq - (pos + 1)), j, 1)
            Else
                otu.SetValue(base & otu.GetValue(j, 1).ToString.Substring(1, lseq - 1), j, 1)


            End If



        Next

        Return otu
    End Function
    

    Public Function SNPmodX(ByVal nseq As Integer, ByVal OTUMAT1(,) As String, ByVal avstates As Boolean) As String(,)
        If avstates = False Then
            Dim seq1, seq2 As String

            Dim i, j As Integer



            Dim sequ As String = OTUMAT1.GetValue(1, 1)
            Dim lseq As Integer
            Dim consta As Boolean
            lseq = sequ.Length



            'eliminar sitios constantes
            Do While i < lseq
                consta = True
                For j = 1 To nseq
                    seq1 = OTUMAT1(j, 1)
                    seq2 = OTUMAT1(j + 1, 1)
                    If seq1.Substring(i, 1) <> seq2.Substring(i, 1) Then
                        consta = False
                        j = nseq
                    End If
                Next
                If consta = True Then

                    For t = 1 To nseq + 1
                        If i <> 0 Then
                            OTUMAT1.SetValue(OTUMAT1(t, 1).ToString.Substring(0, i) & OTUMAT1.GetValue(t, 1).ToString.Substring(i + 1, OTUMAT1.GetValue(t, 1).ToString.Length - (i + 1)), t, 1)
                        Else
                            OTUMAT1.SetValue(OTUMAT1(t, 1).ToString.Substring(1, OTUMAT1.GetValue(t, 1).ToString.Length - 1), t, 1)
                        End If


                    Next
                    i = i - 1
                    lseq = lseq - 1

                End If
                i = i + 1
            Loop

            OTUMAT1 = duplicate(OTUMAT1, lseq, nseq)


        End If
        Return OTUMAT1






    End Function

End Module
