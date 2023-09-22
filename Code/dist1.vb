
Public Class medirdistancia
    Private _seq1 As String
    Private _seq2 As String
    Private boost As Boolean
    Private dist As Single
    Private seed As Integer

    Sub New()
        _seq1 = ""
        _seq2 = ""
        boost = False

    End Sub
    Public Function distance(ByVal _seq1 As String, ByVal _seq2 As String, ByVal length As Integer, ByVal boost As Boolean, ByVal seed As Double, ByVal distx As Single, ByVal bootstrapn() As Integer) As Single
        dist = 0
        Dim seqleng As Integer = length
        'If distx <> 0 Then
        ' Dim borrar As Integer = _seq2.Length
        Dim i As Integer = 0
        Dim g As Integer



        If boost = False Then

            For i = 0 To _seq1.Length - 1
                Dim s1 As Char = _seq1.Chars(i)
                Dim s2 As Char = _seq2.Chars(i)
                If s1 <> s2 Then
                    Dim p As Single
                    p = pesos(s1, s2)
                    If Form1.hethandpp = 2 Then
                        If p <> 1 And p <> 0 Then
                            p = 0
                            seqleng = seqleng - 1
                        End If
                    End If
                    If p = 20 Then
                        seqleng = seqleng - 1

                    Else
                        dist = dist + p
                    End If
                End If

            Next i

        Else
            i = 0

            If bootstrapn Is Nothing Then Exit Function
            Do While i < bootstrapn.Length
                g = bootstrapn(i)

                If _seq1.Chars(g) <> _seq2.Chars(g) Then


                    Dim p As Single
                    p = pesos(_seq1.Chars(g), _seq2.Chars(g))
                    If Form1.hethandpp = 2 Then
                        If p <> 1 And p <> 0 Then
                            p = 0
                            seqleng = seqleng - 1
                        End If
                    End If
                    If p = 20 Then
                        'seqleng = seqleng - 1

                    Else
                        dist = dist + p
                    End If
                End If


                i = i + 1
            Loop
            ' Catch
            ' End Try
        End If

        'End If
        If Form1.hethandpp <> 1 And Form1.pdison = True Then
            dist = dist / seqleng

        End If
        Return dist
    End Function
    Public Function pesos(ByVal subseq1 As Char, ByVal subseq2 As Char) As Single
        Dim peso As Byte
        Dim s1(1), s2(1) As Char


        'If subseq1 = "-" Or subseq1 = "N" Or subseq2 = "-" Or subseq2 = "N" Then
        'peso = 20
        ' Else

        s1 = convertH(subseq1)
        s2 = convertH(subseq2)
        If s1 = Nothing Or s2 = Nothing Then
            peso = 20
            Return peso
            Exit Function
        End If
        '''
        Dim s1l As Byte = s1.Length - 1
        Dim s2l As Byte = s2.Length - 1
        For x = 0 To s1l
            For y = 0 To s2l
                If s1(x) <> s2(y) Then
                    peso = peso + 1
                End If
            Next
        Next

        ' If s1(0) <> s2(0) Then
        'peso = peso + 1
        ' End If
        ' If s1(0) <> s2(1) Then
        'peso = peso + 1
        '  End If
        ' If s1(1) <> s2(0) Then
        'peso = peso + 1
        '   End If
        '  If s1(1) <> s2(1) Then
        'peso = peso + 1

        ' End If
       

        Dim peso1 As Single = peso / (s1.Length * s2.Length)
        Return peso1
    End Function
    Private Function convertH(ByVal s1 As Char) As Char()
        Dim s(0) As Char
        If s1 = "A"c Or s1 = "T"c Or s1 = "C"c Or s1 = "G"c Then

            s(0) = s1
            Return s
            Exit Function

        End If
        Select Case s1
            Case "R"c
                ReDim s(1)
                s(0) = "A"c
                s(1) = "G"c
            Case "Y"c
                ReDim s(1)
                s(0) = "T"c
                s(1) = "C"c
            Case "M"c
                ReDim s(1)
                s(0) = "A"c
                s(1) = "C"c
            Case "W"c
                ReDim s(1)
                s(0) = "A"c
                s(1) = "T"c
            Case "K"c
                ReDim s(1)
                s(0) = "T"c
                s(1) = "G"c
            Case "S"c
                ReDim s(1)
                s(0) = "C"c
                s(1) = "G"c
            Case "-"c
                ReDim s(1)
                s(0) = "X"c
                s(1) = "X"c
            Case "N"c
                Exit Function
            

            Case "V"
                ReDim s(2)
                s(0) = "A"c
                s(1) = "C"c
                s(2) = "G"c
            Case "D"
                ReDim s(2)
                s(0) = "A"c
                s(1) = "T"c
                s(2) = "G"c
            Case "H"
                ReDim s(2)
                s(0) = "A"c
                s(1) = "C"c
                s(2) = "T"c
            Case "B"
                ReDim s(2)
                s(0) = "T"c
                s(1) = "C"c
                s(2) = "G"c
        End Select

        Return s
    End Function
   

    

    Public Property seq1() As String
        Get
            Return _seq1
        End Get
        Set(ByVal cadena1 As String)
            _seq1 = cadena1
        End Set
    End Property
    Public Property seq2() As String
        Get
            Return _seq2
        End Get
        Set(ByVal cadena2 As String)
            _seq2 = cadena2
        End Set
    End Property
End Class
