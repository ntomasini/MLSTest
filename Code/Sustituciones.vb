
Module Sustituciones
 
    Function translate(ByVal codon1 As String) As String
        Dim a As String
        Select Case codon1.Substring(0, 2)

            Case "TT"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "F"
                ElseIf codon1.Substring(2, 1) = "A" Or codon1.Substring(2, 1) = "G" Or codon1.Substring(2, 1) = "R" Then
                    a = "L"
                End If
            Case "CT"
                a = "L"
            Case "AT"
                If codon1.Substring(2, 1) = "G" Then
                    a = "M"
                ElseIf codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Or codon1.Substring(2, 1) = "M" Or codon1.Substring(2, 1) = "W" Or codon1.Substring(2, 1) = "M" Or codon1.Substring(2, 1) = "A" Then
                    a = "I"
                End If
            Case "GT"
                a = "V"
            Case "TC"
                a = "S"
            Case "CC"
                a = "P"
            Case "AC"
                a = "T"
            Case "GC"
                a = "A"
            Case "TA"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "Y"
                ElseIf codon1.Substring(2, 1) = "A" Or codon1.Substring(2, 1) = "G" Or codon1.Substring(2, 1) = "R" Then
                    a = "_"
                End If
            Case "CA"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "H"
                ElseIf codon1.Substring(2, 1) = "A" Or codon1.Substring(2, 1) = "G" Or codon1.Substring(2, 1) = "R" Then
                    a = "Q"
                End If
            Case "AA"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "N"
                ElseIf codon1.Substring(2, 1) = "A" Or codon1.Substring(2, 1) = "G" Or codon1.Substring(2, 1) = "R" Then
                    a = "K"
                End If
            Case "GA"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "D"
                ElseIf codon1.Substring(2, 1) = "A" Or codon1.Substring(2, 1) = "G" Or codon1.Substring(2, 1) = "R" Then
                    a = "E"
                End If
            Case "TG"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "C"
                ElseIf codon1.Substring(2, 1) = "G" Then
                    a = "W"
                Else
                    a = "_"
                End If
            Case "CG"
                a = "R"
            Case "AG"
                If codon1.Substring(2, 1) = "T" Or codon1.Substring(2, 1) = "C" Or codon1.Substring(2, 1) = "Y" Then
                    a = "S"
                ElseIf codon1.Substring(2, 1) = "A" Or codon1.Substring(2, 1) = "G" Or codon1.Substring(2, 1) = "R" Then
                    a = "R"
                End If
            Case "GG"
                a = "G"
            Case Else
                a = "?"
        End Select
        If a = Nothing Then
            a = "?"
        End If
        Return a
    End Function
End Module
