Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Dialog3
    Private FileNamesTrees As String()

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim superSplit As splits
        Dim fsplits As New List(Of splits)
        'Dim fSplits(ListBox1.Items.Count - 1) As splits

        readtree(TextBox1.Text, superSplit)
        superSplit.lseq = 1000000000
        For i = 0 To FileNamesTrees.GetLength(0) - 1
            Dim x As splits
            fsplits.Add(x)

            readtree(FileNamesTrees(i), fsplits(fsplits.Count() - 1), fsplits)


        Next
        Dim suppnames As New List(Of String)

        suppnames.Add("support")



        If CheckBox1.Checked = True Then
            Form1.testnlsupport(superSplit, fsplits.ToArray)
            suppnames.Add("Consensus Support")
        End If
        If CheckBox2.Checked = True Then
            Form1.testcompatib(superSplit, fsplits.ToArray, -1)
            suppnames.Add("TopoIncog")
        End If
        If CheckBox3.Checked = True Then
            Form1.testcompatib(superSplit, fsplits.ToArray, TextBox2.Text)
            suppnames.Add("supported Topoincog")
        End If


        If CheckBox4.Checked = True Then
            Form1.Consnetmaker(fsplits.ToArray, TextBox3.Text, TextBox2.Text)
        End If

        Dim nwkform As New treeviewer2 With {.sptree = superSplit, .Text = "Tree Viewer", ._viewsupport = True, ._supportnames = suppnames.ToArray} '
        nwkform.Show()
        If CheckBox5.Checked = True Then
            Dim sp1 As splits
            Dim spp As New Consensus1
            sp1 = spp.consensussplits(fsplits.ToArray, True)
            Dim supnames(0) As String
            supnames(0) = "Consensus Support"
            Dim nwkformx As New treeviewer2 With {.Text = "Majority Rule Extended", .sptree = sp1, ._supportnames = supnames, ._viewsupport = True}
            nwkformx.Show()

        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        'Me.Close()




    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.ShowDialog()

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As CancelEventArgs) Handles OpenFileDialog1.FileOk
        If Me.OpenFileDialog1.Multiselect = False Then
            TextBox1.Text = OpenFileDialog1.FileName
        Else
            ListBox1.DataSource = OpenFileDialog1.SafeFileNames
            FileNamesTrees = OpenFileDialog1.FileNames
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.ShowDialog()

    End Sub

    Sub readtree(ByVal filename As String, ByRef mySplits As splits, Optional ByRef fsplits As List(Of splits) = Nothing)
        Dim ax(1) As String
        ax(1) = Form1.getfirstFilename()
        mySplits.otumat1 = Form1.conca(ax, 1, Form1.DataGridView7.Rows.Count, vbCrLf, True, True)
        Dim file As New System.IO.StreamReader(filename)
        Dim nwk As String

        nwk = file.ReadLine()
            convertNwkToSplits(nwk, mySplits)
            mySplits.lseq = 1000000000
        If fsplits IsNot Nothing Then

            nwk = file.ReadLine()

            Do While nwk IsNot Nothing AndAlso nwk.StartsWith("(")
                Dim x As splits
                x.lseq = 1000000000
                x.otumat1 = mySplits.otumat1

                convertNwkToSplits(nwk, x)
                fsplits.Add(x)
                nwk = file.ReadLine()

            Loop
        End If

    End Sub
    Function convertNwkToSplits(ByRef nwk As String, ByRef mysplit As splits)
        correctpolit(nwk)
        correctpolit(nwk)
        Dim nseqs As Integer = mysplit.otumat1.GetLength(0) - 1
        Dim nsplits As Integer = (nseqs - 1) * 2
        ReDim mysplit.nOTUs(nsplits, 3)
        ReDim mysplit.notus1(nsplits, 1)
        Dim opensplits(nsplits) As Boolean
        Dim lastsplit As Integer = nsplits + 1
        Dim currsplit As Integer = nsplits + 1


        Dim prevchar As Char = "("c


        Dim length As Integer = 0
        For i = 1 To nwk.Count - 1

            If nwk.Chars(i) = "("c Then

                lastsplit = lastsplit - 1
                opensplits(lastsplit) = True

                currsplit = lastsplit


            ElseIf nwk.Chars(i) = ")"c Then

                If currsplit <= nsplits Then
                    For t = currsplit To nsplits
                        If opensplits(t) = True Then
                            opensplits(t) = False
                            currsplit = t

                            Exit For


                        End If
                    Next



                Else
                    Exit For
                End If

            ElseIf nwk.Chars(i) = ";"c Then
                Exit For
            Else
                If nwk.Chars(i) = "," Then


                End If

                If prevchar = "("c OrElse prevchar = ","c Then
                    Dim id As Integer = readseqname(i, nwk, mysplit)

                    For n = nseqs + 1 To nsplits
                        If opensplits(n) = True Then
                            If mysplit.nOTUs(n, 1) = Nothing Then
                                mysplit.nOTUs(n, 1) = " "
                            End If
                            mysplit.nOTUs(n, 1) = mysplit.nOTUs(n, 1) & id & " "
                        End If
                    Next
                ElseIf prevchar = ")"c Then
                    readsupport(i, nwk, mysplit, currsplit)
                ElseIf prevchar = ":"c Then

                    readbrlenx(i, nwk, mysplit, currsplit)
                End If

            End If


            prevchar = nwk(i)
        Next

        For i = 1 To nsplits
            mysplit.nOTUs(i, 2) = Module1.stdsplit(mysplit.nOTUs(i, 1), nseqs)

        Next
        'Dim suppnames(0) As String
        'suppnames(0) = "support"
        'Dim nwkform As New treeviewer2 With {.sptree = mysplit, .Text = "Tree Viewer", ._viewsupport = True, ._supportnames = suppnames} '
        'nwkform.Show()

    End Function
    Function correctpolit(ByRef nwk As String)
        Dim level As Integer = 0

        Dim startgroup As Integer = 0
        Dim countcomma As Integer = 0
        Dim i As Integer = 0
        Do While i < nwk.Length

            If nwk.Chars(i) = "("c Then
                level = 0
                startgroup = i
                countcomma = 0
                For j = i + 1 To nwk.Length - 1
                    If nwk.Chars(j) = "("c Then
                        level = level + 1
                    ElseIf nwk.Chars(j) = ")"c Then
                        level = level - 1
                    ElseIf nwk.Chars(j) = ","c And level = 0 Then
                        Dim pruebas As String
                        pruebas = nwk.Substring(i, j - i)
                        countcomma = countcomma + 1
                        If countcomma > 1 Then
                            nwk = nwk.Insert(i, "(")
                            j = j + 1
                            nwk = nwk.Insert(j, "):0")
                            j = j + 3
                        End If

                    End If
                    If level = -1 Then
                        Exit For
                    End If
                Next
            End If
            i = i + 1
        Loop
    End Function
    Function readseqname(ByRef currpos As Integer, ByRef nwk As String, ByRef mysplits As splits) As Integer
        Dim name As String = ""
        Dim idname As Integer = -1
        For a = currpos To nwk.Length - 1
            If nwk.Chars(a) = ":"c OrElse nwk.Chars(a) = ")"c OrElse nwk.Chars(a) = "," Then
                currpos = a - 1
                Exit For
            Else
                name = name & nwk.Chars(a)

            End If
        Next

        For i = 1 To mysplits.otumat1.GetLength(0) - 1
            If mysplits.otumat1(i, 0) = name Then
                mysplits.nOTUs(i, 1) = " " & i & " "
                If nwk.Chars(currpos + 1) = ":"c Then
                    Dim a As Integer = currpos + 2

                    readbrlenx(a, nwk, mysplits, i)
                    currpos = a
                End If
                Return i
            End If
        Next

    End Function
    Function readsupport(ByRef currpos As Integer, ByRef nwk As String, ByRef mysplits As splits, ByVal currsplit As Integer)
        Dim supp As Single = 0
        Dim suppstring As String = ""
        For a = currpos To nwk.Length - 1
            If nwk.Chars(a) = ":"c Then
                currpos = a - 1
                Exit For
            Else
                suppstring = suppstring & nwk.Chars(a)

            End If
        Next
        If suppstring <> "" Then

            supp = Convert.ToSingle(suppstring)
        Else
            currpos = currpos + 1
        End If
        mysplits.notus1(currsplit, 1) = supp

    End Function
    Function readbrlenx(ByRef currpos As Integer, ByRef nwk As String, ByRef mysplits As splits, ByVal currsplit As Integer)
        Dim brlen As Single
        Dim brlenstring As String = ""
        For a = currpos To nwk.Length - 1
            If nwk.Chars(a) = ")"c OrElse nwk.Chars(a) = ","c Then
                currpos = a - 1
                Exit For
            Else
                brlenstring = brlenstring & nwk.Chars(a)

            End If
        Next
        brlen = Convert.ToSingle(brlenstring)
        mysplits.notus1(currsplit, 0) = brlen

    End Function

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged, CheckBox5.CheckedChanged

    End Sub
End Class
