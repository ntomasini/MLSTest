Imports System.Windows.Forms

Public Class Dialog1

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Dim locus(Form1.GridView1.Rows.Count - 1) As String
        Dim makefiles As Boolean
        Dim i As Integer
        For i = 0 To Form1.GridView1.Rows.Count - 1
            locus(i) = Form1.GridView1.Item(1, i).Value
        Next
        Dim limit As Integer
        If RadioButton1.Checked = True Then
            limit = 1
            makefiles = True
        Else
            limit = ListBox1.CheckedIndices.Count
            FolderBrowserDialog1.SelectedPath = Nothing
            FolderBrowserDialog1.ShowDialog()
            If FolderBrowserDialog1.SelectedPath <> Nothing Then
                makefiles = True
            Else
                makefiles = False
            End If
        End If
        If makefiles = True Then
            For z = 1 To limit
                Dim a() As String
                If RadioButton1.Checked = True Then
                    ReDim a(ListBox1.CheckedItems.Count)
                    For i = 1 To ListBox1.CheckedItems.Count
                        a(i) = locus(ListBox1.CheckedIndices(i - 1))
                    Next
                Else
                    ReDim a(1)

                    a(1) = locus(ListBox1.CheckedIndices(z - 1))

                End If




                otumat1 = Form1.conca(a, a.Length - 1, Form1.TextBox12.Text, Form1.TextBox13.Text, True, False)
                '''
                'otumat1 = Form1.permuteseqs(otumat1)
                '''
                Dim aaa(otumat1.GetLength(0) - 1) As String
                For ii = 0 To otumat1.GetLength(0) - 1
                    aaa(ii) = otumat1(ii, 0)

                Next
                If CheckBox3.Checked = True Then
                    Array.Sort(aaa)
                End If
                Dim otumat2(otumat1.GetLength(0) - 1, otumat1.GetLength(1) - 1) As String
                For ii = 1 To otumat1.GetLength(0) - 1
                    Dim index As Integer = aaa.IndexOf(aaa, otumat1(ii, 0))
                    otumat2(index, 0) = otumat1(ii, 0)
                    otumat2(index, 1) = otumat1(ii, 1)
                Next
                otumat1 = otumat2
                If CheckBox1.Checked = True Or CheckBox2.Checked = True Then
                    Dim seq1, seq2 As String

                    Dim j As Integer
                    i = 0


                    Dim sequ As String = otumat1.GetValue(1, 1)
                    Dim lseq As Integer
                    Dim consta As Boolean
                    nseq = otumat1.GetLength(0) - 2
                    lseq = sequ.Length



                    'eliminar sitios constantes

                    Do While i < lseq
                        consta = True
                        For j = 1 To nseq
                            seq1 = otumat1(j, 1)
                            seq2 = otumat1(j + 1, 1)
                            If seq1.Substring(i, 1) <> seq2.Substring(i, 1) Then
                                consta = False
                                j = nseq
                            End If
                        Next
                        If consta = True Then
                            If CheckBox1.Checked = True Then
                                For t = 1 To nseq + 1
                                    If i <> 0 Then
                                        otumat1.SetValue(otumat1(t, 1).ToString.Substring(0, i) & otumat1.GetValue(t, 1).ToString.Substring(i + 1, otumat1.GetValue(t, 1).ToString.Length - (i + 1)), t, 1)
                                    Else
                                        otumat1.SetValue(otumat1(t, 1).ToString.Substring(1, otumat1.GetValue(t, 1).ToString.Length - 1), t, 1)
                                    End If


                                Next
                                i = i - 1

                            End If


                        Else


                            If CheckBox2.Checked = True Then
                                otumat1 = duplicateSNPonly(otumat1, lseq, nseq, i)


                                i = i + 1
                            End If

                        End If
                        i = i + 1
                        lseq = otumat1(1, 1).Length
                    Loop

                End If
                Dim texto As String
                Dim xx As String

                If RadioButton3.Checked = True Then

                    xx = "Format " & "FASTA" & "|" & "*." & "fas"

                    SaveFileDialog1.Filter = xx
                    SaveFileDialog1.DefaultExt = "fas"
                    If RadioButton2.Checked = True Then
                        makefasta(otumat1, FolderBrowserDialog1.SelectedPath & "\" & ListBox1.CheckedItems(z - 1) & "." & SaveFileDialog1.DefaultExt)

                    End If
                    If RadioButton1.Checked = True Then
                        SaveFileDialog1.ShowDialog()
                        If SaveFileDialog1.FileName <> Nothing Then
                            makefasta(otumat1, SaveFileDialog1.FileName)
                        End If
                    End If
                ElseIf RadioButton4.Checked = True Then

                    xx = "Format " & "Phylip" & "|" & "*." & "phy"
                    SaveFileDialog1.Filter = xx
                    SaveFileDialog1.DefaultExt = "phy"
                    If RadioButton2.Checked = True Then
                        makephylip(otumat1, FolderBrowserDialog1.SelectedPath & "\" & ListBox1.CheckedItems(z - 1) & "." & SaveFileDialog1.DefaultExt)

                    End If
                    If RadioButton1.Checked = True Then
                        SaveFileDialog1.ShowDialog()
                        If SaveFileDialog1.FileName <> Nothing Then
                            makephylip(otumat1, SaveFileDialog1.FileName)
                        End If
                    End If
                Else
                    'texto = makeNexus(otumat1)
                    xx = "Format " & "Nexus" & "|" & "*." & "nex"
                    SaveFileDialog1.Filter = xx
                    SaveFileDialog1.DefaultExt = "nex"
                    If RadioButton2.Checked = True Then
                        makeNexus(otumat1, FolderBrowserDialog1.SelectedPath & "\" & ListBox1.CheckedItems(z - 1) & "." & SaveFileDialog1.DefaultExt)

                    End If
                    If RadioButton1.Checked = True Then
                        SaveFileDialog1.ShowDialog()
                        If SaveFileDialog1.FileName <> Nothing Then
                            makeNexus(otumat1, SaveFileDialog1.FileName)
                        End If
                    End If
                End If
                RichTextBox1.Text = texto
                
            Next

            
        End If
        Me.Close()






    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        ListBox1.Items.Clear()
        For i = 0 To Form1.GridView1.Rows.Count - 1
            ListBox1.Items.Add(Form1.GridView1.Item(0, i).Value)
        Next i

    End Sub
    Function makefasta(ByVal otumat1 As String(,), ByVal filename As String)
        Dim filex As IO.TextWriter

        filex = New IO.StreamWriter(filename)
        Dim c As String = Form1.TextBox13.Text

        For i = 1 To otumat1.GetLength(0) - 1
            filex.WriteLine(">" & otumat1(i, 0) & c & otumat1(i, 1))
        Next
        filex.Close()
    End Function
    Function makeNexus(ByVal otumat1 As String(,), ByVal filename As String)
        Dim max As Integer
        For i = 1 To otumat1.GetLength(0) - 1
            If otumat1(i, 0).Length > max Then
                max = otumat1(i, 0).Length

            End If
        Next
        Dim filex As IO.TextWriter

        filex = New IO.StreamWriter(filename)
        Dim c As String = Form1.TextBox13.Text

        filex.WriteLine("#nexus" & c & c & "BEGIN Taxa;" & c & "DIMENSIONS ntax=" & otumat1.GetLength(0) - 1 & ";" & c & "TAXLABELS" & c)
        For i = 1 To otumat1.GetLength(0) - 1
            filex.WriteLine(otumat1(i, 0))
        Next
        filex.WriteLine(";" & c & "END; [Taxa]")
        filex.WriteLine("BEGIN CHARACTERS;" & c & "DIMENSIONS NCHAR=" & otumat1(1, 1).Length & ";")
        filex.WriteLine("FORMAT" & c & "DATATYPE=DNA" & c & "labels=No" & c & "MISSING=?" & c & "GAP=-" & c & ";" & c & "MATRIX")
        For i = 1 To otumat1.GetLength(0) - 1
            Dim txt As String
            txt = otumat1(i, 0)
            Dim n As Integer = otumat1(i, 0).Length + 1
            For x = n To max + 1
                txt = txt & " "
            Next
            filex.WriteLine(txt & otumat1(i, 1))
        Next
        filex.WriteLine(";" & c & "END; [Characters]")
        filex.Close()


    End Function
    Function makephylip(ByVal otumat1 As String(,), ByVal filename As String) As String
        Dim filex As IO.TextWriter

        filex = New IO.StreamWriter(filename)
        Dim c As String = Form1.TextBox13.Text
        filex.WriteLine(otumat1.GetLength(0) - 1 & " " & otumat1(1, 1).Length)
        For i = 1 To otumat1.GetLength(0) - 1
            Dim name As String = otumat1(i, 0)
            If otumat1(i, 0).Length < 10 Then
                For j = 1 To 10 - otumat1(i, 0).Length
                    name = name & " "
                Next
            End If
            filex.WriteLine(name & otumat1(i, 1))

        Next
        filex.Close()
    End Function
 

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.PlainText)
    End Sub

  

    Private Sub CheckBox4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetItemChecked(i, True)
            Next
        Else
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub
End Class
