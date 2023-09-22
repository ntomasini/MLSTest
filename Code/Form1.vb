Imports System.IO
Imports System.Math
Imports System.Web
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class Form1
    'Class variable declaration
    Public locusname As String()
    Private hethand As Integer = 0 'the default mode of handle heterozygosities, 0=Average states
    Private pdis As Boolean = True
    Private textito As String
    Private iHeight, iwidth As Double 'parameters for windows size
    Private externo As Boolean = False
    Private nameofproject As String
    Private nlocmin As Integer
    Private strees As Boolean
    Private stopp As Boolean
    Private start As Boolean = False
    Private inidir As String
    Private bs As New BindingSource
    Private tests As String()
    Private numerodereplicas As Integer
    Private changesmade As Boolean
    Private changesmade1 As Boolean



    Public Property hethandpp() As Integer
        Get
            Return hethand
        End Get
        Set(ByVal cadena1 As Integer)
            hethand = cadena1
        End Set
    End Property
    Public Property pdison() As Boolean
        Get
            Return pdis
        End Get
        Set(ByVal cadena1 As Boolean)
            pdis = cadena1
        End Set
    End Property

    'Data Menu
    Private Sub NewProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewProjectToolStripMenuItem.Click
        If GridView1.RowCount <> 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("Do you want to save the current project?", MsgBoxStyle.YesNoCancel)
            If aa = MsgBoxResult.Yes Then
                save()
                nameofproject = Nothing
                Me.Text = "MLSTest - New MLSTest Project"
                clearseqs()
                DataGridView6.Rows.Clear()
                loadseqs()
            ElseIf aa = MsgBoxResult.No Then
                nameofproject = Nothing
                clearseqs()
                DataGridView6.Rows.Clear()
                Me.Text = "MLSTest - New MLSTest Project"
                loadseqs()
            End If

        End If

    End Sub 'button make a new project
    Private Sub OpenProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenProjectToolStripMenuItem.Click
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "MLSTest Project(*.mls)|*.mls"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName <> "" Then
           
            DataGridView6.Rows.Clear()
            clearseqs()
            leerproject(OpenFileDialog1.FileName)
            nameofproject = OpenFileDialog1.FileName
            Me.Text = "MLSTest - " & OpenFileDialog1.SafeFileName
            TabControl1.SelectTab(TabPage1)

        End If
    End Sub 'button open a previous project
    Private Sub SaveProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveProjectToolStripMenuItem.Click
        save()
    End Sub ' Save project button
    Private Sub SaveProjectAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveProjectAsToolStripMenuItem.Click
        nameofproject = Nothing
        save()
    End Sub ' button Save project as...
    Private Sub CargarSecuenciasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadSequencesToolStripMenuItem.Click
        loadseqs()
    End Sub 'button Load Sequences
    Private Sub ViewDataFilesToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ViewDataFilesToolStripMenuItem1.Click
        TabControl1.SelectTab(TabPage1)
    End Sub 'view Data Files button
    Private Sub delunfix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles delunfix.Click
        Dim i As Integer
        Do While i < GridView1.RowCount
            If GridView1.Item(2, i).Value = False Then

                GridView1.Rows.RemoveAt(i)
                i = i - 1
            End If
            i = i + 1
        Loop
        If GridView1.RowCount = 0 Then
            clearseqs()
        End If
    End Sub 'Delete unselected loci
    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        If GridView1.RowCount > 0 Then
            If ToolStripMenuItem10.Text = "Select all Loci" Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = False
                Next
                ToolStripMenuItem10.Text = "Select all Loci"
            End If
        End If
    End Sub 'Select/unselect all loci
    Private Sub Clearseqsmi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clearseqsmi.Click
        clearseqs()
    End Sub 'Delete all sequences button
    Private Sub CloseTreeWindowsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseTreeWindowsToolStripMenuItem.Click
        Dim a As Integer = Application.OpenForms.Count - 1
        For i = 1 To a
            Application.OpenForms.Item(1).Close()
        Next

    End Sub 'Close tree windows
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        If GridView1.RowCount <> 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("Do you want to save the project before quit?", MsgBoxStyle.YesNoCancel)
            If aa = MsgBoxResult.No Then
                Me.Close()
            ElseIf aa = MsgBoxResult.Yes Then
                save()
            End If
        Else
            Me.Close()
        End If
    End Sub 'Exit button
    Private Sub GridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellEndEdit
        ToolStripMenuItem6.Enabled = False
        CalculateBootstrapImprovementToolStripMenuItem.Enabled = False
    End Sub

    ''   subs of Data menu

    Sub loadseqs()
        TabControl1.SelectTab(TabPage1)
        DataGridView1.Rows.Clear()
        OpenFileDialog1.Multiselect = True


        OpenFileDialog1.Reset()
        OpenFileDialog1.InitialDirectory = inidir

        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Title = "Load Sequences in Fasta - Multiselection is allowed"
        OpenFileDialog1.ShowDialog()
        OpenFileDialog1.Title = Nothing
        If OpenFileDialog1.FileNames.Length <> 0 Then
            inidir = OpenFileDialog1.FileNames(0).Substring(0, OpenFileDialog1.FileNames(0).Length - OpenFileDialog1.SafeFileNames(0).Length)
        End If
        If OpenFileDialog1.FileNames.Length <> 0 Then
            Dim a As Integer = GridView1.Rows.Count + 1
            GridView1.Rows.Add(OpenFileDialog1.FileNames.Length)
            GridView1.Visible = True
            DataGridView7.Visible = True
            Panel1.Visible = True
            ToolStripComboBox1.Visible = True
            Label4.Visible = True
            For i = 0 To OpenFileDialog1.FileNames.Length - 1
                Dim po As Integer = OpenFileDialog1.SafeFileNames(i).IndexOf(".")

                Try
                    GridView1.Item(0, i + a - 1).Value = OpenFileDialog1.SafeFileNames(i).Substring(0, po)
                Catch ex As Exception
                    GridView1.Item(0, i + a - 1).Value = OpenFileDialog1.SafeFileNames(i)

                End Try

                GridView1.Item(1, i + a - 1).Value = OpenFileDialog1.FileNames(i)


            Next
            If OpenFileDialog1.FileNames IsNot Nothing Then
                Dim nsq, nsq1 As Integer

                For i = 0 To GridView1.RowCount - 2
                    nsq = nseqx(GridView1.Item(1, i).Value)

                    nsq1 = nseqx(GridView1.Item(1, i + 1).Value)
                    
                    If nsq = -1 Then
                        MsgBox("Locus " & GridView1.Item(0, i).Value & " is not in Fasta Format or sequences are not correctly aligned. Sequences cannot be loaded")
                        clearseqs()
                        Exit Sub
                    End If
                    If nsq1 = -1 Then
                        MsgBox("Locus " & GridView1.Item(0, i + 1).Value & " is not in Fasta Format or sequences are not correctly aligned. Sequences cannot be loaded")
                        clearseqs()
                        Exit Sub
                    End If
                    If nsq <> nsq1 Then
                        Dim msg As Boolean
                        i = 100
                        If MsgBox("Data sets have different number of sequences. Check this carefully", MsgBoxStyle.OkOnly) = MsgBoxResult.Ok Then

                            clearseqs()
                            Exit Sub


                        End If
                    End If

                Next

            End If
            Dim otumat1(,) As String
            Dim ax(1) As String
            ax(1) = GridView1.Item(1, 0).Value
            If nseqx(ax(1)) = -1 Then
                MsgBox("Sequences not in Fasta Format. They cannot be loaded")
                clearseqs()
                Exit Sub
            End If
            otumat1 = conca(ax, 1, nseqx(GridView1.Item(1, 0).Value) + 1, TextBox13.Text, True, True)
            DataGridView7.RowCount = otumat1.GetLength(0) - 1
            Dim len As Integer = otumat1(1, 1).Length
          

            ToolStripComboBox1.Items.Add("NONE")
            For i = 1 To DataGridView7.RowCount
                DataGridView7.Item(0, i - 1).Value = otumat1(i, 0)
                ToolStripComboBox1.Items.Add(otumat1(i, 0))
                DataGridView7.Item(1, i - 1).Value = True
            Next
            ToolStripComboBox1.SelectedIndex = 0
            OpenFileDialog1.Multiselect = False
            Try
                Dim count As Integer
                For i = 0 To DataGridView7.RowCount - 1
                    If DataGridView7.Item(1, i).Value = True Then
                        count = count + 1
                    End If
                Next
                'TextBox12.Text = Module1.nseqx(GridView1.Item(1, 0).Value) + 1
                TextBox12.Text = count
                ToolStripStatusLabel1.Text = "Number of Sequences:" & TextBox12.Text
                ToolStripStatusLabel2.Text = "Outgroup:" & ToolStripComboBox1.Text
                ToolStripStatusLabel5.Text = "Number of Loci:" & GridView1.RowCount
                enable()
            Catch
                If TabControl1.SelectedIndex <> 1 Then


                    MsgBox("No se puede encontrar el archivo del Locus1 o no existen secuencias cargadas")
                    TabControl1.SelectTab(TabPage8)

                End If
            End Try
        End If
    End Sub 'Load Sequences names and path into gridviews

    Sub clearseqs()
        'Delete all loci from the gridview1 and strains in the datagridview7
        DataGridView7.Rows.Clear()
        DataGridView6.Rows.Clear()
        DataGridView4.DataSource = Nothing
        ToolStripComboBox1.Items.Clear()
        GridView1.Visible = False
        DataGridView1.Rows.Clear()
        GridView1.Rows.Clear()
        DataGridView7.Visible = False
        Label4.Visible = False
        Panel1.Visible = False
        ToolStripComboBox1.Visible = False
        Button16.Text = "Fix all"
        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""
        ToolStripStatusLabel5.Text = ""
        TabControl1.SelectTab(TabPage1)
        ToolStripComboBox2.Items.Clear()
        disable()
    End Sub 'Clear sequences
    Sub save() 'Save the project
        If GridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim texto As String
        Dim c As String = TextBox13.Text
        Dim t As Char = Char.ConvertFromUtf32(Keys.Tab)



        texto = "MLSTest Project" & c
        For i = 0 To GridView1.RowCount - 1

            texto = texto & GridView1.Item(0, i).Value & t
            texto = texto & GridView1.Item(1, i).Value & t
            texto = texto & GridView1.Item(2, i).Value & c

        Next
        texto = texto & ";" & c & c
        For i = 0 To DataGridView7.RowCount - 1
            If DataGridView7.Item(1, i).Value = True Then
                texto = texto & "+"
            Else : texto = texto & "-"
            End If

        Next
        texto = texto & c & ";" & c & c
        If DataGridView6.RowCount <> 0 Then
            For i = 0 To DataGridView6.RowCount - 1

                texto = texto & DataGridView6.Item(0, i).Value & t
                texto = texto & DataGridView6.Item(1, i).Value & t
                texto = texto & DataGridView6.Item(2, i).Value & t
                texto = texto & DataGridView6.Item(3, i).Value & c
            Next
        End If
        texto = texto & ";" & c & c & hethand & c & c & ";" & c & c


        texto = texto & ToolStripComboBox1.SelectedIndex & c & ";" & c & c
        For i = 1 To tests.Length - 1
            texto = texto & tests(i) & c
        Next
        texto = texto & ";" & c & c

        Dim xx As String

        Application.DoEvents()
        If nameofproject = Nothing Then
            SaveFileDialog1.Title = "Save Project"
            xx = "MLSTest Project" & "|" & "*." & "mls"
            SaveFileDialog1.Filter = xx
            SaveFileDialog1.DefaultExt = "mls"
            SaveFileDialog1.FileName = ""
            SaveFileDialog1.ShowDialog()
            If SaveFileDialog1.FileName <> "" Then
                Dim save As New StreamWriter(SaveFileDialog1.FileName)
                save.WriteLine(texto)
                save.Close()
                nameofproject = SaveFileDialog1.FileName

            End If
        Else
            Dim save As New StreamWriter(nameofproject)
            save.WriteLine(texto)
            save.Close()
        End If
    End Sub 'save the current project
    Sub leerproject(ByVal file As String)
        Dim lector As TextReader
        lector = New StreamReader(file)
        Dim texto As String = ""
        Dim indexout As Integer
        Dim seq(1) As String
        Dim lin As String

        Dim i As Integer = 0
        lin = lector.ReadLine
        If lin = "MLSTest Project" Then
            Dim n As Integer
            Do While lin <> ";"
                lin = lector.ReadLine()
                If lin <> ";" Then



                    If lin <> Nothing Then

                        Dim count As Integer = 3


                        Dim start, fin As Integer
                        start = 0
                        fin = 0
                        For b = 0 To 1
                            For t = start To lin.Length - 1

                                If t = lin.Length - 1 Then
                                    fin = lin.Length - 1
                                    Exit For
                                End If
                                If lin.Substring(t, 1) = Char.ConvertFromUtf32(Keys.Tab) Then
                                    fin = t

                                    Exit For
                                End If

                            Next
                            If b = 0 Then
                                GridView1.Rows.Add()
                            End If
                            GridView1.Item(b, n).Value = lin.Substring(start, fin - start)

                            start = fin + 1
                        Next

                        n = n + 1
                    End If
                End If
            Loop
            GridView1.Visible = True
            lin = lector.ReadLine

            n = 0
            checkall()

            If GridView1.RowCount = 0 Then Exit Sub
            n = 0
            Do
                lin = lector.ReadLine
                If lin <> ";" Then

                    If lin <> Nothing Then
                        For i = 0 To lin.Length - 1
                            If lin.Chars(i) = "+" Then
                                DataGridView7.Item(1, i).Value = True
                            Else
                                DataGridView7.Item(1, i).Value = False
                            End If

                        Next
                    End If
                End If
            Loop While lin <> ";"
            actseqs()
            n = 0
            lin = lector.ReadLine
            Do
                lin = lector.ReadLine
                If lin <> ";" Then



                    If lin <> Nothing Then

                        Dim count As Integer = 3


                        Dim start, fin As Integer
                        start = 0
                        fin = 0
                        For b = 0 To 3
                            For t = start To lin.Length

                                If t = lin.Length Then
                                    fin = lin.Length
                                    Exit For
                                End If
                                If lin.Substring(t, 1) = Char.ConvertFromUtf32(Keys.Tab) Then
                                    fin = t

                                    Exit For
                                End If

                            Next
                            If b = 0 Then
                                DataGridView6.Rows.Add()
                            End If
                            DataGridView6.Item(b, n).Value = lin.Substring(start, fin - start)

                            start = fin + 1
                        Next

                        n = n + 1
                    End If
                End If
            Loop While lin <> ";"
            lin = lector.ReadLine

            n = 0
            Do
                lin = lector.ReadLine
                If lin <> ";" Then

                    If lin <> Nothing Then
                        hethand = lin


                    End If
                End If
            Loop While lin <> ";"


            Do
                lin = lector.ReadLine
                If lin <> ";" Then

                    If lin <> Nothing Then

                        indexout = lin


                    End If
                End If
            Loop While lin <> ";"
            ReDim tests(0)
            Do
                lin = lector.ReadLine
                If lin <> Nothing And lin <> ";" Then
                    If IO.File.Exists(lin) Then
                        Array.Resize(tests, tests.Length + 1)
                        tests(tests.Length - 1) = lin
                        ToolStripComboBox2.Items.Add(tests(tests.Length - 1).Remove(0, tests(tests.Length - 1).LastIndexOf("\"c) + 1))
                    End If
                End If

            Loop While (lin <> ";")

            ToolStripComboBox1.SelectedIndex = indexout
            ToolStripStatusLabel5.Text = "Number of Loci:" & GridView1.RowCount
            lector.Close()
            If GridView1.RowCount <> 0 Then
                enable()
            End If
        End If

    End Sub 'read a previous project
    Sub actseqs()

        DataGridView6.Rows.Clear()
        ToolStripComboBox1.Items.Clear()

        Dim count As Integer
        ToolStripComboBox1.Items.Add("none")
        For i = 0 To DataGridView7.RowCount - 1
            If DataGridView7.Item(1, i).Value = True Then
                count = count + 1
                ToolStripComboBox1.Items.Add(DataGridView7.Item(0, i).Value)

            End If
        Next
        If ToolStripComboBox1.Items.Contains(ToolStripComboBox1.Text) = True Then
            ToolStripComboBox1.SelectedItem = ToolStripComboBox1.Text
        Else
            ToolStripComboBox1.SelectedIndex = 0


        End If
        'TextBox12.Text = Module1.nseqx(GridView1.Item(1, 0).Value) + 1
        TextBox12.Text = count
        ToolStripStatusLabel1.Text = "Number of Sequences:" & TextBox12.Text

    End Sub 'Correct the sequences showed in the outgroup combobox when any of them is selected or deselected
    Sub disable()
        AlignmentToolStripMenuItem.Enabled = False
        DistanceMatrixToolStripMenuItem.Enabled = False
        AllelicProfilesToolStripMenuItem.Enabled = False
        DTUToTestToolStripMenuItem.Enabled = False
        TestAllCombinationsToolStripMenuItem.Enabled = False
        AddOneToolStripMenuItem.Enabled = False
        DeleteOneToolStripMenuItem.Enabled = False
        TreeToolStripMenuItem.Enabled = False
        HaplotipesToolStripMenuItem.Enabled = False
        CongruenceTestsToolStripMenuItem.Enabled = False
        ToolStripMenuItem9.Enabled = False
        ToolStripMenuItem3.Enabled = False
        CalcToolStripMenuItem.Enabled = False
        SelectHighMeanBootstrapTreesToolStripMenuItem.Enabled = False
        ToolStripMenuItem6.Enabled = False
        CalculateBootstrapImprovementToolStripMenuItem.Enabled = False

    End Sub 'disable menu options when are not available
    Sub blockmenusytab()
        MenuStrip1.Enabled = False
        TabControl1.Enabled = False
    End Sub
    Private Sub TtMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TtMenuItem2.Click
        If GridView1.SelectedCells.Count <> 0 Then
            Dim cellpos As Integer = GridView1.SelectedCells.Item(0).RowIndex
            GridView1.Rows.RemoveAt(cellpos)
            ToolStripStatusLabel5.Text = "Number of Loci:" & GridView1.RowCount
            DataGridView1.Rows.Clear()
            If GridView1.RowCount = 0 Then
                clearseqs()
            End If
        End If
    End Sub 'Delete selected loci
    Sub enable()
        AlignmentToolStripMenuItem.Enabled = True
        DistanceMatrixToolStripMenuItem.Enabled = True
        AllelicProfilesToolStripMenuItem.Enabled = True
        DTUToTestToolStripMenuItem.Enabled = True
        TestAllCombinationsToolStripMenuItem.Enabled = True
        AddOneToolStripMenuItem.Enabled = True
        ToolStripMenuItem9.Enabled = True
        DeleteOneToolStripMenuItem.Enabled = True
        TreeToolStripMenuItem.Enabled = True
        HaplotipesToolStripMenuItem.Enabled = True
        CongruenceTestsToolStripMenuItem.Enabled = True
        ToolStripMenuItem3.Enabled = True
        CalcToolStripMenuItem.Enabled = True
        SelectHighMeanBootstrapTreesToolStripMenuItem.Enabled = True

    End Sub 'enable menu options when are  available
    Sub checkall()
        TabControl1.SelectTab(TabPage1)
        DataGridView1.Rows.Clear()





        DataGridView7.Visible = True
        Panel1.Visible = True
        ToolStripComboBox1.Visible = True
        Label4.Visible = True

        Dim nsq, nsq1 As Integer

        For i = 0 To GridView1.RowCount - 2
            nsq = nseqx(GridView1.Item(1, i).Value)

            nsq1 = nseqx(GridView1.Item(1, i + 1).Value)
            If nsq = -1 Then
                MsgBox("Locus " & GridView1.Item(0, i).Value & " is not in Fasta Format. Sequences cannot be loaded")
                clearseqs()
                Exit Sub
            End If
            If nsq1 = -1 Then
                MsgBox("Locus " & GridView1.Item(0, i + 1).Value & " is not in Fasta Format. Sequences cannot be loaded")
                clearseqs()
                Exit Sub
            End If
            If nsq <> nsq1 Then
                Dim msg As Boolean
                i = 100
                If MsgBox("Data sets have different number of sequences. Check this carefully", MsgBoxStyle.OkOnly) = MsgBoxResult.Ok Then

                    clearseqs()
                    Exit Sub
                End If
            End If

        Next


        Dim otumat1(,) As String
        Dim ax(1) As String
        ax(1) = GridView1.Item(1, 0).Value
        If nseqx(ax(1)) = -1 Then
            MsgBox("Sequences not in Fasta Format. They cannot be loaded")
            clearseqs()
            Exit Sub
        End If
        otumat1 = conca(ax, 1, nseqx(GridView1.Item(1, 0).Value) + 1, TextBox13.Text, True, True)
        DataGridView7.RowCount = otumat1.GetLength(0) - 1
        ToolStripComboBox1.Items.Add("NONE")
        For i = 1 To DataGridView7.RowCount
            DataGridView7.Item(0, i - 1).Value = otumat1(i, 0)
            ToolStripComboBox1.Items.Add(otumat1(i, 0))
            DataGridView7.Item(1, i - 1).Value = True
        Next
        ToolStripComboBox1.SelectedIndex = 0

        Try
            Dim count As Integer
            For i = 0 To DataGridView7.RowCount - 1
                If DataGridView7.Item(1, i).Value = True Then
                    count = count + 1
                End If
            Next
            'TextBox12.Text = Module1.nseqx(GridView1.Item(1, 0).Value) + 1
            TextBox12.Text = count
            ToolStripStatusLabel1.Text = "Number of Sequences:" & TextBox12.Text
            ToolStripStatusLabel2.Text = "Outgroup:" & ToolStripComboBox1.Text
        Catch
            If TabControl1.SelectedIndex <> 1 Then


                MsgBox("No se puede encontrar el archivo del Locus1 o no existen secuencias cargadas")
                TabControl1.SelectTab(TabPage8)

            End If
        End Try

    End Sub 'check if the fasta files contain the same number of sequences
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        ToolStripStatusLabel2.Text = "Outgroup:" & ToolStripComboBox1.Text
    End Sub

    Private Sub DataGridView7_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView7.CellBeginEdit
        DataGridView7.EndEdit()
        actseqs()
    End Sub

    Private Sub DataGridView7_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView7.ColumnHeaderMouseClick
        If DataGridView7.Item(1, 0).Value = False Then
            If e.ColumnIndex = 1 Then


                For i = 0 To DataGridView7.RowCount - 1
                    DataGridView7.Item(1, i).Value = True
                Next

            End If
            DataGridView7.EndEdit()
            actseqs()
        Else
            If e.ColumnIndex = 1 Then


                For i = 0 To DataGridView7.RowCount - 1
                    DataGridView7.Item(1, i).Value = False
                Next

            End If
            DataGridView7.EndEdit()
            actseqs()
        End If

    End Sub 'to select an outgroup taxon
    Private Sub DataGridView7_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView7.MouseLeave
        DataGridView7.EndEdit()
        'actseqs()
    End Sub
    Private Sub MenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.Click
        GridView1.EndEdit()
    End Sub
    Private Sub DataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataToolStripMenuItem.Click
        Dim status As Boolean
        For i = 0 To GridView1.Rows.Count - 1

            If GridView1.Item(2, i).Value = True Then
                status = True
            End If
        Next
        If status = True Then
            ToolStripMenuItem10.Text = "Unselect all Loci"
        Else
            ToolStripMenuItem10.Text = "Select all Loci"
        End If
    End Sub

    ' Alignment Menu
    Private Sub VisorDeAlineamientosToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles VisorDeAlineamientosToolStripMenuItem.Click
        ListBox1.Items.Clear()
        For i = 0 To GridView1.Rows.Count - 1
            ListBox1.Items.Add(GridView1.Item(0, i).Value)
        Next i

        TabControl1.SelectedTab = TabPage5

    End Sub 'Shows the alignment viewer
    Private Sub ExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        Dialog1.Show()
    End Sub 'Shows the alignment export Dialog1
    Private Sub ViewerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewerToolStripMenuItem.Click
        ListBox1.Items.Clear()
        For i = 0 To GridView1.Rows.Count - 1
            ListBox1.Items.Add(GridView1.Item(0, i).Value)
        Next i
        ListBox1.SelectedIndex = 0
        DataGridView7.EndEdit()
        TabControl1.SelectedTab = TabPage5
        viewer()
    End Sub

    ''Alignment window
    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        viewer()
    End Sub 'select an locus to view the alignment
    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick, ListBox7.DoubleClick
        viewer()
    End Sub 'select an locus to view the alignment
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        viewer()
    End Sub 'select an locus to view the alignment
    Private Sub dataGridView2_EditingControlShowing(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView2.EditingControlShowing
        Dim validar As TextBox = CType(e.Control, TextBox)
        AddHandler validar.KeyPress, AddressOf validar_Keypress
    End Sub
    Private Sub validar_Keypress( _
        ByVal sender As Object, _
        ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If RadioButton4.Checked = False And RadioButton6.Checked = False Then
            Exit Sub
        End If



        DataGridView2.CurrentCell.Style.BackColor = Color.White

        Dim c As Char = Char.ToUpperInvariant(e.KeyChar)


        Select Case c
            Case "G"
                DataGridView2.CurrentCell.Style.BackColor = Color.Purple
            Case "A"
                DataGridView2.CurrentCell.Style.BackColor = Color.Green
            Case "T"
                DataGridView2.CurrentCell.Style.BackColor = Color.Red
            Case "C"
                DataGridView2.CurrentCell.Style.BackColor = Color.SkyBlue
            Case "R"
            Case "S"
            Case "M"
            Case "K"
            Case "W"
            Case "Y"
            Case "-"
            Case "N"
            Case "B"
            Case "V"
            Case "H"
            Case "D"

            Case Else


                e.KeyChar = Chr(0)

        End Select
        If DataGridView2.CurrentCell.Value <> c Then
            changesmade = True
        End If
        DataGridView2.CancelEdit()


    End Sub 'to control base replacement
    Private Sub DataGridView2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit

        If DataGridView2.CurrentCell.Value = Nothing Then
            DataGridView2.CurrentCell.Value = "-"
        Else
            DataGridView2.CurrentCell.Value = DataGridView2.CurrentCell.Value.ToString.ToUpper()
        End If
        If changesmade = True Then
            Dim c As Char = DataGridView2.CurrentCell.Value
            If c <> Nothing Then
                Dim r As Integer = DataGridView2.CurrentCell.RowIndex
                Dim col As Integer
                If ListBox7.Items.Count > 0 Then
                    col = ListBox7.Items(DataGridView2.CurrentCell.ColumnIndex)
                Else
                    col = DataGridView2.CurrentCell.ColumnIndex + 1
                End If
                alignmf(r + 1, 1) = alignmf(r + 1, 1).Remove(col - 1, 1)
                alignmf(r + 1, 1) = alignmf(r + 1, 1).Insert(col - 1, c)
                changesmade1 = True
                savemodifica()
                DataGridView2.CurrentCell.ToolTipText = "Click on another locus to save changes"
            End If
        End If
        changesmade = False

    End Sub

    '' subs of alignment 
    Private alignmf(,) As String = Nothing
    Private l1index As Integer
    Sub viewer()

        savemodifica()

        ListBox7.Items.Clear()
        Dim locus(GridView1.Rows.Count - 1) As String
        DataGridView2.Visible = True

        Dim i As Integer
        For i = 0 To GridView1.Rows.Count - 1
            locus(i) = GridView1.Item(1, i).Value
        Next

        Dim a(ListBox1.SelectedItems.Count) As String
        For i = 1 To ListBox1.SelectedItems.Count
            a(i) = locus(ListBox1.SelectedIndices(i - 1))
        Next

        alignmf = Nothing
        Dim alignm(,) As String
        Dim frame As Integer = NumericUpDown1.Value
        Dim pos As Integer = 0
        alignmf = conca(a, a.Length - 1, TextBox12.Text, TextBox13.Text, True, False)
        alignm = alignmf.Clone
        nseq = alignm.GetLength(0)
        lseq = alignm(1, 1).Length - 1
        TextBox9.Text = lseq + 1
        ' Dim alignm1(lseq) As String
        Dim seq1, seq2 As String

        If RadioButton6.Checked = True Then 'delete  constant sites if this option is selected
            Dim consta As Boolean
            i = 0
            Do While i < lseq + 1
                consta = True
                For j = 1 To nseq - 2
                    seq1 = alignm(j, 1)
                    seq2 = alignm(j + 1, 1)
                    If seq1.Substring(i, 1) <> seq2.Substring(i, 1) Then
                        consta = False
                        j = nseq
                    End If
                Next
                If consta = True Then

                    For t = 1 To nseq - 1
                        If i <> 0 Then
                            alignm.SetValue(alignm(t, 1).ToString.Substring(0, i) & alignm.GetValue(t, 1).ToString.Substring(i + 1, alignm.GetValue(t, 1).ToString.Length - (i + 1)), t, 1)
                        Else
                            alignm.SetValue(alignm(t, 1).ToString.Substring(1, alignm.GetValue(t, 1).ToString.Length - 1), t, 1)
                        End If



                    Next
                    i = i - 1
                    lseq = lseq - 1
                    pos = pos + 1
                Else
                    ListBox7.Items.Add(pos + 1)
                    pos = pos + 1
                End If



                i = i + 1
            Loop
        End If

        '----
        If RadioButton7.Checked = True Or RadioButton8.Checked = True Then 'protein translation
            Dim consta As Boolean
            pos = frame - 1
            For i = 1 To nseq - 1
                alignm(i, 1) = alignm(i, 1).Substring(frame - 1, alignm(i, 1).Length - (frame - 1))

            Next
            lseq = alignm(1, 1).Length - 1
            i = 0
            Do While i < lseq - 2
                consta = True
                If RadioButton7.Checked = True Then
                    For j = 1 To nseq - 2
                        seq1 = alignm(j, 1)
                        seq2 = alignm(j + 1, 1)
                        If seq1.Substring(i, 3) <> seq2.Substring(i, 3) Then
                            consta = False
                            j = nseq
                        End If
                    Next
                ElseIf RadioButton8.Checked = True Then
                    For j = 1 To nseq - 2
                        seq1 = alignm(j, 1)
                        seq2 = alignm(j + 1, 1)
                        If translate(seq1.Substring(i, 3)) <> translate(seq2.Substring(i, 3)) Then
                            consta = False
                            j = nseq
                        End If
                    Next

                End If
                If consta = True Then

                    For t = 1 To nseq - 1
                        If i <> 0 Then
                            alignm(t, 1) = alignm(t, 1).Substring(0, i) & alignm(t, 1).Substring(i + 3, alignm(t, 1).Length - (i + 3))
                        Else

                            alignm(t, 1) = alignm(t, 1).Substring(3, alignm(t, 1).Length - 3)


                        End If


                    Next
                    i = i - 3
                    lseq = lseq - 3
                    pos = pos + 3
                Else
                    If CheckBox7.Checked = False Then
                        For t = 1 To nseq - 1
                            If i <> 0 Then
                                alignm(t, 1) = alignm(t, 1).Substring(0, i + 3) & translate(alignm(t, 1).Substring(i, 3)) & alignm(t, 1).Substring(i + 3, alignm(t, 1).Length - (i + 3))
                            Else

                                alignm(t, 1) = alignm(t, 1).Substring(0, 3) & translate(alignm(t, 1).Substring(0, 3)) & alignm(t, 1).Substring(3, alignm(t, 1).Length - 3)


                            End If


                        Next
                        ListBox7.Items.Add(pos + 1)
                        ListBox7.Items.Add(pos + 2)
                        ListBox7.Items.Add(pos + 3)
                        ListBox7.Items.Add(pos + 1 & "-" & pos + 3)
                        pos = pos + 3
                    Else
                        For t = 1 To nseq - 1
                            If i <> 0 Then
                                alignm(t, 1) = alignm(t, 1).Substring(0, i) & translate(alignm(t, 1).Substring(i, 3)) & alignm(t, 1).Substring(i + 3, alignm(t, 1).Length - (i + 3))
                            Else

                                alignm(t, 1) = translate(alignm(t, 1).Substring(0, 3)) & alignm(t, 1).Substring(3, alignm(t, 1).Length - 3)


                            End If


                        Next
                        i = i - 3

                        ListBox7.Items.Add(pos + 1 & "-" & pos + 3)
                        pos = pos + 3
                    End If
                    i = i + 1

                End If



                i = i + 3
                lseq = alignm(1, 1).Length - 1
            Loop
            lseq = lseq - (lseq - i + 1)
        End If



        DataGridView2.ColumnCount = 0
        For i = 0 To lseq
            DataGridView2.Columns.Add(i, "")
            DataGridView2.Columns(i).FillWeight = 1
        Next
        DataGridView2.RowCount = nseq - 1

        For i = 1 To nseq - 1 'writes the alignment on the screen in a gridview 
            For j = 0 To lseq + 1
                If j = 0 Then

                    DataGridView2.Rows(i - 1).HeaderCell.Value = alignm(i, 0)
                Else

                    DataGridView2.Item(j - 1, i - 1).Value = alignm(i, 1).ToString.Substring(j - 1, 1)
                    DataGridView2.Columns(j - 1).Width = 14
                    If RadioButton7.Checked = False And RadioButton8.Checked = False Then
                        Select Case alignm(i, 1).ToString.Substring(j - 1, 1)
                            Case "G"
                                DataGridView2.Item(j - 1, i - 1).Style.BackColor = Color.Purple
                            Case "A"
                                DataGridView2.Item(j - 1, i - 1).Style.BackColor = Color.Green
                            Case "T"
                                DataGridView2.Item(j - 1, i - 1).Style.BackColor = Color.Red
                            Case "C"
                                DataGridView2.Item(j - 1, i - 1).Style.BackColor = Color.SkyBlue
                            Case Else
                                DataGridView2.Item(j - 1, i - 1).Style.BackColor = Color.White
                        End Select
                    Else
                    End If
                End If

            Next
        Next
        If RadioButton7.Checked = True Or RadioButton8.Checked = True Then
            If CheckBox7.Checked = False Then
                For i = 1 To nseq - 1
                    For j = 4 To lseq + 1
                        DataGridView2.Item(j - 1, i - 1).Style.BackColor = Color.Aquamarine

                        j = j + 3
                    Next
                Next
            End If
        End If
        l1index = ListBox1.SelectedIndex
        DataGridView2.Focus()
    End Sub 'shows the alignment on the screen
    Sub position()
        If DataGridView2.SelectedCells.Count <> 0 Then
            Dim cellpos As Integer = DataGridView2.SelectedCells.Item(0).ColumnIndex

            If RadioButton4.Checked = True Then
                TextBox14.Text = "pos:" & cellpos + 1 & " |Length:" & TextBox9.Text
            Else
                Try
                    TextBox14.Text = "pos:" & ListBox7.Items(cellpos) & " |Length:" & TextBox9.Text
                Catch
                End Try
            End If
        End If
    End Sub 'determine the cursor position in the alignment 
    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        position()
    End Sub
    Private Sub DataGridView2_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellEnter
        position()
    End Sub
    Private Sub DataGridView2_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.VisibleChanged
        position()

    End Sub
    Private Sub RadioButton4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.Click, RadioButton6.Click, RadioButton7.Click, RadioButton8.Click
        viewer()
    End Sub
    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        viewer()
    End Sub
    Private Sub NumericUpDown1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.Click
        If RadioButton7.Checked Or RadioButton8.Checked = True Then
            viewer()
        End If
    End Sub
    Sub savemodifica()
        If changesmade1 = True Then
            If alignmf IsNot Nothing Then
                Dim texto, path As String
                Dim a(1) As String

                path = GridView1.Item(1, l1index).Value
                a(1) = path
                Dim otu(,) As String = conca(a, 1, DataGridView7.RowCount, TextBox13.Text, False, True)
                Dim pos As Integer = 0
                For i = 0 To DataGridView7.RowCount - 1
                    If DataGridView7.Item(1, i).Value = True Then
                        otu(i + 1, 1) = alignmf(pos + 1, 1)
                        pos = pos + 1
                    End If

                Next

                Dialog1.makefasta(otu, path)
                
            End If
        End If
        changesmade1 = False
    End Sub 'Saves modified alignments

    'Distance Matrix Menu
    Private Sub ViewAllelicProfilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewAllelicProfilesToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci?", MsgBoxStyle.YesNo, "Distance Matrix")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                Button16.Text = "Unfix all"
            Else
                Exit Sub
            End If
        End If

        ProgressBar1.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        distancematperfal()
        For Each column As DataGridViewColumn In DataGridView9.Columns
            column.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
        TabControl1.SelectTab(TabPage2)
        restore()
    End Sub 'Shows the distance Matrix of allelic profiles
    Private Sub ViewDistanceMatrixToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewDistanceMatrixToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci?", MsgBoxStyle.YesNo, "Distance Matrix")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                Button16.Text = "Unfix all"
            Else
                Exit Sub
            End If
        End If

        ProgressBar1.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        distancemat()
        For Each column As DataGridViewColumn In DataGridView9.Columns
            column.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
        TabControl1.SelectTab(TabPage2)
        restore()

    End Sub 'shows the distance matrix between sequences
    Private Sub ExportDistanceMatrixToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportDistanceMatrixToolStripMenuItem.Click
        If DataGridView9.RowCount <> 0 Then
            Dim texto As String
            texto = "    " & DataGridView9.RowCount & vbNewLine
            For i = 0 To DataGridView9.RowCount - 1
                Dim name As String = DataGridView9.Rows(i).HeaderCell.Value
                If name.Length < 12 Then
                    Do While name.Length < 12
                        name = name & " "
                    Loop
                Else
                    name = name.Substring(0, 11) & " "
                End If
                texto = texto & name '& Char.ConvertFromUtf32(Keys.Tab)
                For j = 0 To DataGridView9.ColumnCount - 1

                    texto = texto & phylformat(DataGridView9.Item(j, i).Value) & "  " 'Char.ConvertFromUtf32(Keys.Tab)


                Next
                texto = texto & vbNewLine
            Next


            SaveFileDialog1.FileName = Nothing
            SaveFileDialog1.Title = "Save Distance Matrix..."
            SaveFileDialog1.ShowDialog()
            If SaveFileDialog1.FileName <> Nothing Then
                Dim save As New StreamWriter(SaveFileDialog1.FileName)
                save.WriteLine(texto)
                save.Close()

            End If

        End If

    End Sub 'Exports distance matrix
    Private Function phylformat(ByVal n As Single) As String
        Dim a As Integer = n.ToString.Length
        If a = 6 Then
            Return n.ToString
        ElseIf a > 6 Then
            Return Math.Round(n, 4)
        ElseIf n = 0 Then
e1:         Return "0.0000"
        Else
            Dim c As Single = Math.Round(n, 4)
            Dim str As String = c
            If c = 0 Then
                GoTo e1
            Else
                For m = a + 1 To 6
                    str = str & "0"
                Next
                Return str
            End If

        End If
        

    End Function
    Private Sub DistanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DistanceToolStripMenuItem.Click
        Dim a As New Dialog2 With {.Text = "Heterozygosities", .page = 2}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            Dim x As String = hethand.ToString
            hethand = a._hethand
            If x <> hethand Then
                If TabControl1.SelectedTab Is TabPage2 Then
                    Dim result As MsgBoxResult
                    result = MsgBox("do you want to refresh distance matrix viewer?", MsgBoxStyle.OkCancel, "Handling heterozygosities")
                    If result = MsgBoxResult.Ok Then
                        ViewDistanceMatrixToolStripMenuItem_Click(Nothing, Nothing)
                    End If
                End If
            End If


        End If
    End Sub 'Shows a dialog to select the way to handle heterozygoses

    Sub distancematperfal()
        Dim locus(GridView1.Rows.Count) As String
        Dim i As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        textito = Nothing
        Dim arch As String = Nothing
        ProgressBar2.Value = 0


        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0



        i = 1


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value

        Next


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)


                fi = fi + 1
            End If
        Next

        Dim perfiles(,) = Perfal.perf(a, TextBox12.Text, locusname, True)

        Dim otumat1(,) As String
        otumat1 = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, True, False)

        nseq = perfiles.GetLength(0) - 2

        lseq = perfiles.GetLength(1) - 2



        Dim seqs(1) As String
        Dim nseq1 As Integer = nseq + 1
        Dim distancias(nseq1, nseq1) As Single

        Dim bootstrapn() As Integer


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
        writedist(otumat1, distancias)


    End Sub 'Calculates distances between allelic profiles and make the matrix
    Sub distancemat()
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        ProgressBar1.Value = 0
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next

        Dim arch As String = Nothing



        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0




        i = 1


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)

                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If

                fi = fi + 1
            End If
        Next
        otumat1 = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        nseq = otumat1.GetLength(0) - 2

        lseq = otumat1(1, 1).ToString.Length

        otumat1 = reductor(otumat1)
        lseqred = otumat1.GetValue(1, 1).ToString.Length





        Dim seqs(1) As String
        Dim nseq1 As Integer = nseq + 1
        Dim distancias(nseq1, nseq1) As Double
        Dim j As Integer
        Dim bootstrapn() As Integer
        Dim dist As medirdistancia

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
        'Dim texto As String
        'For i = 1 To distancias.GetLength(0) - 1
        'For j = 1 To distancias.GetLength(0) - 1
        'texto = texto & " " & distancias(i, j)
        'Next
        'texto = texto & Form1.TextBox13.Text
        'Next
        writedist(otumat1, distancias)
    End Sub 'Calculates distances between sequences
    Sub writedist(ByVal otumat1, ByVal distancias)
        If DataGridView9.ColumnCount <> 0 Then
            Do While DataGridView9.ColumnCount <> 0
                DataGridView9.Columns.RemoveAt(0)
            Loop
        End If

        DataGridView9.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders
        If otumat1.getlength(0) - 1 > 665 Then
            GoTo lx
        End If
        DataGridView9.ColumnCount = TextBox12.Text
        DataGridView9.RowCount = TextBox12.Text

        For i = 0 To DataGridView9.ColumnCount - 1
            DataGridView9.Rows(i).HeaderCell.Value = otumat1(i + 1, 0)

            DataGridView9.Columns(i).HeaderText = otumat1(i + 1, 0)
            DataGridView9.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            For j = 0 To DataGridView9.RowCount - 1
                DataGridView9.Item(i, j).Value = Math.Round(distancias(j + 1, i + 1), 4)

            Next
        Next
        GoTo l3


lx:


        Dim txt2 As String = Nothing
        For x = 1 To otumat1.getlength(0) - 1

            DataGridView9.Columns.Add(x, otumat1(x, 0))
            DataGridView9.Columns(x - 1).FillWeight = 1

        Next
        DataGridView9.RowCount = TextBox12.Text
        For i = 0 To DataGridView9.ColumnCount - 1
            DataGridView9.Rows(i).HeaderCell.Value = otumat1(i + 1, 0)

            DataGridView9.Columns(i).HeaderText = otumat1(i + 1, 0)
            DataGridView9.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            For j = 0 To DataGridView9.RowCount - 1
                DataGridView9.Item(i, j).Value = Math.Round(distancias(j + 1, i + 1), 4)

            Next
        Next
l3:

    End Sub 'write the distances in gridview on the screen

    'Allelic Profiles Menu
    Private Sub MakeAllelicProfilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MakeAllelicProfilesToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Or DataGridView1.RowCount = 1 Then
            If DataGridView1.RowCount = 1 Then
                If DataGridView1.Item(0, 0).Value = Nothing Then
                    profiles()
                End If
            Else
                profiles()
            End If

        Else
            Dim aa As MsgBoxResult
            aa = MsgBox("There are allelic profiles saved into the memory. If you want to discard them and calculate Allelic Profiles again click on Yes, elsewhere, if you want to see saved allelic profiles then click on NO", MsgBoxStyle.YesNoCancel, "MLSTest")
            If aa = MsgBoxResult.Yes Then
                profiles()
            ElseIf aa = MsgBoxResult.No Then
                TabControl1.SelectTab(TabPage4)

            End If

        End If

    End Sub 'Make/view allelic Profiles button
    Private Sub UPGMAtreeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPGMAtreeToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "UPGMA-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "UPGMA with allelic profiles", .page = 1}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            CancelButton1.Visible = True
            blockmenusytab()
            Dim resethethand As Integer = hethand
            hethand = 0
            UPGMAtree(True, False, a.nrep)
            hethand = resethethand
        End If

    End Sub 'UPGMA button
    Private Sub NJtreeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NJtreeToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Neighbor Joining with allelic profiles", .page = 1}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            CancelButton1.Visible = True
            blockmenusytab()
            Dim resethethand = hethand
            hethand = 0
            Dim testsplits As Boolean
            If a.viewnmin = True Then
                Dim respuesta As MsgBoxResult
                respuesta = MsgBox("Do you want to test only selected groups?, Click NO to test all nodes into the tree", MsgBoxStyle.YesNo, "MLSTest1.0")
                If respuesta = MsgBoxResult.Yes Then
                    testsplits = True
                End If
            End If
            NJtree(True, False, False, a.nrep, a._bionj, a.viewnmin, testsplits, False)
            hethand = resethethand

        End If


    End Sub 'NJ button
    Private Sub SingleBurst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SingleBurst.Click

        Dim count As Integer
        For i = 0 To GridView1.RowCount - 1
            If GridView1.Item(2, i).Value = True Then
                count = count + 1
            End If

        Next
        If count = 0 Then count = GridView1.RowCount
        Dim dial As New Dialog2 With {.Text = "Select the group definition", .page = 0, .nloc = count - 1}
        dial.ShowDialog()

        If dial.DialogResult = Windows.Forms.DialogResult.OK Then


            If DataGridView1.RowCount = 0 Then
                profiles()
            End If

            eBurstproc(False, count - dial.gdef, False)
        End If

    End Sub 'Burst button, calls the Burst procedure
    Private Sub SeqBurst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SeqBurst.Click
        textito = Nothing
        If DataGridView1.RowCount = 0 Then
            profiles()
        End If
        eBurstproc(True, 0, False)
    End Sub 'Sequential Burst button
    Private Sub CopyToClipboardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToClipboardToolStripMenuItem.Click
        If DataGridView1.RowCount <> 0 Then
            DataGridView1.SelectAll()
            Clipboard.SetDataObject(DataGridView1.GetClipboardContent())
        End If
    End Sub 'copy allelic profiles to clipboard
    Private Sub ExportToFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToFileToolStripMenuItem.Click
        If DataGridView1.RowCount <> 0 Then
            Dim texto As String
            For i = 0 To DataGridView1.RowCount
                For j = 0 To DataGridView1.ColumnCount - 1
                    If i = 0 Then
                        texto = texto & DataGridView1.Columns(j).HeaderText & Char.ConvertFromUtf32(Keys.Tab)
                    Else
                        texto = texto & DataGridView1.Item(j, i - 1).Value & Char.ConvertFromUtf32(Keys.Tab)
                    End If

                Next
                texto = texto & TextBox13.Text
            Next


            SaveFileDialog1.FileName = Nothing
            SaveFileDialog1.Title = "Save Allelic Profiles As..."
            SaveFileDialog1.ShowDialog()
            If SaveFileDialog1.FileName <> Nothing Then
                Dim save As New StreamWriter(SaveFileDialog1.FileName)
                save.WriteLine(texto)
                save.Close()

            End If

        End If

    End Sub 'Exports the allelic profiles to a file

    ''Subs of Allelic Profiles
    Private Sub profiles()



        For x = 1 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns.Remove(x)
        Next
        Try
            DataGridView1.Columns.Remove("ST")
        Catch
        End Try
        Dim locus(GridView1.Rows.Count) As String
        Dim locusname(GridView1.Rows.Count - 1) As String

        Dim i As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
            locusname(i - 1) = GridView1.Item(0, i - 1).Value
        Next

        Dim arch As String = Nothing
        ProgressBar2.Value = 0
        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0

        Dim ch(10) As Boolean


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)



                fi = fi + 1
            End If
        Next



        i = 1

        If a Is Nothing Then
            a = locus
        End If

        Perfal.perf(a, TextBox12.Text, locusname, False)

        Dim perfil1(DataGridView1.RowCount - 4) As String
        Dim perfil2(DataGridView1.RowCount - 4) As Integer
        Dim dst As Integer = 0
        For j = 0 To DataGridView1.RowCount - 4
            For i = 1 To DataGridView1.ColumnCount - 1
                perfil1(j) = perfil1(j) & DataGridView1.Item(i, j).Value

            Next
        Next
        Dim col As Integer = DataGridView1.ColumnCount - 1

        For i = 0 To perfil1.Length - 2
            If perfil2(i) = Nothing Then
                dst = dst + 1
                perfil2(i) = dst
            End If
            For j = i + 1 To perfil1.Length - 1

                If perfil2(j) = Nothing Then
                    If perfil1(j) = perfil1(i) Then

                        perfil2(j) = perfil2(i)



                    End If

                End If

            Next


        Next
        Dim suma As Integer
        For p = 1 To nseq + 1
            Dim count As Integer = 0
            For q = 1 To nseq + 1

                If perfil2(q - 1) = p Then
                    count = count + 1
                End If

            Next
            count = count * (count - 1)
            suma = suma + count
        Next
        Dim dp As Single = 1 - (suma / ((nseq + 1) * nseq)) ' (npol(i, 1) / nsq1)
        DataGridView1.Item(DataGridView1.ColumnCount - 1, DataGridView1.RowCount - 1).Value = Math.Round(dp, 3)
        DataGridView1.Item(DataGridView1.ColumnCount - 1, DataGridView1.RowCount - 4).Value = perfil2.Max()
        Dim xa As Integer = perfil2.Max()
        For i = 0 To perfil1.Length - 2
            DataGridView1.Item(col, i).Value = perfil2(i)
        Next
        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
        TabControl1.SelectTab(TabPage4)
        Application.DoEvents()
    End Sub 'clear previous profiles and write the profiles on screen
    Private Sub eBurstproc(ByVal seq As Boolean, ByVal ndif As Integer, ByVal testnet As Boolean, Optional ByVal netsize As Integer = 4)
        TextBox6.Text = Nothing
        Dim profi(DataGridView1.Rows.Count - 5, DataGridView1.Columns.Count - 2) As Integer
        Dim otunames(DataGridView1.Rows.Count - 5, 1) As String
        For i = 0 To DataGridView1.Rows.Count - 5
            For j = 1 To DataGridView1.Columns.Count - 1

                profi(i, j - 1) = DataGridView1.Item(j, i).Value
            Next
        Next
        For i = 0 To DataGridView1.Rows.Count - 5


            otunames(i, 0) = DataGridView1.Item(0, i).Value
            otunames(i, 1) = DataGridView1.Item(DataGridView1.Columns.Count - 1, i).Value

        Next

        Dim sp As splits
        Dim gruburst As reburst
        sp = eburst.makeburstgraph(profi, otunames, seq, ndif, testnet, gruburst, netsize)

        CheckBox1.Checked = False
        If seq = True Then
            ' writefiles(sp, 0, True, False)
            Dim vs As Boolean = True
            Dim supnames(0) As String
            supnames(0) = "Group Definition"
            'Dim nwkform As New treeviewer1 With {.nwktree = textito, ._seqburst = True}
            Dim nwkform As New treeviewer2 With {.sptree = sp, .Text = "Tree Viewer", ._viewsupport = vs, ._supportnames = supnames}

            nwkform.Show()


        End If
        TabControl1.SelectTab(TabPage10)
    End Sub 'it Call BURST and shows the results
    Private Sub HiddenSLV(ByVal ndif As Integer)
        TextBox6.Text = Nothing
        Dim profi(DataGridView1.Rows.Count - 5, DataGridView1.Columns.Count - 2) As Integer
        Dim otunames(DataGridView1.Rows.Count - 5, 1) As String
        For i = 0 To DataGridView1.Rows.Count - 5
            For j = 1 To DataGridView1.Columns.Count - 1

                profi(i, j - 1) = DataGridView1.Item(j, i).Value
            Next
        Next
        For i = 0 To DataGridView1.Rows.Count - 5


            otunames(i, 0) = DataGridView1.Item(0, i).Value
            otunames(i, 1) = DataGridView1.Item(DataGridView1.Columns.Count - 1, i).Value

        Next


        Dim gruburst As reburst
        eburst.makeburstgraph(profi, otunames, False, ndif, False, gruburst)
        Dim xx(gruburst.gs.Length - 1, gruburst.gs.Length - 1) As Integer
        For j = 0 To gruburst.gs.Length - 3

            For k = j + 1 To gruburst.gs.Length - 2
                xx(j, k) = testforclusep(j, k, gruburst)
            Next
        Next
        CheckBox1.Checked = False

        TabControl1.SelectTab(TabPage10)
    End Sub 'test for HIdden SLV
    Function testforclusep(ByVal a As Integer, ByVal b As Integer, ByVal _gruburst As reburst) As Integer
        Dim size As Integer = GridView1.RowCount
        Dim otum(size) As splits
        For k = 1 To size

            Dim x(1) As String
            x(1) = GridView1.Item(1, k - 1).Value
            otum(k).otumat1 = conca(x, 1, TextBox12.Text, TextBox13.Text, True, False)
            reductor(otum(k).otumat1)
        Next
        Dim dist As New medirdistancia
        hethandpp = 0
        Dim contacts As Integer = 0
        For i = 1 To _gruburst.gs(a).matg.GetLength(0) - 1
            For j = 1 To _gruburst.gs(b).matg.GetLength(0) - 1

                Dim dif As Integer = 0
                Dim realdif As Integer = 0
                For k = 1 To size
                    Dim distance As Single = 0
                    For l = 0 To otum(k).otumat1(1, 1).Length - 1
                        Dim d As Single = dist.pesos(otum(k).otumat1(_gruburst.gs(a).matg(i, 0), 1).Chars(l), otum(k).otumat1(_gruburst.gs(b).matg(j, 0), 1).Chars(l))
                        distance = distance + d
                        If d > 0.5 Then
                            'distance = distance + d
                            dif = dif + 1
                            Exit For
                        End If

                    Next
                    If distance > 0 Then
                        realdif = realdif + 1
                    End If
                Next
                If dif <= 0 Then
                    contacts = contacts + 1
                End If
            Next
        Next
        Return contacts
        Exit Function
    End Function

    ''Congruence Analyasis menu
    Private Sub SupportToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupportToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "BioNJ-ILD")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "NJ Localized Incongruence Length Diference Test", .page = 6, .viewnmin = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            CancelButton1.Visible = True
            blockmenusytab()
            ProgressBar1.Visible = True
            If a.temple = False Then
                ToolStripStatusLabel3.Text = "Permuting"
                Dim x As Integer = hethandpp
                If x = 1 Then
                    MsgBox("Just average states option for handling heterozygosities is allowed for this test, the parameter will be changed", MsgBoxStyle.Critical, "MLSTest")
                End If
                pdis = False
                Dim splits(,) As String = Module1.splitestmaker()
                hethand = 0
                support_conc(a.nrep, a._savetrees, splits)
                pdis = True
                hethand = x
            Else
                wstc1()
            End If
        End If
        restore()

    End Sub 'NJ-LILD button
    Private Sub ToolStripMenuItem2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "BioNJ-ILD")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "BioNJ Incongruence Length Diference Test", .page = 6}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Permuting"
            CancelButton1.Visible = True
            blockmenusytab()
            pdis = False
            ILD_bionj(a.nrep)
            pdis = True
        End If
        restore()



    End Sub 'ILD BIONJ test button
    Private Sub TopologicalIncongruenceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TopologicalIncongruenceToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        CancelButton1.Visible = True
        blockmenusytab()
        Dim bb As MsgBoxResult
        bb = MsgBox("Do you want to view a detailed analysis?. This includes wich loci are incongruent and the number of topological incongruences with a certain branch?", MsgBoxStyle.YesNo, "NJ-Analisis")
        If bb = MsgBoxResult.Yes Then
            NJtree(False, False, False, 0, False, False, False, True)
        Else
            NJtree(False, False, True, 0, False, False, False, False)
        End If

        


    End Sub 'topological incongruence button
    Private Sub TopologicalIncongrueceAnalisysToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TopologicalIncongrueceAnalisysToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Topological incongruence")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If

        ProgressBar2.Visible = True

        ToolStripStatusLabel4.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()

        top_incong(False)

        restore()



    End Sub 'tree distance among loci
    Function MCST(ByVal bionj As Boolean) As Integer

        Dim avstates As Boolean
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim arch As String = Nothing
        Dim a(1) As String
        Dim fi As Integer = 1
        Dim t As Integer = 0

        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next


        i = 1



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        ProgressBar2.Maximum = ch.Length - 1
        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then


                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If


            End If
        Next

        i = 1

        Dim sp(nmax - 1) As splits
        Dim resultado() As Integer

        Dim newa As New newicker
        Dim N As Integer
        Dim j As Integer
        For i = 1 To ch.Length - 1
            'This block make the splits tree for individual selected loci and save the splits in sp (splits) structure
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                arch1 = arch
                If arch1 = Nothing Then
                    arch1 = GridView1.Item(0, i - 1).Value
                Else
                    arch1 = arch1 & "_" & GridView1.Item(0, i - 1).Value
                End If
                t = t + 1

                sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, bionj, False, False)
                'Makes NJ trees
                N = sp(j).otumat1.GetLength(0) - 1
                For xi = N + 1 To sp(j).nOTUs.GetLength(0) - 1
                    sp(j).nOTUs(xi, 0) = newa.rewritesplits(sp(j).nOTUs(xi, 1), N)

                Next
                ProgressBar2.Increment(1)

                j = j + 1
                Application.DoEvents()

            End If
        Next
        Dim tr1 As Integer
        Dim tr2 As Integer
        Dim arrinco(10000, 3) As String

        Dim arri As Integer = 0
        For tree = 0 To sp.Length - 2

            tr1 = tree
            tr2 = tree + 1
            For ni = sp(tr1).otumat1.GetLength(0) To sp(tr1).nOTUs.GetLength(0) - 2
                For nx = sp(tr2).otumat1.GetLength(0) To sp(tr2).nOTUs.GetLength(0) - 2
                    If sp(tr1).notus1(ni, 0) > 10 ^ -9 And sp(tr2).notus1(nx, 0) > 10 ^ -9 Then
                        Dim comp As Boolean = checkcomp(sp(tr1).nOTUs(ni, 0), sp(tr2).nOTUs(nx, 0), sp(0).otumat1.GetLength(0) - 1)
                        If comp = False Then
                            Dim x() As String = pdeleted(sp(tr1).nOTUs(ni, 0), sp(tr2).nOTUs(nx, 0))


                            arrinco(arri, 0) = x(0)
                            arrinco(arri, 1) = x(1)
                            arrinco(arri, 2) = x(2)
                            arrinco(arri, 3) = x(3)
                            arri = arri + 1

                        End If
                    End If

                Next
            Next
        Next
        If arri <= 1 Then GoTo l222
        Dim indexes(arri - 1) As Byte
        Dim selindexes() As Byte
        Dim sum As Integer = 100000
        Dim selrndarr() As Integer
        Dim randomarray(arri - 1) As Integer
        For i = 0 To arri - 1
            randomarray(i) = i
        Next
        randomarray = permuteranksx(randomarray, arri - 1)

        Dim index As Integer
        Dim savedCST As String = Nothing
        Dim conse As String = Nothing
        Dim conseclone As String = Nothing

        For iii = 0 To N - 1
            conseclone = conseclone & "."
        Next
        Dim id As Integer

        Dim nrep As Integer
        nrep = 500
        Dim minconse As String = Nothing

        For r = 1 To 50
            Array.Clear(indexes, 0, indexes.Length)
            Dim sele(3) As Integer
            Dim min As Integer = 50000
            For i = 0 To 3
                Dim x As Integer = countg(arrinco(randomarray(0), i))
                sele(i) = x

            Next

            Dim minis(0) As Integer
            For i = 1 To N
                For j = 0 To 3
                    If sele(j) = i Then
                        minis(minis.Length - 1) = j
                        Array.Resize(minis, minis.Length + 1)
                        i = N + 1
                    End If
                Next
            Next
            If minis.Length > 1 Then
                Dim aa As New Random
                Dim bb As Integer = aa.Next(0, minis.Length - 1)
                id = minis(bb)
            Else
                id = minis(0)
            End If

            indexes(randomarray(0)) = id

            Dim act As Integer = 1
            Do

                For z = 0 To 3
                    conse = conseclone
                    min = 5000
                    id = 0
                    For jx = 0 To act
                        Dim jz As Integer = randomarray(jx)
                        Dim seldel As String = arrinco(jz, indexes(jz))
                        Dim conse1 As String = Nothing
                        For zz = 0 To N - 1
                            If seldel.Chars(zz) = "-" Or conse.Chars(zz) = "-" Then
                                conse1 = conse1 & "-"
                            Else
                                conse1 = conse1 & "."
                            End If
                        Next
                        conse = conse1

                    Next

                    Dim count As Integer = countg(conse)
                    sele(z) = count
                    indexes(randomarray(act)) = indexes(randomarray(act)) + 1
                Next
                Dim minis1(0) As Integer
                For i = 1 To N
                    For j = 0 To 3
                        If sele(j) = i Then
                            minis1(minis1.Length - 1) = j
                            Array.Resize(minis1, minis1.Length + 1)
                            i = N + 1
                        End If
                    Next
                Next
                If minis1.Length > 1 Then
                    Dim aa As New Random
                    Dim bb As Integer = aa.Next(0, minis1.Length - 1)

                    id = minis1(bb)
                Else
                    id = minis1(0)
                End If
                indexes(randomarray(act)) = id
                act = act + 1

            Loop While act < arri

            randomarray = permuteranksx(randomarray, arri - 1)
            If minconse = Nothing Then
                minconse = conse
                selrndarr = randomarray.Clone
                selindexes = indexes.Clone
            Else
                If countg(conse) < countg(minconse) Then
                    minconse = conse
                    selrndarr = randomarray.Clone
                    selindexes = indexes.Clone
                End If
            End If
        Next
        Dim minN As Integer = countg(minconse)

        'heuristic search
        Dim ori As Integer = 0
        For r = 0 To arri - 1
            Array.Clear(indexes, 0, indexes.Length)
            Dim min As Integer = 5000
            For i = 0 To 3
                Dim x As Integer = countg(arrinco(selrndarr(0), i))
                If x < min Then
                    min = x
                    id = i
                End If
            Next

            indexes(selrndarr(0)) = id

            Dim act As Integer = 1
            Do
                Dim arrmin(3) As Integer
                For z = 0 To 3
                    conse = conseclone
                    min = 5000
                    id = 0
                    For jx = 0 To act

                        Dim jz As Integer = selrndarr(jx)
                        Dim seldel As String = arrinco(jz, indexes(jz))
                        Dim conse1 As String = Nothing
                        For zz = 0 To N - 1
                            If seldel.Chars(zz) = "-" Or conse.Chars(zz) = "-" Then
                                conse1 = conse1 & "-"
                            Else
                                conse1 = conse1 & "."
                            End If
                        Next
                        conse = conse1

                    Next

                    Dim count As Integer = countg(conse)
                    arrmin(z) = count
                    If count < min Then
                        min = count
                        id = z
                    End If
                    indexes(selrndarr(act)) = indexes(selrndarr(act)) + 1
                Next
                If act = r Then
                    min = 5000
                    Dim idmin As Byte
                    For t = 0 To 3
                        If t <> id Then
                            If arrmin(t) < min Then
                                idmin = t
                                min = arrmin(t)
                            End If
                        End If
                    Next
                    ori = id
                    id = idmin
                End If
                indexes(selrndarr(act)) = id
                act = act + 1

            Loop While act < arri


            If minconse = Nothing Then
                minconse = conse
            Else
                If countg(conse) < countg(minconse) Then
                    minconse = conse
                End If
            End If
            selindexes(selrndarr(r)) = ori
        Next
        'createMCST(sp, conse)
        minN = N - countg(minconse)
        Return minN
        GoTo l223

l222:
        If arri = 0 Then
            Return N
        Else
            Return N - 1
        End If
l223:

    End Function
    Function countg(ByVal x As String) As Integer
        For i = 0 To x.Length - 1
            If x(i) = "-"c Then
                countg = countg + 1

            End If
        Next
        Return countg
    End Function
    Function createMCST(ByVal sp() As splits, ByVal excluded As String) As splits
        Dim deletN As Integer
        Dim exclclone As String = excluded
        Dim j As Integer
        Do
            If excluded(j) = "-"c Then
                For k = 0 To sp.Length - 1
                    For l = sp(0).otumat1.GetLength(0) To sp(0).nOTUs.GetLength(0) - 1
                        sp(k).nOTUs(l, 0) = sp(k).nOTUs(l, 0).Remove(j, 1)

                    Next
                Next
                excluded = excluded.Remove(j, 1)

                j = j - 1
            End If
            j = j + 1
        Loop While j <= excluded.Length - 1

        Dim spf(sp.Length - 1) As splits

        For k = 0 To sp.Length - 1
            ReDim spf(k).otumat1(excluded.Length, 1)
            ReDim spf(k).nOTUs((excluded.Length * 2) - 2, 3)
            ReDim spf(k).notus1((excluded.Length * 2) - 2, 1)
            Dim index As Integer = 0
            For l = 1 To sp(0).otumat1.GetLength(0) - 1
                If exclclone(l - 1) = "-"c Then
                    index = index + 1
                Else
                    spf(k).otumat1(l - index, 0) = sp(k).otumat1(l, 0)
                End If

            Next
        Next

        For k = 0 To sp.Length - 1
            Dim count As Integer = exclclone.Length - excluded.Length
            For l = sp(0).otumat1.GetLength(0) To sp(0).nOTUs.GetLength(0) - 2
                For m = l To sp(0).nOTUs.GetLength(0) - 1
                    If sp(k).nOTUs(l, 0) = sp(k).nOTUs(m, 0) Then
                        sp(k).nOTUs(m, 0) = "X"
                        count = count - 1
                    End If
                    If count = -1 Then
                        Exit For
                    End If
                Next
                If count = -1 Then
                    Exit For
                End If
            Next
        Next
        For k = 0 To sp.Length - 1
            Dim index As Integer = 0
            For l = sp(0).otumat1.GetLength(0) To sp(0).nOTUs.GetLength(0) - 2
                If sp(k).nOTUs(l, 0) = "X" Then
                    index = index + 1
                Else
                    spf(k).nOTUs(l - index, 0) = sp(k).nOTUs(l, 0)
                End If



            Next
        Next


    End Function
    Function pdeleted(ByVal split1 As String, ByVal split2 As String) As String()
        Dim N As Integer = split1.Length
        Dim x(3) As String

        For i = 0 To N - 1
            If split1.Chars(i) = split2.Chars(i) Then
                x(0) = x(0) & "."
                x(1) = x(1) & "."
                If split1.Chars(i) = "." Then
                    x(2) = x(2) & "-"
                    x(3) = x(3) & "."
                Else
                    x(2) = x(2) & "."
                    x(3) = x(3) & "-"
                End If
            Else
                If split1.Chars(i) = "*"c Then
                    x(0) = x(0) & "-"
                    x(1) = x(1) & "."
                Else
                    x(0) = x(0) & "."
                    x(1) = x(1) & "-"
                End If

                x(2) = x(2) & "."
                x(3) = x(3) & "."
            End If
        Next
        Return x
    End Function
    Sub top_incong(ByVal bionj As Boolean)
        Dim avstates As Boolean
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim arch As String = Nothing
        Dim a(1) As String
        Dim fi As Integer = 1
        Dim t As Integer = 0

        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next


        i = 1



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        ProgressBar2.Maximum = ch.Length - 1
        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then


                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If


            End If
        Next

        i = 1

        Dim sp(nmax - 1) As splits

        Dim j As Integer
        For i = 1 To ch.Length - 1
            'This block make the splits tree for individual selected loci and save the splits in sp (splits) structure
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                arch1 = arch
                If arch1 = Nothing Then
                    arch1 = GridView1.Item(0, i - 1).Value
                Else
                    arch1 = arch1 & "_" & GridView1.Item(0, i - 1).Value
                End If
                t = t + 1

                sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, bionj, False, False)
                'Makes NJ trees
               
                ProgressBar2.Increment(1)

                j = j + 1
                Application.DoEvents()

            End If
        Next


        Dim seqs(1) As String
        Dim nseq1 As Integer = sp.Length
        nseq = nseq1 - 1
        Dim distancias(nseq1, nseq1) As Single

        Dim bootstrapn() As Integer
        Dim dist As medirdistancia
        Dim DSTarr(10) As Integer

        Dim splitest(,) As String = Module1.splitestmaker


        For i = 1 To nseq + 1

            splitest(splitest.GetLength(0) - 1, 0) = splitest(splitest.GetLength(0) - 1, 0) & i & " "


        Next






        For i = 0 To nseq1

            For j = i + 1 To nseq1
                If i <> 0 Then
                    If i <> j Then
                        For x = 1 To sp(i - 1).nOTUs.GetLength(0) - 2
                            If sp(i - 1).notus1(x, 0) > 10 ^ -9 Then
                                Dim contain As Boolean = False
                                For z = 1 To sp(i - 1).nOTUs.GetLength(0) - 2
                                    If sp(j - 1).notus1(z, 0) > 10 ^ -9 Then
                                        If sp(i - 1).nOTUs(x, 2) = sp(j - 1).nOTUs(z, 2) Then
                                            contain = True
                                            Exit For
                                        End If
                                    End If
                                Next
                                If contain = False Then
                                  
                                    distancias(i, j) = distancias(i, j) + 1
                                End If
                            End If
                        Next

                        For x = 1 To sp(i - 1).nOTUs.GetLength(0) - 2
                            If sp(j - 1).notus1(x, 0) > 10 ^ -9 Then
                                Dim contain As Boolean = False
                                For z = 1 To sp(i - 1).nOTUs.GetLength(0) - 2
                                    If sp(i - 1).notus1(z, 0) > 10 ^ -9 Then
                                        If sp(j - 1).nOTUs(x, 2) = sp(i - 1).nOTUs(z, 2) Then
                                            contain = True
                                            Exit For
                                        End If
                                    End If
                                Next
                                If contain = False Then

                                    distancias(i, j) = distancias(i, j) + 1

                                End If
                            End If

                        Next
                        distancias(j, i) = distancias(i, j)

                    End If
                Else
                    distancias(i, j) = j
                    distancias(j, i) = j
                End If

            Next j
        Next i

       
        ReDim otumat1(nseq1, 1)
        Dim II As Integer = 1
        For i = 1 To ch.Length - 1

            If ch(i) = True Then
                otumat1(II, 0) = GridView1.Item(0, i - 1).Value
                II = II + 1
            End If
        Next

        distancias1 = distancias
        Dim sp1 As splits
        sp1 = UPGMAproc(nseq1, distancias, splitest, DSTarr, otumat1, False, 0)

        Dim nwkform As New treeviewer2 With {.sptree = sp1, .Text = "Tree Viewer", ._upgma = True}
        nwkform.Show()
    End Sub 'It calculates tree distance among loci
    Private Sub TestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Consensus Network", .page = 3, .nloc = GridView1.RowCount - 1}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then

            ProgressBar2.Visible = True

            ToolStripStatusLabel4.Visible = True
            CancelButton1.Visible = True
            blockmenusytab()
            nlocmin = a.gdef

            ConsNet(a.gdef)
            restore()

        End If

    End Sub 'Consensus Network button
    Private Sub WToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WToolStripMenuItem.Click
        If checksel() = 0 Then

            MsgBox("None locus has been fixed. You have to select only one loci", MsgBoxStyle.OkOnly, "Winning-Sites test")
            Exit Sub
            'ElseIf checksel() <> 1 Then
            'MsgBox("There are multiple loci selected. You have to select only one loci", MsgBoxStyle.OkOnly, "Winning-Sites test")

            'Exit Sub
        End If


        ProgressBar1.Visible = True
        ToolStripStatusLabel3.Text = "calculating"
        CancelButton1.Visible = True
        blockmenusytab()
        pdis = False
        wstc()
        pdis = True

        restore()
    End Sub 'Winning sites test button
    Sub wstc()

        stopp = False
        Dim respuesta As MsgBoxResult
        respuesta = MsgBox("Do you want a branch by branch analysis?", MsgBoxStyle.YesNo)

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False




        Dim a(checksel()) As String

        Dim ch(GridView1.Rows.Count) As String
        Dim fi As Integer = 1
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(1, i - 1).Value
            If GridView1.Item(2, i - 1).Value = True Then
                a(fi) = GridView1.Item(1, i - 1).Value
                fi = fi + 1
            End If

        Next
        Dim topoc, topoi, topoc1 As splits

        topoc = Module1.NJ(conca(ch, GridView1.RowCount, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)


        Dim otumat1(,) As String = conca(a, a.Length - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        otumat1 = reductor(otumat1)
        If hethand = 1 Then

            otumat1 = SNPmod.duplicate(otumat1, otumat1(1, 1).Length, otumat1.GetLength(0) - 2)
        End If
        Dim distam As dsx
        'distam = NJgo(conca(ch, GridView1.RowCount, TextBox12.Text, TextBox13.Text, avstates, False), False, False)
        distam = NJgo(otumat1, False, False)
        topoi = NJproc(distam.distancias.GetLength(0) - 1, distam.distancias, otumat1, False, 0, Nothing, False)
        Dim testbranchs As Boolean

        If respuesta = MsgBoxResult.Yes Then
            testbranchs = True
        Else
            testbranchs = False
        End If
        If testbranchs = True Then
            Dim newa As New newicker

            ProgressBar1.Maximum = topoc.nOTUs.GetLength(0) - 2 - otumat1.GetLength(0)
            For branch = otumat1.GetLength(0) To topoc.nOTUs.GetLength(0) - 2
                Dim split As String = newa.rewritesplits(" " & topoc.nOTUs(branch, 2), otumat1.GetLength(0) - 1)
                topoc1 = NJproc1(distam.distancias.GetLength(0) - 1, distam.distancias, Nothing, False, split, True)
                topoc.notus1(branch, 1) = wst(topoc1, topoi.nOTUs, otumat1)
                ProgressBar1.Increment(1)
            Next
            Dim supname(0) As String
            supname(0) = "p-value"
            Dim nwkform As New treeviewer2 With {.sptree = topoc, .Text = "Tree Viewer", ._viewsupport = True, ._supportnames = supname} '

            nwkform.Show()
        Else
            Dim pval As Single = wst(topoc, topoi.nOTUs, otumat1)
            MsgBox("p value is " & pval)

        End If
        restore()
        pdis = True
    End Sub 'Makes Wining Sites test for trees 
    Function wstc1() As Single()

        'Calculates NJ-ILD for each branch on a tree

        stopp = False
        ProgressBar1.Visible = True


        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0
        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        Dim selocus(nmax - 1) As String
        Dim p As Integer
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                selocus(p) = GridView1(1, i - 1).Value
                p = p + 1
            End If
        Next

        i = 1

        '-------------------------------------------------------
        'makes individual alignments
        Dim sp(nmax - 1) As splits
        Dim j As Integer = 0
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                Dim concatenate(,) As String
                concatenate = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                concatenate = Module1.reductor(concatenate)
                sp(j).nsites = concatenate(1, 1).Length
                ProgressBar2.Increment(1)
                j = j + 1
                Application.DoEvents()
            End If
        Next
        ProgressBar1.Maximum = DataGridView7.RowCount

        Dim sp1 As splits
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)
                fi = fi + 1
            End If
        Next
        '-----------------------------------------------

        'Makes concatenated tree
        Dim concat(,) As String = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        Dim treelengths(nmax - 1) As Single
        sp1 = Module1.NJ(concat, Module1.splitestmaker, 0, False, False, False)
        Dim newa As New newicker
        For i = 1 To sp1.nOTUs.GetLength(0) - 1
            sp1.nOTUs(i, 0) = newa.rewritesplits(sp1.nOTUs(i, 1), sp1.otumat1.GetLength(0) - 1)
        Next
        a = Nothing

        '-------------------------------------------------------------
        Dim topoc, topoi, topoc1 As splits


        Dim distam As dsx
        'Makes individual trees 
        ReDim a(1)
        ReDim Module1.otum(nmax - 1)
        Dim ssp(nmax - 1) As splits
        For i = 0 To nmax - 1

            otum(i).treelength = 0
            a(1) = GridView1.Item(1, i).Value
            otum(i).otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, True, False)
            otum(i).dista = distanciasx(otum(i).otumat1, otum(i).otumat1.GetLength(0) - 1)

            ssp(i) = Module1.NJ(otum(i).otumat1, Module1.splitestmaker, 0, False, False, False)
           




        Next

        For branch = sp1.otumat1.GetLength(0) To sp1.nOTUs.GetLength(0) - 2
            Dim x(0) As Single
            For i = 0 To nmax - 1

                Dim split As String = newa.rewritesplits(" " & sp1.nOTUs(branch, 2), sp1.otumat1.GetLength(0) - 1)
                topoc1 = NJproc1(otum(i).dista.GetLength(0) - 1, otum(i).dista, Nothing, False, split, True)
                x = templ(topoc1, ssp(i).nOTUs, otum(i).otumat1, x, otum(i).treelength)

            Next
            ProgressBar1.Increment(1)

            sp1.notus1(branch, 1) = wilcoxon(x)
            ProgressBar1.Increment(1)
            Application.DoEvents()
        Next
        Dim R(7) As Single
        For m = 0 To 5
            R(m) = 999
        Next
        Dim COUNTTP As Single = 0.05
        Dim COUNTFP, countbc As Integer
        Dim bcorr As Boolean = False
       

        For f = sp1.otumat1.GetLength(0) To sp1.notus1.GetLength(0) - 2


            If sp1.notus1(f, 1) < 0.0028 Then
                bcorr = True
                R(7) = R(7) + 1
            End If
            Select Case sp1.nOTUs(f, 2)
                Case "1 2 ,"

                    R(0) = sp1.notus1(f, 1)

                Case "1 2 3 4 ,"

                    R(4) = sp1.notus1(f, 1)

                Case "1 2 3 4 7 8 ,"

                    R(3) = sp1.notus1(f, 1)

                Case "1 2 5 6 7 8 ,"

                    R(2) = sp1.notus1(f, 1)

                Case "1 2 3 4 5 6 ,"

                    R(1) = sp1.notus1(f, 1)

                Case Else
                    R(6) = R(6) + 1
                    If CSng(sp1.notus1(f, 1)) < R(5) Then
                        R(5) = sp1.notus1(f, 1)
                    End If
            End Select


        Next






        wstc1 = R
        Dim re As Integer
        If sp1.otumat1.GetLength(0) - 3 <= 50 Then
            re = 4
        Else
            re = 6
        End If



        For i = sp1.otumat1.GetLength(0) To sp1.notus1.GetLength(0) - 1

            sp1.notus1(i, 1) = Math.Round(CDbl(sp1.notus1(i, 1)), re)
        Next
        Dim supname(0) As String
        supname(0) = "p-value"
        Dim nwkform As New treeviewer2 With {.sptree = sp1, .Text = "Tree Viewer", ._viewsupport = True, ._supportnames = supname} '

        nwkform.Show()

        restore()
        pdis = True
    End Function 'Makes Wining Sites test for trees 
    Function wst(ByVal topoc As splits, ByVal topoi As String(,), ByVal alignm(,) As String) As Single
        Dim sitenmat1(alignm(1, 1).Length - 1) As Double
        Dim N As Integer = alignm.GetLength(0) - 1
        Dim newa As New newicker
        For i = N + 1 To topoc.nOTUs.GetLength(0) - 1
            topoc.nOTUs(i, 0) = newa.rewritesplits(topoc.nOTUs(i, 1), alignm.GetLength(0) - 1)
            topoi(i, 0) = newa.rewritesplits(topoi(i, 1), alignm.GetLength(0) - 1)
        Next

        Dim x(0) As Single
        Dim count, count1 As Integer
        For i = 0 To sitenmat1.Length - 1
            Dim L As Double = 0
            Dim L1 As Double = 0
            For j = 1 To N
                Dim dist As New medirdistancia
                Dim dista As Single
                For k = j + 1 To N

                    If alignm(j, 1).Substring(i, 1) <> alignm(k, 1).Substring(i, 1) Then
                        dista = dist.pesos(alignm(j, 1).Substring(i, 1), alignm(k, 1).Substring(i, 1))
                    Else : dista = 0
                    End If

                    If dista > 0 Then



                        Dim topo1 As Single = (calcB(topoi, j, k) - 1)
                        Dim topo2 As Single = (calcB(topoc.nOTUs, j, k) - 1)


                        L = L + dista / 2 ^ topo1
                        L1 = L1 + dista / 2 ^ topo2
                    End If
                        Next 
            Next
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

        Dim a As Single = wilcoxon(x)
        'If c < 0 Then Stop

        RFD = (f - c) / f

        Return a 'RFD 'count / count1 ' 

        '    End If

        '
        '                Next

        '            Next
        'Application.DoEvents()
        'If L1 - L > 10 ^ -9 Then
        'count = count + 1

        'End If

        '   If Math.Abs(L - L1) > 10 ^ -6 Then
        'count1 = count1 + 1

        '                    End If



        '               Next
        'If count1 > 0 Then
        'Return combinatF1(0.5, count1, count)
        'Else : Return 1
        'End If

    End Function 'Calculates p value for a comparison between two topologies using sign test 
    Function templ(ByVal topoc As splits, ByVal topoi As String(,), ByVal alignm(,) As String, ByVal x() As Single, ByVal tl As Single) As Single()
        Dim sitenmat1(alignm(1, 1).Length - 1) As Double
        Dim N As Integer = alignm.GetLength(0) - 1
        Dim newa As New newicker
        For i = N + 1 To topoc.nOTUs.GetLength(0) - 1
            topoc.nOTUs(i, 0) = newa.rewritesplits(topoc.nOTUs(i, 1), alignm.GetLength(0) - 1)
            topoi(i, 0) = newa.rewritesplits(topoi(i, 1), alignm.GetLength(0) - 1)
        Next


        Dim count, count1 As Integer
        count1 = x.Length - 1
        For i = 0 To sitenmat1.Length - 1
            Dim L As Double = 0
            Dim L1 As Double = 0
            For j = 1 To N
                Dim dist As New medirdistancia
                Dim dista As Single
                For k = j + 1 To N

                    If alignm(j, 1).Substring(i, 1) <> alignm(k, 1).Substring(i, 1) Then
                        dista = dist.pesos(alignm(j, 1).Substring(i, 1), alignm(k, 1).Substring(i, 1))
                    Else : dista = 0
                    End If

                    If dista > 0 Then



                        Dim topo1 As Single = (calcB(topoi, j, k) - 1)
                        Dim topo2 As Single = (calcB(topoc.nOTUs, j, k) - 1)


                        L = L + dista / 2 ^ topo1
                        L1 = L1 + dista / 2 ^ topo2
                    End If
                Next
            Next
            If L1 - L > 0 Then
                count = count + 1

            End If

            If Math.Abs(L - L1) > 0 Then
                count1 = count1 + 1
                Array.Resize(x, count1)
                x(count1 - 1) = (L1 - L)
            End If


        Next






        'If c < 0 Then Stop

        Return x 'RFD 'count / count1 ' 

        '    End If

        '
        '                Next

        '            Next
        'Application.DoEvents()
        'If L1 - L > 10 ^ -9 Then
        'count = count + 1

        'End If

        '   If Math.Abs(L - L1) > 10 ^ -6 Then
        'count1 = count1 + 1

        '                    End If



        '               Next
        'If count1 > 0 Then
        'Return combinatF1(0.5, count1, count)
        'Else : Return 1
        'End If

    End Function 'Calculates p value for a comparison between two topologies using sign test 
    'Minimum set menu

    Private Sub DTUToTestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTUToTestToolStripMenuItem.Click
        chargeDTUtotest()

    End Sub 'Groups to test menu
    Private Sub UseUserDeterminedDTUsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseUserDeterminedDTUsToolStripMenuItem.Click
        If DataGridView6.RowCount = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("First, you have to select Groups to test. Do you Want to do this now?", MsgBoxStyle.OkCancel)
            If aa = MsgBoxResult.Ok Then
                chargeDTUtotest()


            End If

            Exit Sub
        End If

        Dim a As New Dialog2 With {.Text = "Test combination of loci", .page = 4, .nloc = GridView1.RowCount - 1, .viewnmin = True, ._supp = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            'ProgressBar1.Visible = True
            ProgressBar2.Visible = True
            'ToolStripStatusLabel3.Text = "Bootstrapping"
            ToolStripStatusLabel4.Visible = True
            CancelButton1.Visible = True
            blockmenusytab()
            'TabControl1.Enabled = False
            nlocmin = a.gdef
            strees = a._savetrees
            testseldtu(a.gdef, a.nrep, a._savetrees, False, a._supp)
            externo = False
            Panel4.Visible = True
            restore()
        End If




    End Sub 'Test user determined groups button
    Private Sub CompareAgainstAllLociTreeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompareAgainstAllLociTreeToolStripMenuItem.Click
        Dim a As New Dialog2 With {.Text = "Test combination of loci", .page = 4, .nloc = GridView1.RowCount - 1, .viewnmin = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            'ProgressBar1.Visible = True
            ProgressBar2.Visible = True
            'ToolStripStatusLabel3.Text = "Bootstrapping"
            ToolStripStatusLabel4.Visible = True
            CancelButton1.Visible = True
            blockmenusytab()
            nlocmin = a.gdef
            strees = a._savetrees
            testseldtu(a.gdef, a.nrep, a._savetrees, True, a._supp)

            restore()

            Panel4.Visible = True
        End If


    End Sub 'test against all loci tree button
    Private Sub checklocus(ByVal nmin As Integer, ByVal nrep As Integer, ByVal savetrees As Boolean)

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String



        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim texto As String
        ' DataGridView4.Rows.Clear()

        TabControl1.SelectTab(TabPage11)
        Panel3.Enabled = False
        Dim nlocus As Integer = GridView1.Rows.Count
        Dim minimum(nlocus) As Integer
        Dim maximum(nlocus) As Single
        Dim promedio(nlocus) As Single
        Dim combinations(nlocus) As Integer
        For nmin = 1 To nlocus
            ProgressBar2.Visible = True
            Dim a(nmin) As String
            Dim an(nmin) As Integer

            Dim nmax As Integer = factorial(nlocus) / (factorial(nmin) * factorial(nlocus - nmin))
            Dim arrayx(nmax) As Integer
            Dim arrayarch(nmax) As String
            ProgressBar2.Maximum = nmax
            Dim t As Integer = 0


            For i = 1 To locus.Length - 1
                locus(i) = GridView1.Item(1, i - 1).Value
            Next

            Dim matr(nlocus) As dsx
            Dim otumat1 As String(,)
            pdis = False
            For i = 1 To locus.Length - 1
                locus(i) = GridView1.Item(1, i - 1).Value
                a(1) = locus(i)
                otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                matr(i) = Module1.NJgo(otumat1, False, False)
            Next
            pdis = True
            stopp = False


            For i = 1 To nmax
                textito = Nothing
                If an(1) = 0 Then
                    For j = 1 To nmin
                        an(j) = j
                    Next

                Else


                    If an(nmin) < nlocus Then
                        an(nmin) = an(nmin) + 1
                    Else
                        For j = 1 To nmin
                            If an(nmin - j) < nlocus - j Then
                                an(nmin - j) = an(nmin - j) + 1
                                For k = nmin - j + 1 To nmin
                                    an(k) = an(k - 1) + 1
                                Next


                                j = nmin


                            End If


                        Next
                    End If
                End If
                arch = Nothing

                For k = 1 To an.Length - 1
                    a(k) = locus(an(k))
                    If arch = Nothing Then
                        arch = GridView1.Item(0, an(k) - 1).Value
                    Else
                        arch = arch & ", " & GridView1.Item(0, an(k) - 1).Value
                    End If
                Next
                arch = arch & ","
                Dim matf As dsx = Nothing
                ReDim matf.distancias(matr(1).distancias.GetLength(0) - 1, matr(1).distancias.GetLength(0) - 1)


                For k = 1 To an.Length - 1
                    For r = 1 To matr(1).distancias.GetLength(0) - 1
                        For s = 1 To matr(1).distancias.GetLength(0) - 1
                            matf.distancias(r, s) = matf.distancias(r, s) + matr(an(k)).distancias(r, s)
                        Next
                        matf.distancias(r, 0) = r
                        matf.distancias(0, r) = r
                    Next
                Next

                Dim Dst As Integer

                Dst = Module1.dstcalc(matf.distancias)

                arrayx(i) = Dst
                arrayarch(i) = arch
                ProgressBar2.Increment(1)
                t = t + 1
                Label20.Text = "Evaluated trees: " & t & "/" & ProgressBar2.Maximum

                If stopp = True Then

                    Exit Sub
                End If
                Application.DoEvents()
            Next

            promedio(nmin) = ponderado(arrayx)
            'promedio(nmin) = probmax(arrayx)
            maximum(nmin) = arrayx.Max
            texto = texto & nmin & " loci found a maximum of " & maximum(nmin) & " ST" & vbNewLine & "Combinations analyzed: " & nmax & vbNewLine
            For s = 1 To arrayx.Length - 1
                If arrayx(s) = maximum(nmin) Then
                    texto = texto & arrayarch(s) & vbNewLine
                End If
            Next
            arrayx(0) = 1000
            texto = texto & Environment.NewLine

            minimum(nmin) = arrayx.Min
            combinations(nmin) = nmax
            ProgressBar2.Value = 0

        Next
        RichTextBox1.Text = Nothing
        RichTextBox1.Text = texto

        TabControl1.SelectTab(TabPage11)
        DataGridView5.Visible = True
        DataGridView5.RowCount = minimum.Length - 1
        DataGridView5.Columns(0).HeaderText = "Number of Loci (Number of combinations)"
        DataGridView5.Columns(1).HeaderText = "Minimum Number of ST found"
        DataGridView5.Columns(2).HeaderText = "Mean Number of ST found"
        DataGridView5.Columns(3).HeaderText = "Maximum Number of ST found"
        Dim xypoints(minimum.Length - 1, 2) As Double
        For i = 1 To minimum.Length - 1
            xypoints(i, 0) = i
            xypoints(i, 1) = i
            xypoints(i, 2) = promedio(i)
            DataGridView5.Item(0, i - 1).Value = i & " (" & combinations(i) & ")"
            DataGridView5.Item(1, i - 1).Value = minimum(i)
            DataGridView5.Item(2, i - 1).Value = promedio(i)
            DataGridView5.Item(3, i - 1).Value = maximum(i)
        Next
        ProgressBar2.Value = 0
        Panel3.Enabled = True


        Dim mdsviewer As New treeviewer2 With {.xypoints = xypoints, .Text = "Number of Loci(x) vs Mean GD(y)"}
        mdsviewer.Show()
        restore()
    End Sub 'Calculates minimum, mean, maximum GD for every possible combination from 2 to n-1 loci
    Private Sub LoadTestsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadTestsToolStripMenuItem.Click
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Output of MLSTest (*.xml)|*.xml"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName <> "" Then

            leeroutput(OpenFileDialog1.FileName)
            TabControl1.SelectTab(TabPage9)
            Panel4.Visible = False
        End If

    End Sub 'Load output button
    Private Sub DeleteOneToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteOneToolStripMenuItem.Click
        If DataGridView4.RowCount > 0 Then
            exporto()
            restore()
        Else
            MsgBox("No test to save", MsgBoxStyle.OkOnly, "MLSTest")

        End If

        
    End Sub 'Export button
    Private Sub DataGridView4_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.RowHeaderMouseClick

        Dim tree As String
        tree = DataGridView4.Item(DataGridView4.ColumnCount - 1, DataGridView4.CurrentCell.RowIndex).Value.ToString
        If tree <> "" Then
            Dim sp As splits
            Dim nwk As New newicker
            sp = nwk.tosplit(tree)
            Dim nwkform As New treeviewer2 With {.sptree = sp, .Text = DataGridView4.Item(DataGridView4.ColumnCount - 2, DataGridView4.CurrentCell.RowIndex).Value.ToString} '


            nwkform.Show()
        End If
        treeinfo()
    End Sub
    Private Sub ListBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox4.SelectedIndexChanged
        If ListBox4.SelectedIndices.Count <> 0 Then
            For i = 0 To ListBox4.SelectedIndices.Count - 1
                If ListBox6.Items.Contains(ListBox4.SelectedItems(i)) = False Then
                    ListBox6.Items.Add(ListBox4.SelectedItems(i))
                End If
            Next
        End If
        Button24.Focus()
    End Sub
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        If ListBox6.Items.Count = 0 Then
            Exit Sub
        End If
        DataGridView6.Rows.Add()
        DataGridView6.Item(0, DataGridView6.RowCount - 1).Value = TextBox8.Text
        Dim split As String = " "
        Dim strains As String
        For i = 0 To ListBox6.Items.Count - 1

            split = split & ListBox4.Items.IndexOf(ListBox6.Items(i)) + 1 & " "
            strains = strains & ListBox6.Items(i) & ", "

        Next
        DataGridView6.Item(3, DataGridView6.RowCount - 1).Value = split
        DataGridView6.Item(2, DataGridView6.RowCount - 1).Value = stdsplit(split, TextBox12.Text)
        DataGridView6.Item(1, DataGridView6.RowCount - 1).Value = strains
        ListBox6.Items.Clear()

    End Sub '"Add" button in select Groups to test window. It adds a group to test
    Private Sub HideSelectedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HideSelectedToolStripMenuItem.Click

        If DataGridView4.SelectedRows.Count <> 0 Then
            Dim arr(DataGridView4.SelectedRows.Count - 1) As Integer
            For i = 0 To arr.Length - 1
                arr(i) = DataGridView4.SelectedRows(i).Index
            Next
            Dim b As String = bs.Filter
            For i = 0 To arr.Length - 1
                If DataGridView4.Rows.GetRowCount(DataGridViewElementStates.Visible) > 1 Then
                    Dim cellpos As Integer = arr(i)
                    Dim a As String
                    a = DataGridView4.Columns(DataGridView4.ColumnCount - 2).Name

                    If b = Nothing Then
                        b = a & " <> '" & DataGridView4.Item(a, cellpos).Value.ToString & "'"
                    Else
                        b = b & " And " & a & " <> '" & DataGridView4.Item(a, cellpos).Value.ToString & "'"
                    End If


                End If

            Next
            bs.Filter = b
        Else
            MsgBox("Please select entire rows clicking on the row header", MsgBoxStyle.OkOnly, "MLSTest1")
        End If

    End Sub 'Hide selected combinations
    Private Sub UserDeterminedGroupsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserDeterminedGroupsToolStripMenuItem.Click, AllGroupsToolStripMenuItem.Click
        Dim all As Boolean
        If sender Is AllGroupsToolStripMenuItem Then all = True
        If all = False Then
            If DataGridView6.RowCount = 0 Then
                Dim aa As MsgBoxResult
                aa = MsgBox("First, you have to select DTU to test. Do you Want to do this now?", MsgBoxStyle.OkCancel)
                If aa = MsgBoxResult.Ok Then
                    chargeDTUtotest()


                End If

                Exit Sub
            End If
        End If


        Dim s As Integer
        For i = 0 To GridView1.Rows.Count - 1
            If GridView1.Item(2, i).Value = True Then
                s = s + 1
            End If
        Next

        If s = 0 Or s = 1 Then
            For i = 0 To GridView1.Rows.Count - 1
                GridView1.Item(2, i).Value = True
            Next
        End If

        Dim a As New Dialog2 With {.Text = "Test combination of loci", .page = 1, .nloc = GridView1.RowCount - 1, .viewnmin = False, ._savetrees = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ProgressBar2.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            ToolStripStatusLabel4.Visible = True
            CancelButton1.Visible = True
            blockmenusytab()
            nlocmin = a.gdef
            strees = a._savetrees
            numerodereplicas = a.nrep
            stepwiseDEL(a.nrep, a._savetrees, all, False)
            If all = False And a.nrep > 0 Then ToolStripMenuItem6.Enabled = True
            If a.nrep > 0 Then
                CheckBox3.Checked = True

                CheckBox2.Checked = True
                CheckBox2.Checked = False
            End If
            Panel4.Visible = True

        End If
    End Sub 'test all combinations of loci button
    Private Sub DiscardLowBootstrapTreesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiscardLowBootstrapTreesToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Discard Low bootstrap Trees")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If

        ProgressBar2.Visible = True

        ToolStripStatusLabel4.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        Dim ss As Boolean
        If DataGridView6.RowCount > 0 Then
            ss = True
        End If
        Dim a As New Dialog2 With {.Text = "Discard low bootstrap trees", .page = 5, .viewboo = ss, .viewnmin = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrap cutoff"
            CancelButton1.Visible = True
            blockmenusytab()
            'NJtree(False, a._supp, a.nrep, a._bionj)
            unsellowboots(a._bionj, a._cutoff, a.nrep, a._savetrees, False)
        End If

        restore()
    End Sub 'Discard low bootstrap trees button
    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        If DataGridView4.RowCount > 0 Then
            TabControl1.SelectTab(TabPage9)
        Else
            MsgBox("There is no Tests loaded", MsgBoxStyle.OkOnly)

        End If
    End Sub 'View tests button
    Private Sub OrderLociByBootstrapMeanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrderLociByBootstrapMeanToolStripMenuItem.Click


        For i = 0 To GridView1.RowCount - 1
            GridView1.Item(2, i).Value = True
        Next




        ProgressBar2.Visible = True

        ToolStripStatusLabel4.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        Dim ss As Boolean
        If DataGridView6.RowCount > 0 Then
            ss = True
        End If
        Dim a As New Dialog2 With {.Text = "Order loci By bootstrap mean", .page = 5, .viewboo = ss, .viewnmin = False}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrap cutoff"
            CancelButton1.Visible = True
            blockmenusytab()
            unsellowboots(a._bionj, a._cutoff, a.nrep, a._savetrees, True)
            MsgBox("Loci were ordered from higher to lower mean bootstrap, move your mouse over the locus to view bootstrap mean value", , "MLSTest1")
        End If


        restore()

    End Sub 'order loci by bootstrap value
    Private Sub AddOneToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddOneToolStripMenuItem.Click
        If DataGridView6.RowCount = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("First, you have to select DTU to test. Do you Want to do this now?", MsgBoxStyle.OkCancel)
            If aa = MsgBoxResult.Ok Then
                chargeDTUtotest()


            End If

            Exit Sub
        End If

        Dim a As New Dialog2 With {.Text = "Test combination of loci", .page = 1, .nloc = GridView1.RowCount - 1, .viewnmin = False, ._savetrees = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ProgressBar2.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            ToolStripStatusLabel4.Visible = True
            CancelButton1.Visible = True
            blockmenusytab()
            nlocmin = a.gdef
            strees = a._savetrees
            stepwiseadd(a.nrep, a._savetrees)
            numerodereplicas = a.nrep

            externo = False
            Panel4.Visible = True
            If a.nrep > 0 Then
                CalculateBootstrapImprovementToolStripMenuItem.Enabled = True
                CheckBox3.Checked = True
                CheckBox2.Checked = True
                CheckBox2.Checked = False
            End If

        End If


    End Sub 'stepwise addition button
    Private Sub CalculateBootstrapImprovementToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculateBootstrapImprovementToolStripMenuItem.Click
        Dim ss As MsgBoxResult
        ss = MsgBox("Absolute bootstrap values will be replaced by relative values. Do you want to export them previously?", MsgBoxStyle.YesNo)
        If ss = MsgBoxResult.Yes Then
            exporto()
        End If
        Dim avstates As Boolean
        Dim hethand
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        comparetofull(avstates, numerodereplicas)
        CalculateBootstrapImprovementToolStripMenuItem.Enabled = False
    End Sub 'bootstrap improvement/impairmnent calculation
    Private Sub ToolStripMenuItem6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        Dim ss As MsgBoxResult
        ss = MsgBox("Absolute bootstrap values will be replaced by relative value. Do you want to export them previously?", MsgBoxStyle.YesNo)
        If ss = MsgBoxResult.Yes Then
            exporto()
        End If
        Dim avstates As Boolean
        Dim hethand
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        comparetofull(avstates, numerodereplicas)
        ToolStripMenuItem6.Enabled = False
    End Sub 'bootstrap improvement/impairmnent calculation
    Private Sub ToolStripComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        If ToolStripComboBox2.SelectedIndex <> -1 Then
            Dim fileq As String = tests(ToolStripComboBox2.SelectedIndex + 1)
            If IO.File.Exists(fileq) Then
                leeroutput(fileq)
            Else
                MsgBox("The file " & fileq & " does not exists", MsgBoxStyle.OkOnly, "MLSTest")
                Dim tests1(tests.Length - 2) As String
                Dim b As Integer = 1
                For i = 1 To tests.Length - 1
                    If i <> ToolStripComboBox2.SelectedIndex + 1 Then
                        tests1(b) = tests(i)
                        b = b + 1

                    End If
                Next
                tests = tests1.Clone
                ToolStripComboBox2.Items.RemoveAt(ToolStripComboBox2.SelectedIndex)
                Exit Sub
            End If

            leeroutput(tests(ToolStripComboBox2.SelectedIndex + 1))
        End If
    End Sub 'load a test into the project
    Private Sub BootstrapSignificanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BootstrapSignificanceToolStripMenuItem.Click
        Form2.Show()
    End Sub 'bootstrap significance button

    ''subs of Minimum set 
    Sub chargeDTUtotest()
        ListBox4.Items.Clear()
        DataGridView7.EndEdit()
        If DataGridView7.RowCount <> 0 Then
            For i = 0 To DataGridView7.RowCount - 1
                If DataGridView7.Item(1, i).Value = True Then
                    ListBox4.Items.Add(DataGridView7.Item(0, i).Value)
                End If
            Next
        End If
        TabControl1.SelectTab(TabPage3)
    End Sub 'load the selected strains in the listbox of DTU to test
    Private Sub testseldtu(ByVal nmin As Integer, ByVal nrep As Integer, ByVal savetrees As Boolean, ByVal all As Boolean, ByVal facs As Boolean)
        Dim SPLITEST As String(,)
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String

        Dim a(1) As String
        Dim an(nmin) As Integer

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        DataGridView4.DataSource = Nothing

        TabControl1.SelectTab(TabPage9)
        Panel3.Enabled = False
        Panel4.Enabled = False
        Dim nlocus As Integer = GridView1.Rows.Count
        Dim nmax As Long
        'factorial(nlocus) / (factorial(nmin) * factorial(nlocus - nmin))
        nmax = 1
        For ff = nlocus - nmin + 1 To nlocus
            nmax = nmax * ff
        Next
        nmax = nmax / factorial(nmin)

        Dim matb(,) As dsx
        ProgressBar2.Maximum = nmax
        Dim t As Integer = 0
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        If all = True Then
            Dim sp As splits
            ListBox4.Items.Clear()
            DataGridView7.EndEdit()

            If DataGridView7.RowCount <> 0 Then
                For dd = 0 To DataGridView7.RowCount - 1
                    If DataGridView7.Item(1, dd).Value = True Then
                        ListBox4.Items.Add(DataGridView7.Item(0, dd).Value)
                    End If
                Next
            End If
            sp = Module1.NJ(conca(locus, locus.Length - 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)
            SPLITEST = makesplitestfromtree(sp.nOTUs, sp.notus1)

            DataGridView8.RowCount = SPLITEST.GetLength(0) - 2

            For i = 0 To DataGridView8.RowCount - 1

                Dim strains As String = Nothing
                For v = 1 To ListBox4.Items.Count
                    If SPLITEST(i + 1, 0).Contains(" " & v & " ") Then
                        strains = strains & ListBox4.Items(v - 1) & ", "
                    End If
                Next
                DataGridView8.Item(0, i).Value = "split " & i + 1
                DataGridView8.Item(1, i).Value = strains
                DataGridView8.Item(2, i).Value = SPLITEST(i + 1, 0)
                DataGridView8.Item(3, i).Value = SPLITEST(i + 1, 1)

            Next
            externo = True
        Else
            SPLITEST = Module1.splitestmaker

            externo = False
        End If


        ''''
        Dim ch(GridView1.RowCount) As String
        For i = 0 To GridView1.RowCount - 1
            ch(i + 1) = GridView1.Item(1, i).Value
        Next
        Dim spx, spalt As splits
        Dim dista As dsx
        Dim counts(,,) As Integer
        If facs = True Then
            ReDim counts(ch.Length - 1, SPLITEST.GetLength(0) - 2, 1)
            Dim otumatx As String(,)
            otumatx = conca(ch, ch.Length - 1, TextBox12.Text, TextBox13.Text, avstates, False)
            lseq = otumatx(1, 1).Length
            dista = Module1.NJgo(otumatx, False, False)
            spx = Module1.NJproc(dista.distancias.GetLength(0) - 1, dista.distancias, Nothing, False, 0, Nothing, False)

        End If

        ''''


        Dim matr(nlocus) As dsx
        Dim otumat1 As String(,)
        If nrep > 0 Then
            ReDim matb(nlocus, nrep - 1)
            ProgressBar1.Maximum = nrep * nlocus
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            ToolStripStatusLabel3.Visible = True
        ElseIf facs = True Then
            ProgressBar1.Maximum = (SPLITEST.GetLength(0) - 2) * (nlocus - 1)
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "testing sites"
            ToolStripStatusLabel3.Visible = True
        End If
        For i = 1 To locus.Length - 1

            a(1) = locus(i)
            pdis = False
            otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
            If hethand = 1 Then
                otumat1 = reductor(otumat1)
                otumat1 = SNPmod.duplicate(otumat1, otumat1(1, 1).Length, otumat1.GetLength(0) - 2)
            End If

            lseq = otumat1(1, 1).Length

            If nrep > 0 Then
                For x = 0 To nrep - 1
                    matb(i, x) = NJgo(otumat1, False, True)
                    ProgressBar1.Increment(1)
                    Application.DoEvents()
                Next
            Else
                If facs = True Then
                    For ii = 1 To SPLITEST.GetLength(0) - 2
                        Dim contains As Boolean = contiene(spx.nOTUs, 2, SPLITEST(ii, 1))


                        Dim count() As Integer

                        If contains = True Then
                            spalt = Module1.NJproc1(dista.distancias.GetLength(0) - 1, dista.distancias, Nothing, False, SPLITEST(ii, 1), False)
                            count = cladesupport1(spx, spalt, otumat1)

                        Else
                            Dim newa As New newicker
                            Dim split As String = newa.rewritesplits(" " & SPLITEST(ii, 1), dista.distancias.GetLength(0) - 1)
                            spalt = Module1.NJproc1(dista.distancias.GetLength(0) - 1, dista.distancias, Nothing, False, split, True)
                            count = cladesupport1(spalt, spx, otumat1)

                        End If
                        counts(i, ii, 0) = count(0)
                        counts(i, ii, 1) = count(1)
                        ProgressBar1.Increment(1)
                        Application.DoEvents()

                    Next
                End If
            End If

            matr(i) = Module1.NJgo(otumat1, False, False)
            pdis = True
        Next







        stopp = False
        'Dim results(nmax, (SPLITEST.GetLength(0) - 1) * 4) As String
        Dim resultuno(((SPLITEST.GetLength(0) - 1) * 4)) As String
        Dim resultx As New DataTable
        For c = 0 To ((SPLITEST.GetLength(0) - 1) * 4)
            resultx.Columns.Add()
        Next

        For i = 1 To nmax

            textito = Nothing
            If an(1) = 0 Then
                For j = 1 To nmin
                    an(j) = j
                Next

            Else


                If an(nmin) < nlocus Then
                    an(nmin) = an(nmin) + 1
                Else
                    For j = 1 To nmin
                        If an(nmin - j) < nlocus - j Then
                            an(nmin - j) = an(nmin - j) + 1
                            For k = nmin - j + 1 To nmin
                                an(k) = an(k - 1) + 1
                            Next


                            j = nmin


                        End If


                    Next
                End If
            End If
            arch = Nothing
            Dim matf As dsx = Nothing

            ReDim matf.distancias(matr(1).distancias.GetLength(0) - 1, matr(1).distancias.GetLength(0) - 1)


            For k = 1 To an.Length - 1
                For r = 1 To matr(1).distancias.GetLength(0) - 1
                    For s = 1 To matr(1).distancias.GetLength(0) - 1
                        matf.distancias(r, s) = matf.distancias(r, s) + matr(an(k)).distancias(r, s)

                    Next
                    matf.distancias(r, 0) = r
                    matf.distancias(0, r) = r
                Next

                'a(k) = locus(an(k))
                If arch = Nothing Then
                    arch = GridView1.Item(0, an(k) - 1).Value & ", "
                Else
                    arch = arch & GridView1.Item(0, an(k) - 1).Value & ", "
                End If
            Next

            Dim sp As splits


            sp = Module1.NJproc(matf.distancias.GetLength(0) - 1, matf.distancias, otumat1, False, 0, matf.Vij, False)
            If facs = True Then
                For bb = 1 To SPLITEST.GetLength(0) - 2
                    For cc = 1 To sp.nOTUs.GetLength(0) - 2
                        If sp.nOTUs(cc, 2) = SPLITEST(bb, 1) Then
                            Dim count As Integer = 0
                            Dim count1 As Integer = 0
                            For dd = 1 To an.Length - 1
                                count = count + counts(an(dd), bb, 0)
                                count1 = count1 + counts(an(dd), bb, 1)
                            Next
                            sp.notus1(cc, 1) = 1 - combinatF1(0.5, count1, count)
                        End If
                    Next
                Next
            End If
            sp.DSTarr = DSTs(SPLITEST, matf.distancias)
            sp.splitest = SPLITEST

            If nrep > 0 Then
                bootstrapingmatrix(matb, nrep, sp, an)

            End If



            '''''''

            evaluate(sp, resultuno, savetrees)
            '''''''
            'writefiles(sp, ToolStripComboBox1.SelectedIndex + 1, savetrees, True)

            ProgressBar2.Increment(1)
            t = t + 1
            Label20.Text = "Evaluated trees: " & t & "/" & ProgressBar2.Maximum
            If stopp = True Then

                Exit Sub
            End If
            Application.DoEvents()
            resultx.LoadDataRow(resultuno, True)
l20:
        Next





        bs.DataSource = resultx
        DataGridView4.DataSource = Nothing
        DataGridView4.DataSource = bs
        bs.Filter = Nothing
        'DataGridView4.DataSource

        'Dim bs As BindingSource = New BindingSource()
        'bs.DataSource = results
        'DataGridView4.DataSource = bs
        writefiles(SPLITEST)

        CheckBox2.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        CheckBox3.Checked = False
        'RichTextBox1.SaveFile(pathtosave & "\outputrees.nex", RichTextBoxStreamType.PlainText)
        'RichTextBox2.SaveFile(FolderBrowserDialog1.SelectedPath & "\output.txt", RichTextBoxStreamType.PlainText)

        ProgressBar2.Value = 0
        Panel3.Enabled = True

        Panel4.Enabled = True
    End Sub 'Makes all the combination of selected sequences, makes the trees and write the results in a table
    Sub bootstrapingmatrix(ByVal mat(,) As dsx, ByVal nrep As Integer, ByVal sp As splits, ByVal a() As Integer)

        Dim b As Integer = mat(a(1), 1).distancias.GetLength(0) - 1

        Dim rndrep As Integer
        For v = 1 To nrep
            Dim matf As dsx = Nothing
            ReDim matf.distancias(b, b)
            For k = 1 To a.Length - 1
                rndrep = Math.Round(Rnd() * (nrep - 1))
                For r = 1 To b
                    For s = 1 To b
                        matf.distancias(r, s) = matf.distancias(r, s) + mat(a(k), rndrep).distancias(r, s)

                    Next
                    matf.distancias(r, 0) = r
                    matf.distancias(0, r) = r
                Next
            Next
            Dim spb As splits


            spb = Module1.NJproc(matf.distancias.GetLength(0) - 1, matf.distancias, otumat1, False, 0, matf.Vij, False)

            For i = 1 To sp.nOTUs.GetLength(0) - 2

                For j = 1 To spb.nOTUs.GetLength(0) - 2
                    If sp.nOTUs(i, 2) = spb.nOTUs(j, 2) And spb.notus1(j, 0) <> 0 Then
                        sp.notus1(i, 1) = sp.notus1(i, 1) + 1
                        Exit For
                    End If
                Next
            Next
        Next

    End Sub
    Sub evaluate(ByVal sp As splits, ByVal resultuno() As String, ByVal savetrees As Boolean)
        Dim M As Integer = 0
        Dim meanboo As Single
        Dim test As Boolean = True
        For ii = 1 To sp.splitest.GetLength(0) - 2

            If sp.splitest(ii, 1) <> "" Then
                Dim maxerror As Double

                test = False

                For j = 1 To sp.nOTUs.GetLength(0) - 1
                    If sp.splitest(ii, 1) = sp.nOTUs(j, 2) Then
                        If sp.notus1(j, 0) > 0.05 Then
                            test = True






                            resultuno(((ii - 1) * 4) + 3) = "M"
                            resultuno(((ii - 1) * 4) + 4) = sp.notus1(j, 1)
                            resultuno(((ii - 1) * 4) + 5) = sp.notus1(j, 0)

                            meanboo = meanboo + sp.notus1(j, 1)
                            M = M + 1
                            Exit For

                        End If
                    Else
                        Dim splitA As String
                        Dim splitT As String
                        Dim newa As New newicker
                        splitA = newa.rewritesplits(sp.nOTUs(j, 1), sp.otumat1.GetLength(0) - 1)
                        splitT = newa.rewritesplits(sp.splitest(ii, 0), sp.otumat1.GetLength(0) - 1)

                        If sp.notus1(j, 0) > 0.05 AndAlso checkcomp(splitA, splitT, sp.otumat1.GetLength(0) - 1) = False Then

                            test = True
                            Dim dx As Double = 0

                            resultuno(((ii - 1) * 4) + 3) = "I"
                            resultuno(((ii - 1) * 4) + 4) = dx
                            resultuno(((ii - 1) * 4) + 5) = dx
                            Exit For
                        End If
                    End If

                Next
                If test = False Then
                    Dim dx As Double = 0

                    resultuno(((ii - 1) * 4) + 3) = "P"
                    resultuno(((ii - 1) * 4) + 4) = dx
                    resultuno(((ii - 1) * 4) + 5) = dx




                End If

                resultuno(((ii - 1) * 4) + 6) = sp.DSTarr(ii)
            End If

        Next

        resultuno(1) = M / (sp.splitest.GetLength(0) - 2)
        resultuno(2) = meanboo / (sp.splitest.GetLength(0) - 2)
        resultuno(0) = sp.DSTarr(sp.DSTarr.GetLength(0) - 1)
        resultuno(resultuno.Length - 2) = arch
        If savetrees = True Then
            Dim nwk As New newicker
            Dim aa As New treeviewer2
            resultuno(resultuno.Length - 1) = nwk.makenwk(sp.notus1, sp.nOTUs, sp.otumat1.GetLength(0) - 1, sp.otumat1, aa.makeSbMat(sp, sp.nOTUs.GetLength(0) - 1).nodo)
        End If
    End Sub
    Private Sub stepwiseadd(ByVal nrep As Integer, ByVal savetrees As Boolean)
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim countuncheck = 0
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
            If GridView1.Item(2, i - 1).Value = False Then
                countuncheck = countuncheck + 1
            End If
        Next

        stopp = False
        ProgressBar2.Value = 0
        ProgressBar2.Maximum = countuncheck
        Dim a() As String
        Dim fi As Integer = 1


        TabControl1.SelectTab(TabPage9)
        Panel3.Enabled = False
        Panel4.Visible = False
        i = 1
        Dim t As Integer = 0

        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        Dim splitest(,) As String
        splitest = Module1.splitestmaker
        Dim resultuno(((splitest.GetLength(0) - 1) * 4)) As String
        Dim resultx As New DataTable
        For c = 0 To ((splitest.GetLength(0) - 1) * 4)
            resultx.Columns.Add()
        Next


        Dim fixed As String
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)

                If fixed = Nothing Then
                    fixed = GridView1.Item(0, i - 1).Value & ", "
                Else
                    fixed = fixed & GridView1.Item(0, i - 1).Value & ", "
                End If

                fi = fi + 1
            End If
        Next

        i = 1
        Array.Resize(a, fi + 1)
        For i = 1 To ch.Length - 1
            arch = Nothing
            If ch(i) = False Then
                a.SetValue(locus(i), fi)

                If arch = Nothing Then
                    arch = fixed & GridView1.Item(0, i - 1).Value & ", "
                Else
                    arch = arch & GridView1.Item(0, i - 1).Value & ", "
                End If
                t = t + 1
                Dim sp As splits
                sp = Module1.NJ(conca(a, fi, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, False, False, False)
                evaluate(sp, resultuno, savetrees)
                resultx.LoadDataRow(resultuno, True)

                textito = Nothing
                ' writefiles(sp, ToolStripComboBox1.SelectedIndex + 1, True, True)
                ProgressBar2.Increment(1)
                Application.DoEvents()
                If stopp = True Then

                    Exit Sub
                End If
            End If
        Next
        nlocmin = a.GetLength(0) - 1

        bs.DataSource = resultx
        DataGridView4.DataSource = Nothing
        DataGridView4.DataSource = bs
        bs.Filter = Nothing

        writefiles(splitest)


        CheckBox2.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        CheckBox3.Checked = False

        ProgressBar2.Value = 0
        Panel3.Enabled = True
        restore()
    End Sub 'add one of each unselected loci to the selected loci and test the tree 
    Private Sub stepwiseDEL(ByVal nrep As Integer, ByVal savetrees As Boolean, ByVal all As Boolean, ByVal rel As Boolean)
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim countcheck = 0
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
            If GridView1.Item(2, i - 1).Value = True Then
                countcheck = countcheck + 1
            End If
        Next

        stopp = False
        ProgressBar2.Value = 0
        ProgressBar2.Maximum = countcheck - 1
        Dim a() As String
        Dim fi As Integer = 1


        TabControl1.SelectTab(TabPage9)
        Panel3.Enabled = False
        Panel4.Visible = False
        i = 1
        Dim t As Integer = 1


        Dim SPLITEST As String(,)

        If all = True Then
            Dim sp As splits
            ListBox4.Items.Clear()
            DataGridView7.EndEdit()

            If DataGridView7.RowCount <> 0 Then
                For dd = 0 To DataGridView7.RowCount - 1
                    If DataGridView7.Item(1, dd).Value = True Then
                        ListBox4.Items.Add(DataGridView7.Item(0, dd).Value)
                    End If
                Next
            End If
            sp = Module1.NJ(conca(locus, locus.Length - 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)
            SPLITEST = makesplitestfromtree(sp.nOTUs, sp.notus1)

            DataGridView8.RowCount = SPLITEST.GetLength(0) - 2

            For i = 0 To DataGridView8.RowCount - 1

                Dim strains As String = Nothing
                For v = 1 To ListBox4.Items.Count
                    If SPLITEST(i + 1, 0).Contains(" " & v & " ") Then
                        strains = strains & ListBox4.Items(v - 1) & ", "
                    End If
                Next
                DataGridView8.Item(0, i).Value = "split " & i + 1
                DataGridView8.Item(1, i).Value = strains
                DataGridView8.Item(2, i).Value = SPLITEST(i + 1, 0)
                DataGridView8.Item(3, i).Value = SPLITEST(i + 1, 1)

            Next
            externo = True
        Else
            SPLITEST = Module1.splitestmaker

            externo = False
        End If
        Dim resultuno(((SPLITEST.GetLength(0) - 1) * 4)) As String
        Dim resultx As New DataTable
        For c = 0 To ((SPLITEST.GetLength(0) - 1) * 4)
            resultx.Columns.Add()
        Next



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next


        Dim fixed As String
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                fi = fi + 1
            End If
        Next

        i = 1
        Array.Resize(a, fi - 1)
        For i = 1 To ch.Length - 1
            arch = Nothing
            t = 1
            If ch(i) = True Then
                Dim mas As String
                For j = 1 To ch.Length - 1
                    If ch(j) = True And j <> i Then
                        If arch = Nothing Then
                            arch = GridView1.Item(0, j - 1).Value & ", "
                        Else
                            arch = arch & GridView1.Item(0, j - 1).Value & ", "
                        End If
                        a(t) = GridView1.Item(1, j - 1).Value
                        t = t + 1

                    End If

                Next
                arch = arch & "-" & GridView1.Item(0, i - 1).Value & ","

                Dim sp As splits
                sp = Module1.NJ(conca(a, fi - 2, TextBox12.Text, TextBox13.Text, avstates, False), SPLITEST, nrep, False, False, False)


                textito = Nothing
                '  writefiles(sp, ToolStripComboBox1.SelectedIndex + 1, True, True)
                evaluate(sp, resultuno, savetrees)
                resultx.LoadDataRow(resultuno, True)
                ProgressBar2.Increment(1)
                Application.DoEvents()
                If stopp = True Then

                    Exit Sub
                End If
            End If
        Next
        nlocmin = a.GetLength(0) - 1
        bs.DataSource = resultx
        DataGridView4.DataSource = Nothing
        DataGridView4.DataSource = bs
        bs.Filter = Nothing
        'DataGridView4.DataSource

        'Dim bs As BindingSource = New BindingSource()
        'bs.DataSource = results
        'DataGridView4.DataSource = bs
        writefiles(SPLITEST)


        CheckBox2.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        CheckBox3.Checked = False


        ProgressBar2.Value = 0
        Panel3.Enabled = True
        restore()
    End Sub 'add one of each unselected loci to the selected loci and test the tree
    Sub comparetofull(ByVal avstates As Boolean, ByVal nrep As Integer)

        Dim splitest(,) As String = Module1.splitestmaker
        Dim fi As Integer = 1
        Dim fixed As String
        Dim locus(GridView1.RowCount) As String
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value

        Next
        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
        Next

        For i = 1 To ch.Length - 1
            If ch(i) = True Then

                fi = fi + 1
            End If
        Next
        If fi = 1 Then Exit Sub

        Dim resultfull(((splitest.GetLength(0) - 1) * 4)) As String

        Dim spf As splits
        Dim b(fi - 1) As String
        Dim bb As Integer = 1
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                b(bb) = locus(i)
                bb = bb + 1
            End If
        Next
        spf = Module1.NJ(conca(b, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False), splitest, nrep, False, False, False)
        evaluate(spf, resultfull, False)
        Dim dt As New DataTable
        Dim bsx As BindingSource = DirectCast(DataGridView4.DataSource, BindingSource)


        dt = DirectCast(bsx.DataSource, DataTable)
        For h = 0 To dt.Rows.Count - 1
            For i = 4 To resultfull.Length - 5
                Dim a As Double = dt.Rows(h).Item(i)
                dt.Rows(h).Item(i) = Math.Round((a - resultfull(i)) * 10) / 10
                i = i + 3
            Next
        Next


    End Sub 'Bootstrap comparation by adding or deleting one locus
    Sub exporto()
        Dim dt As New DataTable
        Dim dt1 As New DataSet
        Dim bsx As BindingSource = DirectCast(DataGridView4.DataSource, BindingSource)


        dt = DirectCast(bsx.DataSource, DataTable)
        Dim ds As New DataSet
        ds.Tables.Add(dt)





        If externo = True Then
            Dim g As Integer = DataGridView8.RowCount - 1
            For i = 0 To g

                dt.Rows.Add(DataGridView8.Item(0, g - i).Value, DataGridView8.Item(1, g - i).Value)

            Next
        Else

            Dim g As Integer = DataGridView6.RowCount - 1
            For i = 0 To g

                dt.Rows.Add(DataGridView6.Item(0, g - i).Value, DataGridView6.Item(1, g - i).Value)
            Next
        End If

        Dim xx As String





        Application.DoEvents()
        SaveFileDialog1.Title = "Export output"
        xx = "Format " & "xml" & "|" & "*." & "xml"
        SaveFileDialog1.Filter = xx
        SaveFileDialog1.DefaultExt = "xml"
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName <> "" Then
            ds.WriteXml(SaveFileDialog1.FileName)
            Array.Resize(tests, tests.Length + 1)
            tests(tests.Length - 1) = SaveFileDialog1.FileName
            ToolStripComboBox2.Items.Add(tests(tests.Length - 1).Remove(0, tests(tests.Length - 1).LastIndexOf("\"c) + 1))
            ds.Dispose()
        End If
        RichTextBox2.Text = Nothing
    End Sub 'export to a file the output table of tests
    Sub treeinfo()
        Dim ci As Integer = DataGridView4.CurrentCell.RowIndex
        Dim texto As String
        Dim c As String = TextBox13.Text
        Dim i As Integer

        texto = "Tree Information:" & DataGridView4.Item(DataGridView4.ColumnCount - 2, ci).Value & c & c
        texto = texto & "Number of total DSTs: " & DataGridView4.Item(0, ci).Value & c & c

        Dim m As Single
        m = DataGridView4.Item(1, ci).Value

        texto = texto & "proportion of Monophyletic groups: " & m & c & c

        For i = 1 To (DataGridView4.ColumnCount - 4) / 4
            If externo = False Then
                texto = texto & "Group " & i & " Name: " & DataGridView6.Item(0, i - 1).Value & c
                texto = texto & "Strains: " & DataGridView6.Item(1, i - 1).Value & c
            Else
                Dim a As String = DataGridView8.Item(1, i - 1).Value
                texto = texto & "Group " & i & " Name: " & DataGridView8.Item(0, i - 1).Value & c
                texto = texto & "Strains: " & DataGridView8.Item(1, i - 1).Value & c
            End If
            If DataGridView4.Item(((i - 1) * 4) + 3, ci).Value = "M" Then
                texto = texto & "Existing Group: Yes" & c
                texto = texto & "Bootstrap: " & DataGridView4.Item(((i - 1) * 4) + 4, ci).Value & c
                texto = texto & "Branch Length: " & DataGridView4.Item(((i - 1) * 4) + 5, ci).Value & c
            Else
                texto = texto & "Existing Group: No" & c
            End If
            texto = texto & "Number of ST: " & DataGridView4.Item(((i - 1) * 4) + 6, ci).Value & c
            texto = texto & c & c
        Next
        TextBox10.Text = texto

    End Sub 'Gives information of the tree selected in the output table
    Sub writefiles(ByVal splitest(,) As String)





        Dim nmax As Integer = DataGridView4.RowCount - 1
        Dim ncol As Integer = DataGridView4.ColumnCount - 1

        DataGridView4.Columns(ncol).Visible = False
        DataGridView4.Columns(ncol - 1).Visible = True
        'DataGridView4.Rows.Add()
        DataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView4.Columns.Item(0).HeaderCell.Value = "Number of ST"
        DataGridView4.Columns.Item(1).HeaderCell.Value = "% of Monophyly"
        DataGridView4.Columns.Item(2).HeaderCell.Value = "Mean Support"
        DataGridView4.Columns.Item(ncol - 1).HeaderCell.Value = "Loci"
        DataGridView4.Columns.Item(ncol - 1).ReadOnly = True

        CheckedListBox1.Items.Clear()

        If externo = True Then

            For i = 0 To DataGridView8.RowCount - 1
                CheckedListBox1.Items.Add("G " & i + 1 & ": " & DataGridView8.Item(0, i).Value, True)
            Next
        Else


            For i = 0 To DataGridView6.RowCount - 1
                CheckedListBox1.Items.Add("G " & i + 1 & ": " & DataGridView6.Item(0, i).Value, True)
            Next
        End If


        For i = 1 To splitest.GetLength(0) - 2

            If splitest(i, 1) <> "" Then


                DataGridView4.Columns.Item(((i - 1) * 4) + 3).HeaderCell.Value = "G " & i
                If externo = True Then
                    DataGridView4.Columns(((i - 1) * 4) + 3).HeaderCell.ToolTipText = Replace(DataGridView8.Item(1, i - 1).Value, ",", vbNewLine)
                Else
                    DataGridView4.Columns(((i - 1) * 4) + 3).HeaderCell.ToolTipText = Replace(DataGridView6.Item(1, i - 1).Value, ",", vbNewLine)
                End If

                DataGridView4.Columns.Item(((i - 1) * 4) + 4).HeaderCell.Value = "Support"
                DataGridView4.Columns.Item(((i - 1) * 4) + 5).HeaderCell.Value = "br.le"
                DataGridView4.Columns.Item(((i - 1) * 4) + 6).HeaderCell.Value = "ST"


            End If

        Next













        ''



    End Sub 'writes the table of tests
    Sub unsellowboots(ByVal bionj As Boolean, ByVal minboo As Single, ByVal nrep As Integer, ByVal selsplit As Boolean, ByVal order As Boolean)
        Dim avstates As Boolean
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim arch As String = Nothing
        Dim a(1) As String
        Dim fi As Integer = 1
        Dim t As Integer = 0

        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next


        i = 1



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        ProgressBar2.Maximum = ch.Length - 1
        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then


                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If


            End If
        Next

        i = 1

        Dim sp(nmax - 1) As splits
        Dim boots(nmax - 1) As Single
        Dim j As Integer
        Dim splitest(,) As String
        If selsplit = True Then
            splitest = Module1.splitestmaker
        End If
        For i = 1 To ch.Length - 1
            'This block make the splits tree for individual selected loci and save the splits in sp (splits) structure
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                arch1 = arch
                If arch1 = Nothing Then
                    arch1 = GridView1.Item(0, i - 1).Value
                Else
                    arch1 = arch1 & "_" & GridView1.Item(0, i - 1).Value
                End If
                t = t + 1

                sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, bionj, False, False)
                'Makes NJ trees
                Dim count As Integer = 0
                If selsplit = True Then
                    For g = 1 To splitest.GetLength(0) - 2
                        For f = 1 To sp(j).notus1.GetLength(0) - 2

                            If sp(j).nOTUs(f, 2) = splitest(g, 1) Then
                                If sp(j).notus1(f, 0) > 0.0000000001 Or sp(j).notus1(f, 0) < -0.0000000001 Then
                                    boots(j) = boots(j) + sp(j).notus1(f, 1)

                                End If

                            End If

                        Next
                        count = count + 1
                    Next
                Else
                    For f = 1 To sp(j).notus1.GetLength(0) - 2


                        If sp(j).notus1(f, 0) > 0.0000000001 Or sp(j).notus1(f, 0) < -0.0000000001 Then
                            boots(j) = boots(j) + sp(j).notus1(f, 1)

                        End If


                        count = count + 1
                    Next



                End If
                boots(j) = boots(j) / count
                ProgressBar2.Increment(1)

                j = j + 1
                Application.DoEvents()

            End If
        Next
        If order = False Then

            For s = 0 To boots.Length - 1
                If boots(s) >= minboo Then
                    GridView1.Item(2, s).Value = True
                Else
                    GridView1.Item(2, s).Value = False


                End If
                GridView1.Item(0, s).ToolTipText = "Mean Bootstrap:" & boots(s)
            Next
        Else
            Dim boots1(boots.Length - 1) As Single
            Dim arr(GridView1.RowCount - 1, 2) As String
            Dim index As Integer
            For i = 0 To boots.Length - 1
                index = Array.IndexOf(boots, boots.Max)
                boots1(i) = boots(index)
                boots(index) = -1
                arr(i, 0) = GridView1.Item(0, index).Value
                arr(i, 1) = GridView1.Item(1, index).Value

            Next

            For i = 0 To boots.Length - 1
                GridView1.Item(0, i).Value = arr(i, 0)
                GridView1.Item(1, i).Value = arr(i, 1)
                GridView1.Item(0, i).ToolTipText = "Mean Bootstrap:" & boots1(i)
            Next
        End If

    End Sub 'select loci by bootstrap support

    Private Sub ToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        If DataGridView6.SelectedCells.Count <> 0 Then
            For i = 0 To DataGridView6.SelectedCells.Count - 1
                Dim cellpos As Integer = DataGridView6.SelectedCells.Item(0).RowIndex
                DataGridView6.Rows.RemoveAt(cellpos)

            Next
        End If
    End Sub
    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        Dim a As Integer = DataGridView4.ColumnCount - 1

        If CheckBox5.Checked = False Then
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                DataGridView4.Columns.Item((i * 4) + 5).Visible = False

            Next
        Else
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                If CheckedListBox1.CheckedIndices.Contains(i) = True Then
                    DataGridView4.Columns.Item((i * 4) + 5).Visible = True
                End If
            Next
        End If
    End Sub
    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Dim a As Integer = DataGridView4.ColumnCount - 1

        If CheckBox4.Checked = False Then
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                DataGridView4.Columns.Item((i * 4) + 4).Visible = False

            Next
        Else
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                If CheckedListBox1.CheckedIndices.Contains(i) = True Then
                    DataGridView4.Columns.Item((i * 4) + 4).Visible = True
                End If
            Next
        End If
    End Sub
    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        Dim a As Integer = DataGridView4.ColumnCount - 1

        If CheckBox6.Checked = False Then
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                DataGridView4.Columns.Item((i * 4) + 6).Visible = False

            Next
        Else
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                If CheckedListBox1.CheckedIndices.Contains(i) = True Then
                    DataGridView4.Columns.Item((i * 4) + 6).Visible = True
                End If
            Next
        End If
    End Sub
    Private Sub DataGridView7_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellClick
      
        Dim count As Integer

        For i = 0 To DataGridView7.RowCount - 1
            If DataGridView7.Item(1, i).Value = True Then
                count = count + 1
            End If
        Next
        'TextBox12.Text = Module1.nseqx(GridView1.Item(1, 0).Value) + 1
        TextBox12.Text = count

    End Sub
    Private Sub DataGridView7_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellEndEdit
        actseqs()
    End Sub

    ''minimun set output window
    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        Clipboard.SetDataObject(DataGridView4.GetClipboardContent())
    End Sub 'context menu. copy the selected cells to the clipboard
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim max As Double = 0

        Dim s As String

        If ListBox5.SelectedIndices.Contains(0) Then 'filter by monophyly



            For i = 0 To DataGridView4.RowCount - 1
                If DataGridView4.Rows(i).Visible = True Then


                    Dim m As Double = DataGridView4.Item(1, i).Value


                    If m >= max Then
                        max = m
                        s = DataGridView4.Item(1, i).Value
                    End If
                End If
            Next

            If bs.Filter = Nothing Then
                bs.Filter = "Column2 like '" & s & "'"

            Else
                bs.Filter = bs.Filter & " AND " & "Column2 like '" & s & "'"
            End If

        End If
        '---
        If ListBox5.SelectedIndices.Contains(1) Then 'Filtering by Support
            If TextBox5.Text = "" Then
                MsgBox("You have to type a value for Bootstrap cutoff")
                Exit Sub
            End If
            Dim dt As New DataTable
            Dim bsx As BindingSource = DirectCast(DataGridView4.DataSource, BindingSource)


            dt = DirectCast(bsx.DataSource, DataTable)

            Dim x As Integer = dt.Rows.Count - 1
            For i = 0 To x

                Dim m As Integer
                m = 0
                For j = 4 To DataGridView4.ColumnCount - 3
                    If dt.Rows(i).Item(j) <> -1 Then
                        Dim SS As Single = dt.Rows(i).Item(j)
                        If SS >= TextBox5.Text Then
                            m = m + 1
                        End If
                    End If

                    j = j + 3

                Next
                dt.Rows(i).Item(2) = m
                If m >= max Then
                    max = m
                    s = dt.Rows(i).Item(2)
                End If
                'End If
            Next
            Dim bs1 As New BindingSource
            bs1.DataSource = dt
            bs1.Filter = bs.Filter
            If bs1.Filter = Nothing Then
                bs1.Filter = "Column3 like '" & s & "'"

            Else
                bs1.Filter = bs1.Filter & " AND " & "Column3 like '" & s & "'"
            End If

            DataGridView4.DataSource = bs1
            DataGridView4.Columns(2).HeaderText = "# groups higher than cutof"
        End If
        '---

        If ListBox5.SelectedIndices.Contains(2) Then 'Filtering by Number of DSTs

            max = 0
            For i = 0 To DataGridView4.RowCount - 1
                If DataGridView4.Rows(i).Visible = True Then
                    Dim m As Integer
                    m = 0
                    m = DataGridView4.Item(0, i).Value
                    If m >= max Then
                        max = m
                        s = DataGridView4.Item(0, i).Value
                    End If
                End If
            Next

            If bs.Filter = Nothing Then
                bs.Filter = "Column1 like '" & s & "'"

            Else
                bs.Filter = bs.Filter & " AND " & "Column1 like '" & s & "'"
            End If

        End If


    End Sub 'view button in loci selection, It makes filtering of output results
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        bs.Filter = Nothing
        DataGridView4.DataSource = bs
    End Sub 'Delete filters in combinations and shows hidden combinations
    Private Sub CheckedListBox1_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        Dim a As Integer = DataGridView4.ColumnCount - 1
        Try
            If CheckedListBox1.CheckedIndices.Contains(CheckedListBox1.SelectedIndex) = True Then
                Dim i As Integer
                i = CheckedListBox1.SelectedIndex

                DataGridView4.Columns.Item((i * 4) + 3).Visible = False
                DataGridView4.Columns.Item((i * 4) + 4).Visible = False
                DataGridView4.Columns.Item((i * 4) + 5).Visible = False
                DataGridView4.Columns.Item((i * 4) + 6).Visible = False

            Else
                Dim i As Integer
                i = CheckedListBox1.SelectedIndex
                If CheckBox6.Checked = True Then
                    DataGridView4.Columns.Item((i * 4) + 6).Visible = True
                End If
                If CheckBox5.Checked = True Then
                    DataGridView4.Columns.Item((i * 4) + 5).Visible = True
                End If
                If CheckBox4.Checked = True Then
                    DataGridView4.Columns.Item((i * 4) + 4).Visible = True
                End If
                DataGridView4.Columns.Item((i * 4) + 3).Visible = True

            End If
        Catch
        End Try
    End Sub 'filter by group
    Private Sub DataGridView4_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        treeinfo()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ss As MsgBoxResult
        ss = MsgBox("Non visible Combinations will be deleted. Do you want to export them previously?", MsgBoxStyle.YesNo)
        If ss = MsgBoxResult.Yes Then
            exporto()
        End If
        Dim avstates As Boolean
        Dim hethand
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Panel4.Enabled = False
        Dim nmax As Integer = DataGridView4.Rows.GetRowCount(DataGridViewElementStates.Visible) - 1
        Dim nrep As Integer = TextBox1.Text
        ProgressBar1.Visible = True
        ProgressBar1.Maximum = 0
        ProgressBar2.Visible = True
        ToolStripStatusLabel3.Text = "Bootstrapping"
        ToolStripStatusLabel3.Visible = True
        ToolStripStatusLabel4.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        Dim visel(nmax) As Integer
        Dim v As Integer = 0






        Dim locus(GridView1.Rows.Count) As String
        If TextBox1.Text = "" Then
            MsgBox("You must specify the number of bootstrap replications")
            Exit Sub
        End If

        Dim nmin As Integer = nlocmin
        Dim savetrees As Boolean = strees
        Dim a(nmin) As String
        Dim an(nmin) As Integer

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer


        TabControl1.SelectTab(TabPage9)
        Panel3.Enabled = False
        Dim nlocus As Integer = GridView1.Rows.Count


        ProgressBar2.Maximum = nmax
        Dim t As Integer = 0


        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        '''


        Dim matb(,) As dsx
        Dim fast As Boolean = RadioButton2.Checked
        Dim local As Boolean = RadioButton3.Checked
        Dim matsize As Integer

        Dim matr(nlocus) As dsx
        Dim otumat1 As String(,)
        If fast = True Then
            ReDim matb(nlocus, nrep - 1)
            Dim loc() As Boolean = Nothing
            ReDim loc(locus.Length - 1)
            For i = 0 To GridView1.RowCount - 1
                For j = 0 To DataGridView4.RowCount - 1
                    If DataGridView4.Item(DataGridView4.ColumnCount - 2, j).Value.ToString.Contains(GridView1.Item(0, i).Value & ", ") Then
                        loc(i) = True

                    End If
                Next
            Next
            For Each xx As Boolean In loc
                If xx = True Then
                    ProgressBar1.Maximum = ProgressBar1.Maximum + nrep
                End If
            Next
            For i = 1 To locus.Length - 1
                If loc(i - 1) = True Then

                    locus(i) = GridView1.Item(1, i - 1).Value
                    a(1) = locus(i)
                    pdis = False
                    otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                    lseq = otumat1(1, 1).Length - 1

                    For x = 0 To nrep - 1
                        matb(i, x) = NJgo(otumat1, False, True)

                        ProgressBar1.Increment(1)
                        Application.DoEvents()
                    Next


                    matr(i) = Module1.NJgo(otumat1, False, False)
                    matsize = matr(i).distancias.GetLength(0) - 1
                    pdis = True
                End If

            Next
        End If
        '''
        stopp = False
        Dim splitest(,) As String
        If externo = True Then
            splitest = Module1.splitestmakerEX
        Else
            splitest = Module1.splitestmaker
        End If
        Dim resultuno(((splitest.GetLength(0) - 1) * 4)) As String
        Dim resultx As New DataTable
        For c = 0 To ((splitest.GetLength(0) - 1) * 4)
            resultx.Columns.Add()
        Next

        For i = 0 To DataGridView4.Rows.Count - 1
            Dim pos As Integer = 1
            For s = 0 To GridView1.RowCount - 1
                Dim aaa As String = DataGridView4.Item(DataGridView4.ColumnCount - 2, i).Value
                Dim bbb As String = GridView1.Item(0, s).Value & ", "
                If aaa.StartsWith(bbb) Then
                    If aaa.Contains(bbb) = True Then

                        an(pos) = s + 1
                        pos = pos + 1
                    End If
                Else
                    If aaa.Contains(", " & bbb) = True Then

                        an(pos) = s + 1
                        pos = pos + 1
                    End If
                End If

            Next
            arch = Nothing
            For k = 1 To an.Length - 1

                a(k) = locus(an(k))
                If arch = Nothing Then
                    arch = GridView1.Item(0, an(k) - 1).Value & ", "
                Else
                    arch = arch & GridView1.Item(0, an(k) - 1).Value & ", "
                End If
            Next
            Dim sp As splits
            If fast = False Then
                pdis = False
                sp = Module1.NJ(conca(a, nmin, TextBox12.Text, TextBox13.Text, avstates, False), splitestmaker1, nrep, False, local, True)
                pdis = True

            Else

                Dim matf As dsx = Nothing

                ReDim matf.distancias(matsize, matsize)


                For k = 1 To an.Length - 1
                    For r = 1 To matr(an(1)).distancias.GetLength(0) - 1
                        For s = 1 To matr(an(1)).distancias.GetLength(0) - 1
                            matf.distancias(r, s) = matf.distancias(r, s) + matr(an(k)).distancias(r, s)

                        Next
                        matf.distancias(r, 0) = r
                        matf.distancias(0, r) = r
                    Next


                Next




                sp = Module1.NJproc(matf.distancias.GetLength(0) - 1, matf.distancias, otumat1, False, 0, matf.Vij, False)

                sp.DSTarr = DSTs(splitest, matf.distancias)
                sp.splitest = splitest


                bootstrapingmatrix(matb, nrep, sp, an)


            End If



            evaluate(sp, resultuno, False)

            CheckBox3.Checked = True
            ProgressBar2.Increment(1)
            If stopp = True Then
                Panel3.Enabled = True
                Panel4.Enabled = True
                Exit Sub
            End If
            Application.DoEvents()

            resultx.LoadDataRow(resultuno, True)
            If stopp = True Then
                restore()
                Exit Sub
            End If
        Next
        bs.Filter = Nothing
        DataGridView4.DataSource = Nothing
        bs.DataSource = Nothing
        bs.DataSource = resultx
        DataGridView4.DataSource = bs
        writefiles(splitest)
        CheckBox2.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        CheckBox3.Checked = False
        ProgressBar2.Value = 0
        Panel3.Enabled = True
        Panel4.Enabled = True
        restore()

    End Sub 'Calculates support for visible combinations
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        Dim a As Integer = DataGridView4.ColumnCount - 1




        If CheckBox2.Checked = True Then
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            CheckBox6.Enabled = True
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                DataGridView4.Columns.Item((i * 4) + 3).Visible = True

            Next
        Else
            CheckBox4.Enabled = False
            CheckBox4.Checked = True
            CheckBox4.Checked = False
            CheckBox5.Enabled = False
            CheckBox5.Checked = True
            CheckBox5.Checked = False
            CheckBox6.Enabled = False
            CheckBox6.Checked = True
            CheckBox6.Checked = False
            Dim i As Integer
            For i = 0 To (a - 7) / 4
                DataGridView4.Columns.Item((i * 4) + 3).Visible = False

            Next
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            DataGridView4.Columns(2).Visible = True
        Else
            DataGridView4.Columns(2).Visible = False
        End If
    End Sub
    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click

        If RichTextBox1.SelectedText = Nothing Then
            MsgBox("You have to select (Paint) loci names", , "MLSTest1")



        Else
            If GridView1.Rows.Count > 0 Then
                For i = 0 To GridView1.Rows.Count - 1
                    If RichTextBox1.SelectedText.Contains(GridView1.Item(0, i).Value & ",") = True Then
                        GridView1.Item(2, i).Value = True
                    Else
                        GridView1.Item(2, i).Value = False
                    End If

                Next
            End If
        End If
    End Sub
    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        If RichTextBox1.SelectedText <> Nothing Then
            Clipboard.SetText(RichTextBox1.SelectedText)
        End If
    End Sub
    Private Sub DataGridView4_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellEnter
        If DataGridView4.SelectedCells.Count > 0 Then
            treeinfo()
        End If
    End Sub
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = False Then
            TextBox1.Enabled = True
            TextBox1.Text = 100
        Else
            TextBox1.Enabled = False
            TextBox1.Text = 0
        End If
    End Sub

    'Trees menu
    Private Sub NJToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NJToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Neighbor Joining Analysis", .page = 1, ._supp = True, ._bionj = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            CancelButton1.Visible = True
            blockmenusytab()
            TabControl1.Enabled = False
            Dim testsplits As Boolean
            If a.viewnmin = True Then
                Dim respuesta As MsgBoxResult
                respuesta = MsgBox("Do you want to test only selected groups?, Click NO to test all nodes into the tree", MsgBoxStyle.YesNo, "MLSTest1.0 - 1-njCS")
                If respuesta = MsgBoxResult.Yes Then
                    testsplits = True
                End If
            End If
            NJtree(False, a._supp, False, a.nrep, a._bionj, a.viewnmin, testsplits, False)
            'TabControl1.Enabled = True
        End If
    End Sub 'Concatenated NJ button
    Private Sub UPGMAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPGMAToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "UPGMA-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "UPGMA Analysis", .page = 1, ._supp = True}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Bootstrapping"
            CancelButton1.Visible = True
            blockmenusytab()
            UPGMAtree(False, a._supp, a.nrep)
        End If
    End Sub 'Concatenated UPGMA button
    Private Sub ConsensusTreeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConsensusTreeToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Consensus NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim ab As MsgBoxResult
        ab = MsgBox("Do you want to use bioNJ instead of NJ algorithm?", MsgBoxStyle.YesNo, "Consensus NJ-Analisis")
        Dim bionj As Boolean
        If ab = MsgBoxResult.Yes Then
            bionj = True
        End If
        ProgressBar2.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()


        Cons(True, bionj)
        restore()

    End Sub 'Consensus NJ button
    Private Sub UPGMAConsensusTreeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPGMAConsensusTreeToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Consensus UPGMA-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If

        ProgressBar2.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        Cons(False, False)
        restore()

    End Sub 'Consensus UPGMA button
    Private Sub NJToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NJToolStripMenuItem1.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Consensus NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim ab As MsgBoxResult
        ab = MsgBox("Do you want to use bioNJ instead of NJ algorithm?", MsgBoxStyle.YesNo, "Consensus NJ-Analisis")
        Dim bionj As Boolean
        If ab = MsgBoxResult.Yes Then
            bionj = True
        End If
        ProgressBar2.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()


        MRe(bionj)
        restore()



    End Sub 'MRe button
    Private Sub LoadMLSTestTreeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadMLSTestTreeToolStripMenuItem.Click
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "MLSTest tree (*.xml)|*.xml"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName <> "" Then


            leeroutputtree(OpenFileDialog1.FileName)



        End If
    End Sub 'Load MLSTest saved tree

    ''subs of Trees menu
    Private Sub NJtree(ByVal perfi As Boolean, ByVal nlsupp As Boolean, ByVal incong As Boolean, ByVal nrep As Integer, ByVal bionj As Boolean, ByVal facs As Boolean, ByVal facsplt As Boolean, ByVal detincong As Boolean)
        Dim vs As Boolean

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        textito = Nothing
        Dim arch As String = Nothing
        ProgressBar2.Value = 0


        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0
        'SaveFileDialog1.ShowDialog()
        'Dim pathtosave As String = SaveFileDialog1.FileName


        i = 1


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1 - nmax


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)


                fi = fi + 1
            End If
        Next



        i = 1
        Dim sp As splits
        If perfi = False Then



            sp = Module1.NJ(conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, bionj, facs, facsplt)


        Else
            Dim perfiles(,) = Perfal.perf(a, TextBox12.Text, locusname, True)
            otumat1 = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
            sp = Module1.NJperfal(perfiles, otumat1, arch, nrep)

        End If


        Application.DoEvents()
        Dim supnames() As String

        If nrep > 0 Then
            vs = True
            ReDim supnames(0)
            supnames(supnames.Length - 1) = "Bootstrap"

        End If
        If facs = True Then
            vs = True
            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "1-njCS"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "1-njCS"
            End If
        End If
        If nlsupp = True Then
            Dim bar As Boolean
            If nrep > 0 Or facs = True Then
                bar = True
            End If
            vs = True

            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "#loci"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "#loci"
            End If
            nlSupport(sp, a, avstates)
        End If

        If incong = True Then
            compatibility(sp, a, avstates)

            vs = True
            ReDim supnames(0)
            supnames(0) = "incongruent loci"
          
        End If
        If detincong = True Then
            relaxedIncong(sp, 0, True, False)
            Dim x As Integer
            ReDim supnames(0)
            vs = True
            For g = 0 To GridView1.RowCount - 1
                If GridView1.Item(2, g).Value = True Then
                    Array.Resize(supnames, x + 1)
                    supnames(x) = GridView1.Item(0, x).Value
                    x = x + 1
                End If
            Next

        End If
        ' writefiles(sp, ToolStripComboBox1.SelectedIndex + 1, True, False)
        'Dim nwkform As New treeviewer1 With {.nwktree = textito, .Text = "Tree Viewer"}


        Dim nwkform As New treeviewer2 With {.sptree = sp, .Text = "Tree Viewer", ._viewsupport = vs, ._supportnames = supnames} '
        nwkform.Show()
        restore()

        'RichTextBox1.SaveFile(pathtosave, RichTextBoxStreamType.PlainText)
    End Sub 'Makes a NJ tree of selected loci (concatenated) and shows the tree in the treeviewer
    Sub nlSupport(ByRef sp As splits, ByVal a() As String, ByVal avstates As Boolean)
        Dim spi(a.Length - 1) As splits

        For i = 1 To a.Length - 1
            Dim an(1) As String
            an(1) = a(i)
            spi(i) = NJ(conca(an, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)

        Next
        testnlsupport(sp, spi)
    End Sub

    Function testnlsupport(ByRef sp As splits, ByRef spi As splits())
        Dim merr As Double = 0.05 / sp.lseq
        Dim newa As New newicker
        Dim N As Integer = sp.otumat1.GetLength(0) - 1


        'For i = N + 1 To sp.nOTUs.GetLength(0) - 1
        ''sp.nOTUs(i, 0) = newa.rewritesplits(sp.nOTUs(i, 1), N)
        '
        'Next
        'For i = 1 To spi.Length - 1


        'For ii = N + 1 To sp.nOTUs.GetLength(0) - 1
        'spi(i).nOTUs(ii, 0) = newa.rewritesplits(spi(i).nOTUs(ii, 1), N)
        '
        'Next
        'Next

        For j = sp.otumat1.GetLength(0) To sp.nOTUs.GetLength(0) - 1
            If sp.notus1(j, 0) > merr Then

                Dim nlsupp As Integer = 0
                For k = 1 To spi.Length - 1
                    Dim merr2 As Double = 0.05 / spi(k).lseq
                    For m = spi(1).otumat1.GetLength(0) To spi(k).nOTUs.GetLength(0) - 1
                        If spi(k).notus1(m, 0) > merr2 Then
                            If sp.nOTUs(j, 2) = spi(k).nOTUs(m, 2) Then
                                nlsupp = nlsupp + 1
                                Exit For
                            End If
                        End If

                    Next
                Next
                If sp.notus1(j, 1) = Nothing Then
                    sp.notus1(j, 1) = nlsupp
                Else
                    sp.notus1(j, 1) = sp.notus1(j, 1) & "/" & nlsupp
                End If

            End If
            If j = sp.otumat1.GetLength(0) - 1 Then
                Dim vrc As String = Nothing
                If sp.notus1(j, 0) > merr Or sp.notus1(j, 0) < -merr Then



                    vrc = sp.notus1(j, 1) & "/"

                End If
                sp.notus1(j, 1) = vrc & "0"

            End If
        Next

    End Function

    Sub compatibility(ByVal sp As splits, ByVal a() As String, ByVal avstates As Boolean)
        Dim spi(a.Length - 1) As splits

        For i = 1 To a.Length - 1
            Dim an(1) As String
            an(1) = a(i)
            spi(i) = NJ(conca(an, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)

        Next
        testcompatib(sp, spi, -1)


    End Sub ' calculates the number of imcompatible splits for each branch of the concatenated tree
    Function testcompatib(ByRef sp As splits, ByRef spi As splits(), ByVal sup As Single)
        Dim merr As Double = 0.05 / sp.lseq
        Dim newa As New newicker
        Dim N As Integer = sp.otumat1.GetLength(0) - 1


        For i = N + 1 To sp.nOTUs.GetLength(0) - 1
            sp.nOTUs(i, 0) = newa.rewritesplits(sp.nOTUs(i, 1), N)

        Next
        For i = 1 To spi.Length - 1


            For ii = N + 1 To sp.nOTUs.GetLength(0) - 1
                spi(i).nOTUs(ii, 0) = newa.rewritesplits(spi(i).nOTUs(ii, 1), N)

            Next
        Next

        For j = sp.otumat1.GetLength(0) To sp.nOTUs.GetLength(0) - 2
            If sp.notus1(j, 0) > merr Then

                Dim incomp As Integer = 0
                For k = 1 To spi.Length - 1
                    Dim merr2 As Double = 0.05 / spi(k).lseq
                    For m = spi(1).otumat1.GetLength(0) To spi(k).nOTUs.GetLength(0) - 2
                        If spi(k).notus1(m, 0) > merr2 Then
                            If spi(k).notus1(m, 1) > sup AndAlso checkcomp(sp.nOTUs(j, 0), spi(k).nOTUs(m, 0), sp.otumat1.GetLength(0) - 1) = False Then
                                incomp = incomp + 1
                                Exit For
                            End If
                        End If

                    Next
                Next
                If sp.notus1(j, 1) = Nothing Then
                    sp.notus1(j, 1) = incomp
                Else
                    sp.notus1(j, 1) = sp.notus1(j, 1) & "/" & incomp
                End If

            End If
        Next

    End Function

    Sub compatibilityA(ByVal sp As splits, ByVal a() As String, ByVal avstates As Boolean)
        Dim spi(a.Length - 1) As splits
        Dim totalsites As Integer

        For i = 1 To a.Length - 1
            Dim an(1) As String
            an(1) = a(i)

            spi(i).otumat1 = conca(an, 1, TextBox12.Text, TextBox13.Text, avstates, False)

            spi(i) = NJ(spi(i).otumat1, Module1.splitestmaker, 0, False, False, False)
            spi(i).nsites = lseq
            totalsites = totalsites + lseq
        Next
        For j = sp.otumat1.GetLength(0) To sp.nOTUs.GetLength(0) - 2
            sp.notus1(j, 1) = 1
            If sp.notus1(j, 0) > 10 ^ -9 Then
                Dim dsx1 As dsx = NJgo(conca(a, a.Length - 1, TextBox12.Text, TextBox13.Text, avstates, False), False, False)
                Dim spalt As splits
                spalt = NJproc1(sp.otumat1.GetLength(0) - 1, dsx1.distancias, Nothing, False, sp.nOTUs(j, 2), False)
                Dim totalfav, totaldesfav As Integer
                totalfav = 0
                totaldesfav = 0

                Dim c(a.Length - 1) As Integer
                'Array.Clear(c, 0, 2)
                Dim incomp As Integer = 0
                For k = 1 To spi.Length - 1
                    Dim an(1) As String
                    an(1) = a(k)



                    Dim b(1) As Integer
                    Dim d() As Integer

                    'If checkcomp(sp.nOTUs(j, 2), spi(k).nOTUs(m, 2), sp.otumat1.GetLength(0) - 1) = False Then
                    'incomp = incomp + 1


                    d = cladesupport1(sp, spalt, spi(k).otumat1)

                    c(k) = d(0)
                    totalfav = totalfav + d(0)
                    totaldesfav = totaldesfav + d(1) - d(0)

                    'End If
                    'Else

                    ' End If


                Next
                totalsites = 500 * (a.Length - 1)
                Dim pfav = totalfav / totalsites
                Dim pdesfav = totalfav / totalsites
                Dim modelP(c.Max) As Single
                Dim emp(c.Max) As Single
                For s = 0 To c.Max
                    Dim count As Integer = 0
                    For i = 1 To c.Length - 1
                        If c(i) = s Then
                            count = count + 1
                        End If
                    Next
                    modelP(s) = combinatF2(pdesfav, 500, s)
                    emp(s) = count / (a.Length - 1)
                Next





                sp.notus1(j, 1) = LRT(modelP, emp, a.Length - 1)

                'End If
            End If
        Next

    End Sub 'code for Beta testing
    Function LRT(ByVal model() As Single, ByVal emp() As Single, ByVal k As Integer) As Single
        Dim sumLN As Single = 0
        Dim df As Integer = -1
        For i = 0 To model.Length - 1
            If emp(i) <> 0 Then
                sumLN = sumLN + k * emp(i) * Math.Log(model(i) / emp(i))
                df = df + 1
            End If

        Next
        If df > 0 Then

        Else
            Return 1
        End If

    End Function 'code for Beta testing
    Sub compatibilityB(ByVal sp As splits, ByVal a() As String, ByVal avstates As Boolean)
        Dim spi(a.Length - 1) As splits


        Dim splitsx(a.Length - 1, 99) As splits
        pdis = False
        For x = 0 To 99
            For y = 1 To a.Length - 1
                Dim an(1) As String
                an(1) = a(y)
                otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                otumat1 = reductor(otumat1)
                lseq = otumat1(1, 1).Length

                Dim dix As dsx = NJgo(otumat1, False, True)
                splitsx(y, x) = NJproc(otumat1.GetLength(0) - 1, dix.distancias, Nothing, False, 0, Nothing, False)
            Next
        Next
        pdis = True
        For i = 1 To a.Length - 1
            Dim an(1) As String
            an(1) = a(i)
            spi(i) = NJ(conca(an, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)
        Next
        Dim count As Integer


        For j = 1 To sp.nOTUs.GetLength(0) - 2

            If sp.notus1(j, 0) > 10 ^ -9 Then

                Dim incomp As Integer = 0
                For k = 1 To a.Length - 1
                    Dim supp As Integer = 0
                    For m = 1 To spi(k).nOTUs.GetLength(0) - 2
                        If spi(k).notus1(m, 0) > 10 ^ -9 Then

                            If checkcomp(sp.nOTUs(j, 2), spi(k).nOTUs(m, 2), sp.otumat1.GetLength(0) - 1) = False Then

                                For XX = 0 To 99
                                    For mm = 1 To spi(k).nOTUs.GetLength(0) - 2
                                        If splitsx(k, XX).notus1(m, 0) > 10 ^ -9 Then
                                            If checkcomp(sp.nOTUs(j, 2), splitsx(k, XX).nOTUs(mm, 2), sp.otumat1.GetLength(0) - 1) = False Then
                                                supp = supp + 1
                                                Exit For
                                            End If
                                        End If
                                    Next

                                Next
                                Exit For
                            End If
                        End If

                    Next
                    If sp.notus1(j, 1) < supp Then
                        sp.notus1(j, 1) = supp
                    End If
                Next



            End If
        Next


    End Sub 'code for Beta testing it calculates the booststrap support for incongruences
    Private Sub UPGMAtree(ByVal perfi As Boolean, ByVal nlsupp As Boolean, ByVal nrep As Integer)
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        textito = Nothing

        ProgressBar2.Value = 0


        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0




        i = 1



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1 - nmax


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)



                fi = fi + 1
            End If
        Next



        i = 1
        Dim sp As splits
        If perfi = False Then
            sp = Module2.UPGMA(conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False), arch, nrep)

        Else
            Dim perfiles(,) = Perfal.perf(a, TextBox12.Text, locusname, True)
            sp = Module2.UPGMAperfal(perfiles, arch, nrep)
        End If
        ''''
        Dim vs As Boolean
        Dim supnames() As String

        If nrep > 0 Then
            vs = True
            ReDim supnames(0)
            supnames(supnames.Length - 1) = "Bootstrap"

        End If
        
        If nlsupp = True Then
            Dim bar As Boolean
            If nrep > 0 Then
                bar = True
            End If
            vs = True

            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "#loci"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "#loci"
            End If
            ''''

            Dim sp2 As splits
            sp2 = consensus(0, False, False) 'make all the splits for individual tree to estimate number of trees support 
            For i = 1 To (sp.nOTUs.GetLength(0) - 1)
                For j = 1 To sp2.nOTUs.GetLength(0) - 1
                    If sp.nOTUs(i, 2) = sp2.nOTUs(j, 2) Then

                        If sp.notus1(i, 0) > 10 ^ -9 Or sp.notus1(i, 0) < -10 ^ -9 Then
                            Dim vrc As String = Nothing
                            If bar = True Then

                                vrc = sp.notus1(i, 1) & "/"
                            End If

                            sp.notus1(i, 1) = vrc & sp2.notus1(j, 1)
                        End If
                        Exit For
                    End If
                Next
            Next
        End If
        'writefiles(sp, 0, True, False)

        'Dim nwkform As New treeviewer1 With {.nwktree = textito}

        Dim nwkform As New treeviewer2 With {.sptree = sp, .Text = "Tree Viewer", ._upgma = True, ._viewsupport = vs, ._supportnames = supnames}
        If nrep > 0 Or nlsupp = True Then
            nwkform._viewsupport = True
        End If
        nwkform.Show()
        restore()



    End Sub 'Makes an UPGMA tree of selected loci (concatenated) and shows the tree in the treeviewer
    Private Sub Cons(ByVal nj As Boolean, ByVal bionj As Boolean)
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next



        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1


        i = 1

        Dim sp(nmax - 1) As splits
        Dim j As Integer
        For i = 1 To ch.Length - 1

            If ch(i) = True Then
                a.SetValue(locus(i), 1)

                If nj = True Then
                    sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, bionj, False, False)

                Else
                    sp(j) = Module2.UPGMA(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), arch, 0)

                End If


                'RichTextBox1.SaveFile(pathtosave & "\" & arch1 & ".nex", RichTextBoxStreamType.PlainText)
                ProgressBar2.Increment(1)

                j = j + 1
                Application.DoEvents()

            End If
        Next
        Dim sp1 As splits
        Dim spp As New Consensus1
        sp1 = spp.consensussplits(sp, False)


        If nj = True Then
            'ff = nwk.makenwk(sp1.notus1, sp1.nOTUs, Toolstripcombobox1.SelectedIndex + 1, nseq + 1, sp1.otumat1, True)

        Else
            'ff = nwk.makenwk(sp1.notus1, sp1.nOTUs, 0, nseq + 1, sp1.otumat1, False)
        End If
        Dim vs As Boolean = True
        Dim supnames(0) As String
        supnames(0) = "Number of loci"
        Dim nwkform As New treeviewer2 With {.sptree = sp1, ._upgma = True, ._viewsupport = vs, ._supportnames = supnames}

        nwkform.Show()

    End Sub 'Makes the consensus tree of individual locus trees (NJ or UPGMA)
    Sub MRe(ByVal bionj As Boolean)
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next



        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1


        i = 1

        Dim sp(nmax - 1) As splits
        Dim j As Integer
        For i = 1 To ch.Length - 1

            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, bionj, False, False)
                ProgressBar2.Increment(1)
                j = j + 1
                Application.DoEvents()

            End If
        Next
        Dim sp1 As splits
        Dim spp As New Consensus1
        sp1 = spp.consensussplits(sp, True)
        Dim supnames(0) As String
        supnames(0) = "support"
        Dim nwkform As New treeviewer2 With {.sptree = sp1, ._upgma = True, ._supportnames = supnames, ._viewsupport = True}
        nwkform.Show()


    End Sub 'Makes the Majority rule extended consensus

    'Others Menu
    Private Sub MultidimensionalScalingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MultidimensionalScalingToolStripMenuItem.Click
        RichTextBox1.Visible = False
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Classical Multidimensional Scaling")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If

        ProgressBar1.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()
        MDSca()
        restore()

    End Sub 'Multidimensional scaling buttonej
    Private Sub ConsensusNetworkToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Consensus Network", .page = 3, .nloc = GridView1.RowCount - 1}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then

            ProgressBar2.Visible = True

            ToolStripStatusLabel4.Visible = True
            CancelButton1.Visible = True
            blockmenusytab()
            nlocmin = a.gdef

            ConsNet(a.gdef)
            restore()

        End If

    End Sub 'Consensus Network button

    'subs of Others menu

    Sub MDSca()
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        ProgressBar1.Value = 0
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next

        Dim arch As String = Nothing



        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0




        i = 1


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)

                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If

                fi = fi + 1
            End If
        Next
        otumat1 = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        nseq = otumat1.GetLength(0) - 2

        lseq = otumat1(1, 1).ToString.Length

        otumat1 = reductor(otumat1)
        lseqred = otumat1.GetValue(1, 1).ToString.Length





        Dim seqs(1) As String
        Dim nseq1 As Integer = nseq + 1
        Dim distancias(nseq1, nseq1) As Double
        Dim j As Integer
        Dim bootstrapn() As Integer
        Dim dist As medirdistancia

        ProgressBar1.Maximum = nseq1 + 4


        Dim sequ1, sequ2 As String

        dist = New medirdistancia


        For i = 1 To nseq1

            For j = i + 1 To nseq1

                If i <> j Then
                    seqs(0) = otumat1(i, 1)
                    seqs(1) = otumat1(j, 1)

                    sequ1 = seqs(0)
                    sequ2 = seqs(1)

                    distancias(i, j) = dist.distance(sequ1, sequ2, lseq, False, 0, 1, bootstrapn)
                    distancias(j, i) = distancias(i, j)

                End If


            Next j
            ProgressBar1.Increment(1)
        Next i
        Dim xydist As Double(,)
        xydist = MDS1.MDS1(distancias)

        DataGridView5.RowCount = xydist.GetLength(0) - 1
        Dim texto As String
        Dim st As Integer = 1
        Dim xypoints(xydist.GetLength(0) - 1, 2) As Double

        For i = 1 To xydist.GetLength(0) - 1
            xypoints(i, 1) = xydist(i, 1)
            xypoints(i, 2) = xydist(i, 2)
            DataGridView5.Item(1, i - 1).Value = otumat1(i, 0)
            DataGridView5.Item(2, i - 1).Value = xydist(i, 1)

            DataGridView5.Item(3, i - 1).Value = xydist(i, 2)

            Dim newst As Boolean = True
            If i = 1 Then
                DataGridView5.Item(0, i - 1).Value = 1
                xypoints(i, 0) = 1
                newst = False
            End If
            For j = 1 To i - 1
                If otumat1(j, 1) = otumat1(i, 1) Then
                    DataGridView5.Item(0, i - 1).Value = DataGridView5.Item(0, j - 1).Value
                    newst = False
                    Exit For

                End If

            Next
            If newst = True Then

                st = st + 1
                DataGridView5.Item(0, i - 1).Value = st
                xypoints(i, 0) = st
            End If

            'texto = texto & xydist(i, 1) & " " & xydist(i, 2) & TextBox13.Text
        Next
        DataGridView5.Columns(0).HeaderText = "ST"
        DataGridView5.Columns(1).HeaderText = "Strains"
        DataGridView5.Columns(2).HeaderText = "X"
        DataGridView5.Columns(3).HeaderText = "Y"
        RichTextBox1.Visible = True
        Dim stat As String
        stat = "x-Axis explains " & Math.Round(xydist(0, 0) * 100, 1) & "% of variability" & vbNewLine & _
        "y-Axis explains " & Math.Round(xydist(0, 1) * 100, 1) & "% of variability" & vbNewLine & _
        "Both Axes explain " & Math.Round((xydist(0, 1) + xydist(0, 0)) * 100, 1) & "% of variability"
        RichTextBox1.Text = stat

        Dim mdsviewer As New treeviewer2 With {.xypoints = xypoints, .Text = "Classical Multidimensional Scaling"}
        mdsviewer.Show()
        TabControl1.SelectTab(TabPage11)

    End Sub 'makes a distance matrix, calls the multidimensional scaling and shows the results in a table and into a viewer
    Private Sub ConsNet(ByVal umbral As Integer)
        Dim avstates As Boolean

        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next

        Dim arch As String = Nothing

        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0

        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1


        i = 1

        Dim sp(nmax - 1) As splits
        Dim j As Integer
        For i = 1 To ch.Length - 1

            If ch(i) = True Then
                a.SetValue(locus(i), 1)


                t = t + 1

                sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)
                'writefiles(sp(j), Toolstripcombobox1.SelectedIndex + 1, False, True)

                ProgressBar2.Increment(1)
                ' Label20.Text = "trees: " & t & "/" & ProgressBar2.Maximum
                j = j + 1
                Application.DoEvents()

            End If
        Next
        Consnetmaker(sp, umbral, -1)

    End Sub
    Function Consnetmaker(ByRef sp() As splits, ByVal umbral As Integer, ByVal minsup As Single)
        Dim sp1 As splits
        Dim spp As New consensusnetwork
        sp1 = spp.consensusnet(sp, umbral, minsup)
        Dim texto As String
        Dim enter As String = TextBox13.Text
        nseq = sp1.otumat1.GetLength(0) - 2
        texto = Nothing
        texto = "#nexus" & enter & enter & "BEGIN Taxa;" & enter & "DIMENSIONS ntax=" & nseq + 1 & ";" & enter & "TAXLABELS" & enter

        For i = 1 To sp1.otumat1.GetLength(0) - 1
            texto = texto & "[" & i & "] '" & sp1.otumat1(i, 0) & "'" & enter
        Next
        texto = texto & ";" & enter & "END; [Taxa]" & enter & "BEGIN Splits;" & enter & "DIMENSIONS ntax=" & nseq + 1 & " nsplits=" & sp1.nOTUs.GetLength(0) - 2 & " ;" & enter & "FORMAT labels=no weights=yes confidences=yes intervals=no;"
        texto = texto & enter & "MATRIX" & enter
        For i = 1 To sp1.nOTUs.GetLength(0) - 2
            texto = texto & sp1.notus1(i, 0) & " " & sp1.notus1(i, 1) & " " & sp1.nOTUs(i, 2) & enter
        Next
        texto = texto & ";" & enter & "END; [Splits]"
        Dim xx As String

        Application.DoEvents()
        SaveFileDialog1.Title = "Export Consensus Splits"
        xx = "Nexus Format" & "|" & "*." & "nex"
        SaveFileDialog1.Filter = xx
        SaveFileDialog1.DefaultExt = "nex"
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName <> "" Then
            Dim save As New StreamWriter(SaveFileDialog1.FileName)
            save.WriteLine(texto)
            save.Close()


        End If
        ProgressBar2.Value = 0
    End Function 'makes individual trees of each loci and then builds the splits for a consensus network with a threshold

    'subs used by different buttons

    Sub restore()

        MenuStrip1.Enabled = True
        ProgressBar1.Visible = False
        ProgressBar1.Value = 0
        ProgressBar2.Value = 0
        ProgressBar2.Visible = False
        ToolStripStatusLabel3.Visible = False
        ToolStripStatusLabel4.Visible = False
        CancelButton1.Visible = False
        TabControl1.Enabled = True
    End Sub 'hide the proggres bars and the cancel button. Enable the menu bar
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        DataGridView7.EndEdit()
        alignmf = Nothing


    End Sub
    Private Sub CancelButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelButton1.Click
        Dim aa As MsgBoxResult
        aa = MsgBox("Do You want to stop the analysis?", MsgBoxStyle.YesNo)
        If aa = MsgBoxResult.Yes Then
            Module1._stopp = True
            Module2._stopp = True
            stopp = True
            restore()
        End If
    End Sub
    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If ListBox6.SelectedIndices.Count <> 0 Then
            Dim m As Integer
            For i = 0 To ListBox6.Items.Count
                If ListBox6.SelectedIndices.Contains(i) Then
                    ListBox6.Items.RemoveAt(i)
                    i = i - 1
                End If

            Next
        End If
    End Sub 'cancel the analysis
    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click
        ListBox4.SelectedIndex = -1
    End Sub
    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "."c Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Private Sub ListBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox5.SelectedIndexChanged
        If ListBox5.SelectedIndex = 1 Then
            Panel2.Visible = True
        Else
            Panel2.Visible = False
        End If
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    'Position and size
    Private Sub Form1_ClientSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ClientSizeChanged
        Dim r As New System.Globalization.CultureInfo("es-ES")
        r.NumberFormat.NumberDecimalSeparator = "."
        System.Threading.Thread.CurrentThread.CurrentCulture = r

        If start = True Then
            If Me.WindowState <> FormWindowState.Minimized Then
                Dim x = Me.Width - iwidth
                Dim y = Me.Height - iHeight
                TabControl1.Width = TabControl1.Width + x
                TabControl1.Height = TabControl1.Height + y
                GridView1.Height = GridView1.Height + y
                GridView1.Width = GridView1.Width + x
                DataGridView7.Height = DataGridView7.Height + y
                DataGridView2.Height = DataGridView2.Height + y
                DataGridView2.Width = DataGridView2.Width + x
                DataGridView5.Height = DataGridView5.Height + y
                DataGridView4.Height = DataGridView4.Height + y
                DataGridView4.Width = DataGridView4.Width + x
                RichTextBox1.Height = RichTextBox1.Height + y
                TextBox10.Height = TextBox10.Height + y
                iwidth = Me.Width
                iHeight = Me.Height
            End If
        End If
    End Sub
    Private Sub Form1_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin
        iHeight = Me.Height
        iwidth = Me.Width
    End Sub
    Private Sub Form1_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Dim x = Me.Width - iwidth
        Dim y = Me.Height - iHeight
        TabControl1.Width = TabControl1.Width + x
        TabControl1.Height = TabControl1.Height + y
        GridView1.Height = GridView1.Height + y
        GridView1.Width = GridView1.Width + x
        DataGridView7.Height = DataGridView7.Height + y
        DataGridView2.Height = DataGridView2.Height + y
        DataGridView2.Width = DataGridView2.Width + x
        DataGridView5.Height = DataGridView5.Height + y
        DataGridView4.Height = DataGridView4.Height + y
        DataGridView4.Width = DataGridView4.Width + x
        TextBox10.Height = TextBox10.Height + y
        RichTextBox1.Height = RichTextBox1.Height + y
        iwidth = Me.Width
        iHeight = Me.Height
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim r As New System.Globalization.CultureInfo("es-ES")
        r.NumberFormat.NumberDecimalSeparator = "."
        System.Threading.Thread.CurrentThread.CurrentCulture = r
        Dim s As String
        Try

            If AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData IsNot Nothing Then
                For i = 0 To AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData.Length - 1
                    s = s & AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData(i) & " "

                Next


                leerproject(s)
                nameofproject = s
            End If
        Catch
        End Try
        TabControl1.SelectTab(TabPage1)



        iwidth = Me.Width
        iHeight = Me.Height
        start = True
        ReDim tests(0)
    End Sub
    Private Sub AboutToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem2.Click
        AboutBox1.Show()
    End Sub 'About

    '-----------------------------------------------------------------------------
    'Functions
    '-----------------------------------------------------------------------------
    Function ponderado(ByVal arrayw() As Integer) As Single
        Dim a As Single
        Dim b As Double = arrayw.Sum()
        a = arrayw.Sum() / (arrayw.Length - 1)
        Return a
    End Function 'it calculates the mean GD
    Function probmax(ByVal arrayw() As Integer) As Single
        Dim a As Single
        Dim b As Double = arrayw.Max
        For i = 1 To arrayw.Length - 1
            If arrayw(i) = b Then
                a = a + 1
            End If




        Next
        a = a / (arrayw.Length - 1)
        Return a
    End Function 'it calculates maximum GD
    Function factorial(ByVal n As Integer) As Double 'return the factorial of n
        Dim a As Double = 1
        For i = 1 To n
            a = i * a

        Next
        Return a
    End Function
    Function checksel() As Integer
        GridView1.EndEdit()

        Dim counta As Integer

        For i = 0 To GridView1.RowCount - 1
            If GridView1.Item(2, i).Value = True Then
                counta = counta + 1
            End If
        Next
        Return counta
    End Function 'counts the number of loci selected in the gridview1
    Function Ndst(ByVal sp As splits) As Integer 'return the number of ST in a split structure

        Dim sum As Integer = 0

        For i = 1 To sp.splitest.GetLength(0) - 2

            If sp.splitest(i, 1) <> "" Then

                Dim j = 1
                Do While j < sp.nOTUs.GetLength(0) - 1
                    If sp.splitest(i, 1) = sp.nOTUs(j, 2) And sp.notus1(j, 0) > 10 ^ -10 Then
                        sum = sum + 1


                    End If

                    j = j + 1
                Loop

            End If

        Next
        Return sum
    End Function 'return the number of ST in a split structure
    Public Function conca(ByVal a() As String, ByVal b As Integer, ByVal nseq As Integer, ByVal c As String, ByVal avstates As Boolean, ByVal all As Boolean) As String(,)
        'reads Fasta files and concatenate the sequences and return a string array with the sequences concatenated and modified
        'a() contain the paths for the files of different loci to be readed and concatenated

        Dim lector As TextReader
        Dim i, j, k, l, m As Integer
        Dim concat(nseq, 1) As String 'this is the returned array

        Dim concat1 As String
        Dim path As String
        Dim n As Integer = 1
        Dim lin As String = 1
        Dim length As Integer
        Dim index As Integer = 0
        For i = 1 To b 'b is the number of loci


            path = a.GetValue(i)

            lector = New StreamReader(path)
            n = 0
            lin = 1
            Dim ax As Integer = 0
            If all = True Then 'if the function is called for load the names of strain showed in the datagridview7
                Do While ax < 2 'This loop read all the sequence for the file specified in path

                    lin = lector.ReadLine()
                    If lin <> Nothing Then
                        If lin.Substring(0, 1) = ">" Then

                            concat(n + 1, 0) = lin.Substring(1, lin.Length - 1)
                            n = n + 1


                            index = index + 1
                            ax = 0
                        Else

                            concat(n, 1) = concat(n, 1) & lin.ToUpper




                        End If
                    Else
                        ax = ax + 1
                    End If

                Loop

            Else 'when the function is called for an analysis



                Do While ax < 2

                    lin = lector.ReadLine()
                    If lin <> Nothing Then
                        If lin.Substring(0, 1) = ">" Then
                            If DataGridView7.Item(1, index).Value = True Then
                                concat(n + 1, 0) = DataGridView7.Item(0, index).Value
                                n = n + 1

                            End If
                            index = index + 1
                            ax = 0
                        Else
                            If DataGridView7.Item(1, index - 1).Value = True Then 'check if the strain is selected for the analysis
                                concat(n, 1) = concat(n, 1) & lin.ToUpper
                            End If


                        End If
                    Else
                        ax = ax + 1
                    End If


                Loop
            End If
            index = 0
        Next i
        lector.Close()
        concat1 = 1




        Return concat
    End Function 'return a string array with the sequences concatenated and modified
    Function consensus(ByVal nrep As Integer, ByVal nj As Boolean, ByVal bionj As Boolean) As splits
        'Return the consensus splits with umbral=0 from splits of different trees
        Dim avstates As Boolean
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim arch As String = Nothing
        Dim a(1) As String
        Dim fi As Integer = 1
        Dim t As Integer = 0

        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next


        i = 1



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        ProgressBar2.Maximum = ch.Length - 1
        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then


                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If


            End If
        Next

        i = 1

        Dim sp(nmax - 1) As splits
        Dim j As Integer
        For i = 1 To ch.Length - 1
            'This block make the splits tree for individual selected loci and save the splits in sp (splits) structure
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                arch1 = arch
                If arch1 = Nothing Then
                    arch1 = GridView1.Item(0, i - 1).Value
                Else
                    arch1 = arch1 & "_" & GridView1.Item(0, i - 1).Value
                End If
                t = t + 1

                If nj = True Then

                    sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, bionj, False, False)
                    'Makes NJ trees
                Else
                    sp(j) = Module2.UPGMA(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), arch, nrep)
                    'Makes UPGMA
                End If
                ProgressBar2.Increment(1)

                j = j + 1
                Application.DoEvents()

            End If
        Next
        Dim sp1 As splits
        Dim spp As New consensusnetwork
        sp1 = spp.consensusnet(sp, 0, -1) 'return a unique list of splits in split structure
        Return sp1
    End Function 'return a unique list of splits in a split structure of a group trees made by NJ or UPGMA
    Sub relaxedIncong(ByVal spconca As splits, ByVal nrep As Integer, ByVal nj As Boolean, ByVal bionj As Boolean)
        'Return the consensus splits with umbral=0 from splits of different trees
        Dim avstates As Boolean
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        Dim arch As String = Nothing
        Dim a(1) As String
        Dim fi As Integer = 1
        Dim t As Integer = 0

        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If

        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next


        i = 1



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next

        ProgressBar2.Maximum = ch.Length - 1
        Dim arch1 As String

        For i = 1 To ch.Length - 1
            If ch(i) = True Then


                If arch = Nothing Then
                    arch = GridView1.Item(0, i - 1).Value
                Else
                    arch = arch & "_" & GridView1.Item(0, i - 1).Value
                End If


            End If
        Next

        i = 1

        Dim sp(nmax - 1) As splits
        Dim j As Integer
        For i = 1 To ch.Length - 1
            'This block make the splits tree for individual selected loci and save the splits in sp (splits) structure
            If ch(i) = True Then
                a.SetValue(locus(i), 1)

                t = t + 1

                If nj = True Then

                    sp(j) = Module1.NJ(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, bionj, False, False)
                    'Makes NJ trees
                Else
                    sp(j) = Module2.UPGMA(conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False), arch, nrep)
                    'Makes UPGMA
                End If
                ProgressBar2.Increment(1)

                j = j + 1
                Application.DoEvents()

            End If
        Next
        Dim sp1 As splits
        Dim n1 As Integer = spconca.otumat1.GetLength(0)
        Dim n2 As Integer = spconca.nOTUs.GetLength(0) - 2
        Dim nwk As New newicker

        For i = n1 To n2
            Dim count As Integer = 0
            Dim splitc As String = nwk.rewritesplits(" " & spconca.nOTUs(i, 2), n1 - 1)
            For k = 0 To nmax - 1
                count = 0
                For j = n1 To n2
                    If sp(k).notus1(j, 0) > 10 ^ -9 Then
                        Dim spliti As String = nwk.rewritesplits(" " & sp(k).nOTUs(j, 2), n1 - 1)

                        If checkcomp(splitc, spliti, n1 - 1) = False Then
                            count = count + 1
                        End If
                    End If
                Next
                If spconca.notus1(i, 1) <> Nothing Then
                    spconca.notus1(i, 1) = spconca.notus1(i, 1) & "/" & Math.Round(count)
                Else
                    spconca.notus1(i, 1) = Math.Round(count)
                End If
            Next
            'spconca.notus1(i, 1) = Math.Round(count / nmax, 2)

        Next

    End Sub 'return a unique list of splits in a split structure of a group trees made by NJ or UPGMA
    Function resolve(ByVal base As String) As String()
        'Resolves ambiguous bases in the posible bases
        Dim res(1) As String
        Select Case base
            Case "R"
                res(0) = "G"
                res(1) = "A"
            Case "S"
                res(0) = "G"
                res(1) = "C"
            Case "K"
                res(0) = "G"
                res(1) = "T"
            Case "W"
                res(0) = "T"
                res(1) = "A"
            Case "M"
                res(0) = "C"
                res(1) = "A"
            Case "Y"
                res(0) = "T"
                res(1) = "C"
        End Select
        Return res
    End Function 'return a string array of the two posible bases of an heterozygote site 
    Private Function makesplitestfromtree(ByVal notus As String(,), ByVal notus1(,) As String) As String(,)
        Dim a As String()
        Dim k As Integer = 2
        For i = 1 To notus.GetLength(0) - 1
            If treeviewer2.splitsize(notus(i, 1)) > 1 Then
                If notus1(i, 0) > 0.0000000000001 Or notus1(i, 0) < -0.0000000000001 Then
                    Array.Resize(a, k)
                    a(k - 1) = notus(i, 1)
                    k = k + 1
                End If
            End If
        Next
        Dim b(a.Length - 1, 1) As String
        For i = 1 To a.Length - 1

            b(i, 0) = a(i)
            b(i, 1) = stdsplit(a(i), TextBox12.Text)
        Next
        Return b
    End Function 'it makes a serie of splits to test from a tree
    Function splitestmaker1() As String(,)
        Dim splitest(,) As String
        If externo = False Then
            ReDim splitest(DataGridView6.RowCount + 1, 1)
            For i = 1 To DataGridView6.RowCount
                splitest(i, 1) = DataGridView6.Item(2, i - 1).Value
                splitest(i, 0) = DataGridView6.Item(3, i - 1).Value

            Next
        Else
            ReDim splitest(DataGridView8.RowCount + 1, 1)
            For i = 1 To DataGridView8.RowCount
                splitest(i, 1) = DataGridView8.Item(3, i - 1).Value
                splitest(i, 0) = DataGridView8.Item(2, i - 1).Value

            Next
        End If
        Return splitest
    End Function 'return an string array of splits to be tested in an analysis writed by the user 
    Function leeroutput(ByVal file As String)
        'allows to read the output saved from a loci selection analysis

        Dim ds As New DataSet
        ds.ReadXml(file)
        DataGridView8.Rows.Clear()
        With ds.Tables(0)
            Dim n As Integer

            If .Rows(.Rows.Count - 1).Item(2).ToString <> "" Then
                MsgBox("the file format is not from MLSTest", MsgBoxStyle.OkOnly)
                Exit Function
            End If
            Do While .Rows(.Rows.Count - 1).Item(2).ToString = ""


                DataGridView8.Rows.Add()



                DataGridView8.Item(0, n).Value = .Rows(.Rows.Count - 1).Item(0)
                DataGridView8.Item(1, n).Value = .Rows(.Rows.Count - 1).Item(1)
                .Rows.RemoveAt(.Rows.Count - 1)
                n = n + 1
            Loop
        End With


        Dim i As Integer = 0




        bs.DataSource = ds.Tables(0)
        DataGridView4.DataSource = bs

        DataGridView4.Columns(DataGridView4.ColumnCount - 1).Visible = False
        DataGridView4.Columns(DataGridView4.ColumnCount - 2).ReadOnly = True
        DataGridView4.Columns.Item(0).HeaderCell.Value = "Number of ST"
        DataGridView4.Columns.Item(1).HeaderCell.Value = "Proportion of Monophyly"
        DataGridView4.Columns.Item(2).HeaderCell.Value = "Mean Support"
        For i = 1 To DataGridView8.RowCount
            DataGridView4.Columns.Item(((i - 1) * 4) + 3).HeaderCell.Value = "G " & i
            DataGridView4.Columns.Item(((i - 1) * 4) + 4).HeaderCell.Value = "Support"
            DataGridView4.Columns.Item(((i - 1) * 4) + 5).HeaderCell.Value = "br.le"
            DataGridView4.Columns.Item(((i - 1) * 4) + 6).HeaderCell.Value = "ST"
            DataGridView4.Columns(((i - 1) * 4) + 3).HeaderCell.ToolTipText = Replace(DataGridView8.Item(1, i - 1).Value, ",", vbNewLine)
        Next
        CheckedListBox1.Items.Clear()
        For i = 0 To DataGridView8.RowCount - 1
            CheckedListBox1.Items.Add("G " & i + 1 & ": " & DataGridView8.Item(0, i).Value, True)
        Next
        TabControl1.SelectTab(TabPage9)
        externo = True
        CheckBox2.Checked = True
        CheckBox2.Checked = False
        If DataGridView4.Item(2, 0).Value > 0 Then
            CheckBox3.Checked = False
            CheckBox3.Checked = True
        Else
            CheckBox3.Checked = True
            CheckBox3.Checked = False

        End If


    End Function 'allows to read the output saved from a loci selection analysis
    Function ILD_bionj(ByVal nper As Integer)
        'It calculates pvalue for ILD-Bionj test
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False
        ProgressBar1.Visible = True
        ToolStripStatusLabel3.Text = "Permuting"
        CancelButton1.Visible = True
        blockmenusytab()


        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next



        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1


        i = 1

        Dim sp(nmax - 1) As splits
        Dim j As Integer
        '---------------------------------------
        'Makes individual trees
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                a.SetValue(locus(i), 1)

                Dim concatenate(,) As String
                concatenate = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                sp(j) = Module1.NJ(concatenate, Module1.splitestmaker, 0, True, False, False)
                sp(j).nsites = concatenate(1, 1).Length 'determine the number of polymorphic sites
                ProgressBar2.Increment(1)
                j = j + 1
                Application.DoEvents()
            End If
        Next
        '---------------------------------------------------
        'Makes concatenated Tree
        Dim sp1 As splits
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)


                fi = fi + 1
            End If
        Next
        Dim concat(,) As String = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        concat = reductor(concat)
        sp1 = Module1.NJ(concat, Module1.splitestmaker, 0, True, False, False)
        sp1.nsites = concat(1, 1).Length

        '---------------------------------------------------
        'Calculate the tree lengths

        Dim concat1(,) As String = concat.Clone

        Dim sumL, sumLc, sumaLp, ILD, ILDp As Double
        For i = 0 To sp.Length - 1
            For j = 0 To sp(i).notus1.GetLength(0) - 1
                If sp(i).notus1(j, 0) > 10 ^ -9 Then
                    sumL = sumL + sp(i).notus1(j, 0)

                End If
            Next
        Next
        For j = 0 To sp1.notus1.GetLength(0) - 1
            If sp1.notus1(j, 0) > 10 ^ -9 Then
                sumLc = sumLc + sp1.notus1(j, 0)

            End If
        Next
        ILD = sumLc - sumL 'represents the extralength
        '------------------------------------------------------

        'permutation test
        Dim count As Integer
        ProgressBar1.Maximum = nper
        ProgressBar1.Value = 0
        Dim h As Integer
        For h = 1 To nper
            concat = Nothing
            concat = concat1.Clone
            sumaLp = 0
            Dim countsp1 As Integer = 0
            For i = 0 To nmax - 1
                Dim conc1(concat.GetLength(0) - 1, concat.GetLength(1) - 1) As String
                For j = 0 To sp(i).nsites - 1
                    Dim pos As Integer = Math.Round(Rnd() * (concat(1, 1).Length - 1))
                    For k = 1 To conc1.GetLength(0) - 1
                        conc1(k, 1) = conc1(k, 1) & concat(k, 1).Substring(pos, 1)

                        concat(k, 1) = concat(k, 1).Remove(pos, 1)
                    Next
                Next

                Dim spp As splits

                If conc1(1, 1) <> Nothing Then
                    spp = Module1.NJ(conc1, Module1.splitestmaker, 0, True, False, False)
                    For j = 0 To spp.notus1.GetLength(0) - 1
                        If spp.notus1(j, 0) > 10 ^ -9 Then
                            sumaLp = sumaLp + spp.notus1(j, 0)
                        End If
                    Next
                End If
            Next

            ILDp = sumLc - sumaLp
            If sumL >= sumaLp - 0.002 * sumaLp Then
                count = count + 1
            End If

            ProgressBar1.Increment(1)
            Application.DoEvents()
            If stopp = True Then
                Dim pval1 = 1 - (count / h)
                MsgBox("the ILDp is " & pval1)
                stopp = False
                Exit Function

            End If
        Next
        '-----------------------------------------------------------

        Dim pval = ((count + 1) / (h))
        MsgBox("the ILDp is " & pval)
        pdis = True

    End Function 'it calculates ILD-bionj Test
    Function support_conc(ByVal nrep As Integer, ByVal selsplits As Boolean, ByVal splits(,) As String) As Single()
        'Calculates NJ-ILD for each branch on a tree
        ProgressBar1.Maximum = nrep
        stopp = False
        If nrep = 0 Then ProgressBar1.Visible = False

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0
        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        Dim selocus(nmax - 1) As String
        Dim p As Integer
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                selocus(p) = GridView1(1, i - 1).Value
                p = p + 1
            End If
        Next

        i = 1

        '-------------------------------------------------------
        'makes individual alignments
        Dim sp(nmax - 1) As splits
        Dim j As Integer = 0
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                Dim concatenate(,) As String
                concatenate = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                concatenate = Module1.reductor(concatenate)
                sp(j).nsites = concatenate(1, 1).Length
                ProgressBar2.Increment(1)
                j = j + 1
                Application.DoEvents()
            End If
        Next
        Dim sp1 As splits
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)
                fi = fi + 1
            End If
        Next
        '-----------------------------------------------

        'Makes concatenated tree
        Dim concat(,) As String = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        
        Dim treelengths(nmax - 1) As Single
        sp1 = Module1.NJ(concat, Module1.splitestmaker, 0, False, False, False)
        Dim newa As New newicker
        For i = 1 To sp1.nOTUs.GetLength(0) - 1
            sp1.nOTUs(i, 0) = newa.rewritesplits(sp1.nOTUs(i, 1), sp1.otumat1.GetLength(0) - 1)
        Next
        a = Nothing

        '-------------------------------------------------------------
        'Makes individual trees and calculates lengths
        ReDim a(1)
        ReDim Module1.otum(nmax - 1)
        For i = 0 To nmax - 1
            otum(i).treelength = 0
            Dim x As Integer = hethandpp
            a(1) = GridView1.Item(1, i).Value
            otum(i).otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
            hethand = 0
            otum(i).dista = distanciasx(otum(i).otumat1, otum(i).otumat1.GetLength(0) - 1)

            Dim ssp As splits
            ssp = Module1.NJ(otum(i).otumat1, Module1.splitestmaker, 0, False, False, False)
            For SS = 1 To ssp.notus1.GetLength(0) - 1
                otum(i).treelength = otum(i).treelength + ssp.notus1(SS, 0)
            Next
            treelengths(i) = otum(i).treelength
            hethand = x
        Next
        '---------------------------------------------------------------

        'calculates extralengths for individual tree when a branch is forced to exist in a partition
        Dim extralengths() As Single
        extralengths = Module1.checkxtrasteps(sp1.nOTUs, 0, concat.GetLength(0) - 2, False, Nothing, treelengths, splits, selsplits)
        sp1.nsites = concat(1, 1).Length

        '--------------------------------------------------------------

        'Permutation test
        Dim concat1(,) As String = concat.Clone
        Dim matcount(sp1.nOTUs.GetLength(0) - 1) As Integer
        Dim count As Integer
        Dim countsp As Integer
        For h = 1 To nrep
            Dim sp2 As splits
            concat = Nothing
            concat = concat1.Clone
            ReDim Module1.otum(nmax - 1)
            For i = 0 To nmax - 1
                Dim conc1(concat.GetLength(0) - 1, concat.GetLength(1) - 1) As String

                For j = 0 To sp(i).nsites - 1
                    Dim pos As Integer = Math.Round(Rnd() * (concat(1, 1).Length - 1))

                    For k = 1 To conc1.GetLength(0) - 1
                        conc1(k, 1) = conc1(k, 1) & concat(k, 1).Substring(pos, 1)

                        concat(k, 1) = concat(k, 1).Remove(pos, 1)
                    Next
                Next

                a(1) = selocus(i)
                otum(i).otumat1 = conc1
                If otum(i).otumat1(1, 1) = Nothing Then
                    GoTo 23
                End If
                otum(i).dista = distanciasx(otum(i).otumat1, otum(i).otumat1.GetLength(0) - 1)
                Dim ssp As splits
                ssp = Module1.NJ(otum(i).otumat1, Module1.splitestmaker, 0, False, False, False)
                For SS = 1 To ssp.notus1.GetLength(0) - 1
                    otum(i).treelength = otum(i).treelength + ssp.notus1(SS, 0)
                Next

                If stopp = True Then
                    Exit Function
                End If

23:         Next

            Dim extralengths1() As Single
            extralengths1 = Module1.checkxtrasteps(sp1.nOTUs, 0, concat.GetLength(0) - 2, False, extralengths, treelengths, splits, selsplits)
            For ss = 1 To extralengths.Length - 1

                Dim xtc As Single = extralengths1(ss)
                If extralengths(ss) > xtc Then
                    matcount(ss) = matcount(ss) + 1
                End If

            Next
            Application.DoEvents()
            ProgressBar1.Increment(1)
        Next

        If nrep = 0 Then
            For f = 1 To sp1.notus1.GetLength(0) - 1
                sp1.notus1(f, 1) = Math.Round(((extralengths(f))) * 1000) / 1000
            Next
        Else
            Dim R(1) As Single
         
            Dim COUNTTP As Single = 0.05
            Dim COUNTFP, countBc As Integer
            Dim bcorr, nbcorr As Boolean

            '''
            For f = sp1.otumat1.GetLength(0) To sp1.notus1.GetLength(0) - 2
                If extralengths(f) > 0 Then
                    countBc = countBc + 1
                End If
            Next
            '''
            For f = sp1.otumat1.GetLength(0) To sp1.notus1.GetLength(0) - 2
                sp1.notus1(f, 1) = (1 - (matcount(f)) / (nrep))

                If sp1.notus1(f, 1) < (0.05 / countBc) Then
                    R(1) = R(1) + 1
                    R(0) = R(0) + 1
                ElseIf sp1.notus1(f, 1) < (0.05) Then
                    R(0) = R(0) + 1
                End If
                


            Next
            If bcorr = True Then

            End If
            '
            support_conc = R
        End If
        Dim supnames(0) As String
        supnames(0) = "p-Val"
        'treeviewer2.Dispose()
        
       
        Dim nwkform As New treeviewer2 With {.sptree = sp1, .Text = "Tree Viewer", ._viewsupport = True, ._supportnames = supnames} '
        nwkform.Show()
        restore()
        pdis = True
    End Function 'Calculates NJ-LILD and pvalue for each branch on the concatenated tree
    Function support_conc1(ByVal nrep As Integer, ByVal selsplits As Boolean, ByVal splits(,) As String) As Single()
        'Calculates NJ-ILD for each branch on a tree
        ProgressBar1.Maximum = nrep
        stopp = False
        If nrep = 0 Then ProgressBar1.Visible = False

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0
        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        Dim selocus(nmax - 1) As String
        Dim p As Integer
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                selocus(p) = GridView1(1, i - 1).Value
                p = p + 1
            End If
        Next

        i = 1

        '-------------------------------------------------------
        'makes individual alignments
        Dim sp(nmax - 1) As splits
        Dim j As Integer = 0
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                Dim concatenate(,) As String
                concatenate = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                concatenate = Module1.reductor(concatenate)
                sp(j).nsites = concatenate(1, 1).Length
                ProgressBar2.Increment(1)
                j = j + 1
                Application.DoEvents()
            End If
        Next
        Dim sp1 As splits
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)
                fi = fi + 1
            End If
        Next
        '-----------------------------------------------

        'Makes concatenated tree
        Dim concat(,) As String = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
        Dim treelengths(nmax - 1) As Single
        sp1 = Module1.NJ(concat, Module1.splitestmaker, 0, False, False, False)
        Dim newa As New newicker
        For i = 1 To sp1.nOTUs.GetLength(0) - 1
            sp1.nOTUs(i, 0) = newa.rewritesplits(sp1.nOTUs(i, 1), sp1.otumat1.GetLength(0) - 1)
        Next
        a = Nothing

        '-------------------------------------------------------------
        'Makes individual trees and calculates lengths
        ReDim a(1)
        ReDim Module1.otum(nmax - 1)

        For i = 0 To nmax - 1
            otum(i).treelength = 0
            a(1) = GridView1.Item(1, i).Value
            otum(i).otumat1 = conca(a, 1, TextBox12.Text, TextBox13.Text, True, False)
            otum(i).dista = distanciasx(otum(i).otumat1, otum(i).otumat1.GetLength(0) - 1)
            Dim ssp As splits
            ssp = Module1.NJ(otum(i).otumat1, Module1.splitestmaker, 0, False, False, False)
            For SS = 1 To ssp.notus1.GetLength(0) - 1
                otum(i).treelength = otum(i).treelength + ssp.notus1(SS, 0)
            Next
            treelengths(i) = otum(i).treelength
        Next
        '---------------------------------------------------------------

        'calculates extralengths for individual tree when a branch is forced to exist in a partition
        Dim extralengths() As Single
        extralengths = Module1.checkxtrasteps(sp1.nOTUs, 0, concat.GetLength(0) - 2, False, Nothing, treelengths, splits, selsplits)
        sp1.nsites = concat(1, 1).Length

        '--------------------------------------------------------------

        'Permutation test
        Dim concat1(,) As String = concat.Clone
        Dim matcount(sp1.nOTUs.GetLength(0) - 1) As Integer
        Dim count As Integer
        Dim countsp As Integer
        For h = 1 To nrep
            Dim sp2 As splits
            concat = Nothing
            concat = concat1.Clone
            ReDim Module1.otum(nmax - 1)
            For i = 0 To nmax - 1
                Dim conc1(concat.GetLength(0) - 1, concat.GetLength(1) - 1) As String

                For j = 0 To sp(i).nsites - 1
                    Dim pos As Integer = Math.Round(Rnd() * (concat(1, 1).Length - 1))

                    For k = 1 To conc1.GetLength(0) - 1
                        conc1(k, 1) = conc1(k, 1) & concat(k, 1).Substring(pos, 1)

                        concat(k, 1) = concat(k, 1).Remove(pos, 1)
                    Next
                Next

                a(1) = selocus(i)
                otum(i).otumat1 = conc1
                If otum(i).otumat1(1, 1) = Nothing Then
                    GoTo 23
                End If
                otum(i).dista = distanciasx(otum(i).otumat1, otum(i).otumat1.GetLength(0) - 1)
                Dim ssp As splits
                ssp = Module1.NJ(otum(i).otumat1, Module1.splitestmaker, 0, False, False, False)
                For SS = 1 To ssp.notus1.GetLength(0) - 1
                    otum(i).treelength = otum(i).treelength + ssp.notus1(SS, 0)
                Next

                If stopp = True Then
                    Exit Function
                End If

23:         Next

            Dim extralengths1() As Single
            extralengths1 = Module1.checkxtrasteps(sp1.nOTUs, 0, concat.GetLength(0) - 2, False, extralengths, treelengths, splits, selsplits)
            For ss = 1 To extralengths.Length - 1

                Dim xtc As Single = extralengths1(ss)
                If extralengths(ss) > xtc Then
                    matcount(ss) = matcount(ss) + 1
                End If

            Next
            Application.DoEvents()
            ProgressBar1.Increment(1)
        Next

        If nrep = 0 Then
            For f = 1 To sp1.notus1.GetLength(0) - 1
                sp1.notus1(f, 1) = Math.Round(((extralengths(f))) * 1000) / 1000
            Next
        Else
            Dim R(7) As Single
            For m = 0 To 5
                R(m) = 999
            Next
            Dim COUNTTP As Single = 0.05
            Dim COUNTFP As Integer = 0.05
            Dim bcorr As Boolean = False
            For f = sp1.otumat1.GetLength(0) To sp1.notus1.GetLength(0) - 2
                sp1.notus1(f, 1) = (1 - (matcount(f)) / (nrep))

                If sp1.notus1(f, 1) < 0.016 Then
                    bcorr = True
                End If
                Select Case sp1.nOTUs(f, 2)
                    Case "1 2 ,"

                        R(0) = sp1.notus1(f, 1)

                    Case "1 2 3 4 ,"

                        R(4) = sp1.notus1(f, 1)

                    Case "1 2 3 4 7 8 ,"

                        R(3) = sp1.notus1(f, 1)

                    Case "1 2 5 6 7 8 ,"

                        R(2) = sp1.notus1(f, 1)

                    Case "1 2 3 4 5 6 ,"

                        R(1) = sp1.notus1(f, 1)

                    Case Else
                        R(6) = R(6) + 1

                        R(5) = sp1.notus1(f, 1)

                End Select


            Next
            If bcorr = True Then
                R(7) = R(7) + 1
            End If
            '
            'support_conc = R
        End If
        Dim supnames(0) As String
        supnames(0) = "p-Val"
        'treeviewer2.Dispose()
        Dim nwkform As New treeviewer2 With {.sptree = sp1, .Text = "Tree Viewer", ._viewsupport = True, ._supportnames = supnames} '
        nwkform.Show()
        restore()
        pdis = True
    End Function 'Calculates NJ-LILD and pvalue for each branch on the concatenated tree
    Function checkcomp(ByVal split1 As String, ByVal split2 As String, ByVal n As Integer) As Boolean

     
        Dim contain As Boolean = False
        Dim xx, yy As Char
        For x = 0 To 1
            If x = 0 Then xx = "*"c Else xx = "."c

            For y = 0 To 1
                If y = 0 Then yy = "*"c Else yy = "."c
                Dim dentro As Boolean = True
                For i = 1 To n

                    If split1.Chars(i - 1) = xx Then
                        If split2.Chars(i - 1) <> yy Then
                            dentro = False
                            Exit For
                        End If
                    End If
                Next
                If dentro = True Then
                    contain = True
                    Return True
                    Exit Function
                End If
            Next
        Next
        Return contain
    End Function 'tests the compatibility between two splits
    Function leeroutputtree(ByVal file As String)
        'allows to read the output saved from a loci selection analysis
        Dim TREEMAT As treemat
        Dim SPLITS As splits
        Dim NODESxx() As nodeap
        Dim vs As Boolean
        Dim supnames As String()
        Dim ds As New DataSet
        ds.ReadXml(file)
        Dim dt1 As DataTable = ds.Tables(0)
        If dt1.Rows(0).Item(1) <> "1 ," Then
            MsgBox("The file is not recognized by MLSTest", MsgBoxStyle.OkOnly)
            Exit Function
        End If
        Dim count As Integer
        With ds.Tables(0)
            Dim n As Integer


            ReDim SPLITS.nOTUs(.Rows.Count, 3)
            ReDim SPLITS.notus1(.Rows.Count, 1)


            For i = 1 To .Rows.Count

                SPLITS.nOTUs(i, 1) = .Rows(i - 1).Item(0)
                SPLITS.nOTUs(i, 2) = .Rows(i - 1).Item(1)
                SPLITS.notus1(i, 0) = .Rows(i - 1).Item(2)
                If .Columns.Count > 3 Then
                    If IsDBNull(.Rows(i - 1).Item(3)) = False Then
                        SPLITS.notus1(i, 1) = .Rows(i - 1).Item(3)
                        vs = True
                        Dim a As Integer = .Rows(i - 1).Item(3).ToString.Split("/"c).Length
                        If a > count Then
                            count = a
                            ReDim supnames(a - 1)
                        End If
                    End If
                End If

            Next

        End With
        With ds.Tables(1)
            Dim n As Integer



            ReDim TREEMAT.nodo(.Rows.Count)
            ReDim TREEMAT.x(.Rows.Count, 1)
            ReDim TREEMAT.y(.Rows.Count, 2)
            ReDim NODESxx(.Rows.Count)
            Dim DT As New DataTable
            DT = ds.Tables(1)
            For i = 1 To .Rows.Count
                If .Rows(i - 1).Item(0) IsNot DBNull.Value Then
                    TREEMAT.nodo(i) = .Rows(i - 1).Item(0)
                    TREEMAT.x(i, 0) = .Rows(i - 1).Item(1)
                    TREEMAT.x(i, 1) = .Rows(i - 1).Item(2)
                End If
                If .Rows(i - 1).Item(3) IsNot DBNull.Value Then
                    TREEMAT.y(i, 0) = .Rows(i - 1).Item(3)
                    TREEMAT.y(i, 1) = .Rows(i - 1).Item(4)
                    TREEMAT.y(i, 2) = .Rows(i - 1).Item(5)
                End If
                If .Rows(i - 1).Item(6) IsNot DBNull.Value Then
                    Dim COLORX As String = .Rows(i - 1).Item(6)
                    Dim START As Integer = COLORX.IndexOf("[") + 1

                    Dim COLORT As String = COLORX.Substring(START, COLORX.Length - START - 1)
                    NODESxx(i).pen = New Pen(Color.Black, 1)
                    NODESxx(i).pen.Color = Color.FromArgb(COLORX)
                    NODESxx(i).pen.Width = (.Rows(i - 1).Item(7))
                End If

            Next

        End With
        With ds.Tables(2)
            ReDim SPLITS.otumat1(.Rows.Count - 1, 1)
            'Dim DT As New DataTable
            'DT = ds.Tables(2)
            For i = 1 To .Rows.Count - 1

                SPLITS.otumat1(i, 0) = .Rows(i).Item(0)
                SPLITS.otumat1(i, 1) = .Rows(i).Item(1)


            Next

        End With
        If supnames IsNot Nothing Then
            For i = 0 To supnames.Length - 1
                supnames(i) = "branch value " & i
            Next
        End If
        Dim tv As New treeviewer2
        tv.sptree = SPLITS
        tv.Text = "Tree Viewer"
        tv._treemat = TREEMAT
        tv._nodes = NODESxx
        tv._viewsupport = vs
        tv._supportnames = supnames
        tv.Show()



        ds.Dispose()







    End Function 'allows to read the output saved from a loci selection analysis
    Function rd() As Integer
        Dim A As Single = Rnd()
        If A > 0.499 Then
            rd = 249
        Else
            rd = 0
        End If
    End Function 'For internal test purposes
    

    

  

   

   

   
  
    Private Sub TestToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestToolStripMenuItem1.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "NJ-Analisis")
            If aa = MsgBoxResult.Yes Then
                For iii = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, iii).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        CancelButton1.Visible = True
        blockmenusytab()
        Dim bbb As MsgBoxResult
        bbb = MsgBox("Do you want to view a detailed analysis?. This includes wich loci are incongruent and the number of topological incongruences with a certain branch?", MsgBoxStyle.YesNo, "NJ-Analisis")
        If bbb = MsgBoxResult.Yes Then
            NJtree(False, False, False, 0, False, False, False, True)
        Else
            NJtreeborrar(False, False, True, 0, False, False, False, False)
        End If

        Exit Sub


        Dim matr(1000, 1000) As Byte
        Dim otumat1(1000, 1) As String
        For iv = 1 To 1000
            otumat1(iv, 0) = "ST" & iv
        Next
        Dim arr(400) As Integer
        Dim count As Integer
        For jjj = 0 To 999
            If DataGridView7.Item(0, jjj).Value.ToString.Contains("[1]") Then
                arr(count) = jjj + 1
                count = count + 1
            End If
        Next
        For ia = 1 To 1000

            For ib = 1 To 1000
                If arr.Contains(ia) And arr.Contains(ib) = False Then
                    matr(ia, ib) = 1
                ElseIf arr.Contains(ia) = False And arr.Contains(ib) = True Then
                    matr(ia, ib) = 1
                End If
            Next

        Next
        Dim txt2 As String = Nothing
        For hh = 1 To 1000
            For jj = 1 To 1000
                txt2 = txt2 & matr(hh, jj) & vbTab
            Next
            txt2 = txt2 & vbNewLine
        Next
        Clipboard.SetText(txt2)
        writedist(otumat1, matr)


        restore()

        Exit Sub


        GoTo l2
        Dim folder As String
        Dim txt1(8) As String
        For zz = 8 To 8
            Select Case zz
                Case 0
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\CONG\"
                Case 1
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\Aspergillus\reps\"
                Case 2
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\Blastocystis\reps\"
                Case 3
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\Leishmania\reps\"
                Case 4
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\Fusarium solani\reps\"
                Case 5
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\Candida albicans\reps\"
                Case 6
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\Candida glabrata\REPS\"
                Case 7
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\random\"
                Case 8
                    folder = "D:\Documentos\nuevo trabajo\MLSTest\cruzi\"
            End Select

            Dim res(9, 1) As Integer
            For z = 1 To 1
                clearseqs()
                leerproject(folder & "set24.mls")
                For xx = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, xx).Value = True
                Next
                pdis = False
                Dim splitsa(,) As String = Module1.splitestmaker()
                hethand = 0
                Dim R As Single()
                R = support_conc(1000, False, splitsa)
                pdis = True
                res(z - 1, 0) = R(0)
                res(z - 1, 1) = R(1)
            Next

            For hh = 1 To 9
                txt1(zz) = txt1(zz) & res(hh, 0) & vbTab & res(hh, 1) & vbNewLine
            Next
            Clipboard.SetText(txt1(zz))
        Next
        MsgBox("LISTO", MsgBoxStyle.OkOnly)
        Exit Sub
l2:
        Dim nlocres(0) As Single
        Dim bootres(0) As Single
        Dim incomp(0) As Single
        For xx = 0 To GridView1.RowCount - 1
            GridView1.Item(2, xx).Value = True
        Next
        For rep = 1 To 1

            For v = 0 To DataGridView7.RowCount - 1
                DataGridView7.Item(1, v).Value = False
                'DataGridView7.Item(0, v).Value = "ST" & v + 1
            Next
l22:
            Dim sem As Integer = System.DateTime.Now.Millisecond
            Dim rand As Random
            rand = New Random(sem)
            For r = 1 To 60
                Dim t As Boolean = False

                Dim b As Integer
                If DataGridView7.Item(1, b).Value = True Then
                    t = True
                    Do While t = True
                        b = rand.Next(0, DataGridView7.RowCount - 1)
                        If DataGridView7.Item(1, b).Value = False Then
                            'With DataGridView7.Item(0, b).Value.ToString
                            'If .EndsWith("[1]") Or .EndsWith("[2]") Or .EndsWith("[3]") Or .EndsWith("[4]") Then
                            DataGridView7.Item(1, b).Value = True

                            t = False
                            'End If
                            ' End With
                        End If
                    Loop
                End If

                With DataGridView7.Item(0, b).Value.ToString
                    'If .EndsWith("[1]") Or .EndsWith("[2]") Or .EndsWith("[3]") Or .EndsWith("[4]") Then
                    DataGridView7.Item(1, b).Value = True


                    'End If
                End With
                'DataGridView7.Item(0, b).Value = "t" & b + 1
            Next

            GoTo lfin
            actseqs()
            Dim sp As splits = NJtree2(True, True, 1000, False, False, False, False)

            nlocres = calculatex(nlocres, 1, sp)
            bootres = calculatex(bootres, 0, sp)
            incomp = calculatex(incomp, 2, sp)
            SAVEBORRAR(rep)
            Application.DoEvents()
        Next
        Dim txt As String = Nothing
        For h = 0 To nlocres.Length - 1
            txt = txt & nlocres(h) & vbTab & bootres(h) & vbTab & incomp(h) & vbNewLine
        Next

        Clipboard.SetText(txt)
        MsgBox("LISTO", MsgBoxStyle.OkOnly)
        Exit Sub
        Exit Sub
lin22:
        DataGridView5.Visible = True
        DataGridView5.ColumnCount = 9
        DataGridView5.RowCount = 1000
        TabControl1.SelectTab(TabPage11)
        DataGridView5.Visible = True
        Dim a(99) As Single
        Dim aL(99, 8) As Single
        Dim ch(10) As String
        Dim splits(,) As String = Module1.splitestmaker()
        Dim i, j As Integer
        Do While i = 0

            Rnd() '(-j)
            Dim AA As Single
            AA = rd()


            For XX = 1 To 3
                'AA = rd()
                Dim bb As Integer = Math.Round(Rnd() * 249)
                'If AA = 0 Then

                ch(XX) = "D:\Documentos\seqgen\LILD\incong500titv3i67l0.01\sim" & bb + 1 & ".dat.fasta"
                GridView1.Item(1, XX - 1).Value = ch(XX)
                'Else
                ' ch(XX) = "D:\seqgen\incong500h0.6l0.001\sim" & bb + 1 + AA & ".dat.fasta"
                ' GridView1.Item(1, XX - 1).Value = ch(XX)
                'End If

            Next
            For XX = 4 To 7
                ch(XX) = "D:\Documentos\seqgen\LILD\incong500titv3i67l0.002\sim" & Math.Round(Rnd() * 249) + 1 & ".dat.fasta"
                GridView1.Item(1, XX - 1).Value = ch(XX)
            Next
            'If i <> 510 And i <> 970 Then
            'a(i)
            'Dim ALX() As Single = support_conc(100, False, splits)
            Dim ALX() As Single = wstc1()
            'aL(i, 0) = ALX(0)
            'aL(i, 1) = ALX(1)

            'If ALX(1) = 0 Then
            'aL(i, 0) = ALX(0)

            ' Dim SP As splits
            'Dim boot, boot1 As Single
            'boot1 = 0
            'boot = 0
            'Dim cc As Integer = Math.Round(Rnd() * 249)
            'ch(5) = "D:\seqgen\incong500titv3i67l0.001\sim" & cc + 1 & ".dat.fasta"
            'SP = Module1.NJ(conca(ch, ch.Length - 1, TextBox12.Text, TextBox13.Text, True, False), Module1.splitestmaker, 1000, False, False, False) '0, False, True, False)

            'For v = 9 To SP.nOTUs.GetLength(0) - 2
            'If SP.nOTUs(v, 2) = "1 2 5 6 7 8 ," Then
            'boot = SP.notus1(v, 1)
            'End If
            'Next

            'ch(5) = "D:\seqgen\incong500titv3i67l0.001\sim" & cc + 251 & ".dat.fasta"
            'SP = Module1.NJ(conca(ch, ch.Length - 1, TextBox12.Text, TextBox13.Text, True, False), Module1.splitestmaker, 1000, False, False, False) '1000, False, False, False)

            'For v = 9 To SP.nOTUs.GetLength(0) - 2
            'If SP.nOTUs(v, 2) = "1 2 5 6 7 8 ," Then
            'boot1 = SP.notus1(v, 1)
            'End If
            'Next
            'aL(i, 0) = boot
            'aL(i, 1) = boot1
            'i = i + 1
            'End If
            'End If
            ''If i = 99 Then Stop
            'If aL(i) < 1000 Then Stop
            For b = 0 To 7
                DataGridView5.Item(b, j).Value = ALX(b)

            Next
            j = j + 1
            If j = 999 Then i = 1

        Loop

        ' Dim sp2 As splits
        ' sp2 = consensus(0, True, False)

        'For j = 1 To sp2.nOTUs.GetLength(0) - 1
        'If "1 2 3 4 ," = sp2.nOTUs(j, 2) And sp2.notus1(j, 0) <> 0 Then
        'a(i) = sp2.notus1(j, 1)
        ' Exit For
        ' End If

        ' Next



        'aL(i) = boot
        'a(i) = ILD_bionj(100)
        Application.DoEvents()
        ' Next

        TabControl1.SelectTab(TabPage11)
        DataGridView5.Visible = True



        ' For i = 0 To 999

        'DataGridView5.Item(0, i).Value = aL(i, 0)
        'DataGridView5.Item(1, i).Value = aL(i, 1)

        ' Next
lfin:
        'Stop

    End Sub 'internal tests BORRAR

    
    
    Private Sub WorkWithSTsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CalcToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalcToolStripMenuItem.Click
        RichTextBox1.Visible = True
        checklocus(1, 0, True)
        ProgressBar2.Visible = True
        restore()
    End Sub

    Private Sub SelectTheseLociToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectTheseLociToolStripMenuItem.Click
        Dim selected As String
        selected = DataGridView4.SelectedCells(0).Value
        selected = ", " & selected
        For i = 0 To GridView1.RowCount - 1
            If selected.Contains(GridView1.Item(0, i).Value) Then
                GridView1.Item(2, i).Value = True
            Else
                GridView1.Item(2, i).Value = False
            End If
        Next
    End Sub

    Private Sub ConvertToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConvertToolStripMenuItem.Click
        For i = 251 To 500
            Dim path As String = "D:\Documentos\seqgen\LILD\Copia de cGLAG1\sim" & i & ".dat"
            Dim lector As TextReader
            lector = New StreamReader(path)
            Dim lin As String
            Dim index As Integer = 1
            Dim txtarr() As String
            txtarr = Nothing
            ReDim txtarr(20)
            Dim filetext As String = Nothing
            lin = lector.ReadLine
            lin = lector.ReadLine
            txtarr(0) = lin
            lin = lector.ReadLine
            Do While lin <> Nothing
                If index = 5 Then '>= 6 And index <= 13 Then
l1:
                    'Dim a As Integer = Rnd() * 7 + 6

                    'If txtarr(a) IsNot Nothing Then GoTo l1
                    'txtarr(a) = lin
                    txtarr(16) = lin
                    txtarr(5) = lin
                ElseIf index = 16 Then
                    txtarr(16) = "txx" & txtarr(16).Substring(3)
                Else
                    txtarr(index) = lin
                End If
                lin = lector.ReadLine
                index = index + 1
            Loop
            lector.Close()
            filetext = Nothing
            For t = 0 To 20
                filetext = filetext & vbNewLine & txtarr(t)
            Next
            Dim save As New StreamWriter(path)
            save.WriteLine(filetext)
            save.Close()
        Next
    End Sub

    Private Sub TempletonModToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        wstc1()
    End Sub

   
    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub DiscriminatoryPowerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiscriminatoryPowerToolStripMenuItem.Click
        If DataGridView1.RowCount > 1 Then
            DPo()
        Else
            Dim result As MsgBoxResult
            result = MsgBox("There are no saved allelic profiles. MLSTest needs to calculate them first. Do you want to continue?", MsgBoxStyle.OkCancel, "Discriminatory Power Confidence Intervals")
            If result = MsgBoxResult.Ok Then
                profiles()
                DPo()
            End If

        End If


    End Sub

    Private Sub TestForNetworkStructuresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestForNetworkStructuresToolStripMenuItem.Click
       
        
        ''''
        Dim count As Integer
        For i = 0 To GridView1.RowCount - 1
            If GridView1.Item(2, i).Value = True Then
                count = count + 1
            End If

        Next
        If count = 0 Then count = GridView1.RowCount
        

        Dim dial As New Dialog2 With {.Text = "Find network structures", .page = 7}
        dial.ShowDialog()

        If dial.DialogResult = Windows.Forms.DialogResult.OK Then

            If DataGridView1.RowCount = 0 Then
                profiles()
            End If


            eBurstproc(False, 1, True, dial.gdef)

        End If

        


    End Sub 'netBurst button, test for network structures

    Private Sub TestForHiddenSLVInDiploidsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestForHiddenSLVInDiploidsToolStripMenuItem.Click
        Dim count As Integer
        For i = 0 To GridView1.RowCount - 1
            If GridView1.Item(2, i).Value = True Then
                count = count + 1
            End If

        Next
        If count = 0 Then count = GridView1.RowCount
       




        If DataGridView1.RowCount = 0 Then
            profiles()
        End If

        HiddenSLV(1)

    End Sub

    Private Sub SelectJustOneIsolatePerSTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectJustOneIsolatePerSTToolStripMenuItem.Click
        If DataGridView1.RowCount > 1 Then
            If DataGridView1.RowCount - 4 < DataGridView7.RowCount Then
                Dim result As MsgBoxResult
                result = MsgBox("MLSTest needs to recalculate Allelic Profiles. Do you want to continue?", MsgBoxStyle.OkCancel, "Select one strain per ST")
                If result = MsgBoxResult.Ok Then
                    For j = 0 To DataGridView7.RowCount - 1
                        DataGridView7.Item(1, j).Value = True
                    Next
                    actseqs()
                    profiles()
                    selOnePerST()
                End If
            Else
                selOnePerST()
            End If
        Else
            Dim result As MsgBoxResult
            result = MsgBox("There are no saved allelic profiles. MLSTest needs to calculate them first. Do you want to continue?", MsgBoxStyle.OkCancel, "Select one strain per ST")
            
            If result = MsgBoxResult.Ok Then
                For j = 0 To DataGridView7.RowCount - 1
                    DataGridView7.Item(1, j).Value = True
                Next
                actseqs()
                profiles()
                selOnePerST()
            End If

        End If
    End Sub
    Sub selOnePerST()
        Dim ncol As Integer = DataGridView1.ColumnCount
        Dim nrow As Integer = DataGridView1.RowCount
        Dim nSTs As Integer = DataGridView1.Item(ncol - 1, nrow - 4).Value
        For j = 0 To DataGridView7.RowCount - 1
            DataGridView7.Item(1, j).Value = False
        Next
        For i = 1 To nSTs
            For k = 1 To nrow - 4
                If DataGridView1.Item(ncol - 1, k - 1).Value = i Then

                    DataGridView7.Item(1, k - 1).Value = True
                    Exit For
                End If
            Next
        Next
        actseqs()
        TabControl1.SelectTab(TabPage1)
    End Sub

    Private Sub GridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellContentClick

    End Sub

    Private Sub GridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles GridView1.RowHeaderMouseClick


    End Sub

    Private Sub GridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles GridView1.ColumnHeaderMouseClick
        If e.ColumnIndex = 2 Then
            If GridView1.RowCount > 0 Then
                If ToolStripMenuItem10.Text = "Select all Loci" Then
                    For i = 0 To GridView1.RowCount - 1
                        GridView1.Item(2, i).Value = True
                    Next
                    ToolStripMenuItem10.Text = "Unselect all Loci"
                Else
                    For i = 0 To GridView1.RowCount - 1
                        GridView1.Item(2, i).Value = False
                    Next
                    ToolStripMenuItem10.Text = "Select all Loci"
                End If
            End If
        End If


    End Sub

    Private Sub DataGridView7_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick

    End Sub

    Private Sub DataGridView4_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub CongruenceAmongDistanceMatricesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CongruenceAmongDistanceMatricesToolStripMenuItem.Click

    End Sub

    Private Sub CADMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CADMToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Basic CADM")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Congruence Among Distance Matrices", .page = 6}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Permuting"
            CancelButton1.Visible = True
            blockmenusytab()

            Dim posteriori As Boolean
            Dim bb As MsgBoxResult
            bb = MsgBox("Do you want to calculate a posteriori tests", MsgBoxStyle.YesNo, "Basic CADM")
            If bb = MsgBoxResult.Yes Then
                posteriori = True
            End If
            CADM(a.nrep, False, posteriori, False)

        End If
        restore()



    End Sub

    Function CADM(ByVal nper As Integer, ByVal topo As Boolean, ByVal post As Boolean, ByVal test As Boolean)

        'It calculates pvalue for CADM
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False
        ProgressBar1.Visible = True
        ToolStripStatusLabel3.Text = "Permuting"
        CancelButton1.Visible = True
        blockmenusytab()


        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next



        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        Dim names(nmax - 1) As String
        ProgressBar2.Maximum = ch.Length - 1


        i = 1

        Dim sp(nmax - 1) As dsx
        If test = True Then
            ReDim sp(nmax)
        End If

        Dim j As Integer
        '---------------------------------------
        'Makes individual matrices
        ProgressBar1.Maximum = nper + nmax
        If post = True Then
            ProgressBar1.Maximum = ProgressBar1.Maximum + nper * nmax
        End If
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                names(j) = GridView1.Item(0, i - 1).Value
                Dim concatenate(,) As String
                concatenate = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                sp(j) = NJgo(concatenate, True, False)

                j = j + 1
                Application.DoEvents()
                ProgressBar1.Increment(1)
            End If

        Next
        If topo = True Then
            Dim sp1(sp.Length - 1) As splits
            For i = 0 To sp1.Length - 1
                sp1(i) = NJproc(sp(0).distancias.GetLength(0) - 1, sp(i).distancias, Nothing, Nothing, 0, sp(i).Vij, True)
                Dim newa As New newicker
                For t = sp(0).distancias.GetLength(0) To sp1(i).nOTUs.GetLength(0) - 1
                    sp1(i).nOTUs(t, 0) = newa.rewritesplits(sp1(i).nOTUs(t, 1), sp(0).distancias.GetLength(0) - 1)
                Next
                For x = 1 To sp(0).distancias.GetLength(0) - 2
                    For y = x + 1 To sp(0).distancias.GetLength(0) - 1
                        If sp(i).distancias(x, y) > 10 ^ -9 Then
                            sp(i).distancias(x, y) = calcB(sp1(i).nOTUs, x, y)
                        End If
                    Next
                Next
            Next

        End If
        Dim elm As Integer = sp(0).distancias.GetLength(0) - 1


        Dim RankLength = (elm ^ 2 - elm) / 2
        Dim ranks(sp.Length - 1) As vectomat
        For i = 0 To sp.Length - 1
            ReDim ranks(i).matv(RankLength - 1)
        Next
        Dim Tie As Long = 0
        calculateranks(sp, ranks, RankLength, Tie)
        Dim W, S As Double
        Dim p As Double = ranks.Length
        Dim n As Double = ranks(0).matv.Length
        S = calcW(ranks, Tie)
        W = (12 * S - (3 * p ^ 2 * n * (n + 1) ^ 2)) / (p ^ 2 * (n ^ 3 - n) - p * Tie)

        Dim COUNT As Integer = 1

        Dim matr(0) As dsx
        Dim rnd As Random

        sem = 1 'System.DateTime.Now.Millisecond
        rnd = New Random(sem)

        For i = 2 To nper
            Dim Wp, Spe As Double
            Dim ranksp(ranks.Length - 1) As vectomat

            For m = 0 To ranks.Length - 1

                ranksp(m).matv = permuteranks1(ranks(m).matv, elm)
                sem = rnd.Next
            Next
            Spe = calcW(ranksp, Tie)

            If Spe >= S Then 'Wp + (Wp * 0.002) >= W Then
                COUNT = COUNT + 1

            End If
            ProgressBar1.Increment(1)
            Application.DoEvents()

        Next
        Dim hyp(sp.Length - 1) As Single

        If post = True Then
            Dim count1 As Integer = 1
            For mat = 0 To sp.Length - 1
                count1 = 1
                For i = 2 To nper

                    Dim Wp As Single
                    Dim ranksp(ranks.Length - 1) As vectomat
                    For j = 0 To ranksp.Length - 1
                        If j <> mat Then
                            ranksp(j).matv = ranks(j).matv.Clone
                        End If
                    Next
                    ranksp(mat).matv = permuteranks1(ranks(mat).matv, elm)
                    sem = rnd.Next
                    Dim Spe As Double
                    Spe = calcW(ranksp, Tie)
                    If Spe >= S Then 'Wp + (Wp * 0.002) >= W Then
                        count1 = count1 + 1

                    End If
                    ProgressBar1.Increment(1)
                    Application.DoEvents()

                Next
                hyp(mat) = count1 / nper
            Next
            Dim str As String
            str = "H0: all fragments incongruent:" & vbNewLine & "Kendall's W=" & W & vbNewLine & "p value=" & COUNT / nper
            For x = 0 To hyp.Length - 1
                str = str & vbNewLine & "H0:" & names(x) & " incongruent with all, p value=" & hyp(x)
            Next
            MsgBox(str)

            Exit Function
        End If

        MsgBox("H0: all fragments incongruent:" & vbNewLine & "Kendall's W=" & W & vbNewLine & "p value=" & COUNT / nper)

    End Function
  

    Function CADMtest(ByVal nper As Integer, ByVal topo As Boolean, ByVal post As Boolean)

        'It calculates pvalue for CADM
        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        pdis = False
        ProgressBar1.Visible = True
        ToolStripStatusLabel3.Text = "Permuting"
        CancelButton1.Visible = True
        blockmenusytab()


        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next



        Dim a(1) As String
        Dim fi As Integer = 1
        i = 1
        Dim t As Integer = 0


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        Dim names(nmax - 1) As String
        ProgressBar2.Maximum = ch.Length - 1


        i = 1

        Dim sp(nmax) As dsx

        

        Dim j As Integer
        '---------------------------------------
        'Makes individual matrices
        ProgressBar1.Maximum = nper + nmax
        If post = True Then
            ProgressBar1.Maximum = ProgressBar1.Maximum + nper * nmax
        End If
        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                a.SetValue(locus(i), 1)
                names(j) = GridView1.Item(0, i - 1).Value
                Dim concatenate(,) As String
                concatenate = conca(a, 1, TextBox12.Text, TextBox13.Text, avstates, False)
                sp(j) = NJgo(concatenate, True, False)

                j = j + 1
                Application.DoEvents()
                ProgressBar1.Increment(1)
            End If

        Next
        Dim Ns As Integer = sp(0).distancias.GetLength(0) - 1
        ReDim sp(nmax).distancias(Ns, Ns)



      
        Dim arr(Ns) As Integer
        Dim countg As Integer
        For jjj = 0 To Ns - 1
            If DataGridView7.Item(0, jjj).Value.ToString.Contains("[1]") Then
                arr(countg) = jjj + 1
                countg = countg + 1
            End If
        Next

        If topo = True Then
            Dim sp1(sp.Length - 2) As splits
            For i = 0 To sp1.Length - 2
                sp1(i) = NJproc(sp(0).distancias.GetLength(0) - 1, sp(i).distancias, Nothing, Nothing, 0, sp(i).Vij, True)
                Dim newa As New newicker
                For t = sp(0).distancias.GetLength(0) To sp1(i).nOTUs.GetLength(0) - 1
                    sp1(i).nOTUs(t, 0) = newa.rewritesplits(sp1(i).nOTUs(t, 1), sp(0).distancias.GetLength(0) - 1)
                Next
                For x = 1 To sp(0).distancias.GetLength(0) - 2
                    For y = x + 1 To sp(0).distancias.GetLength(0) - 1
                        If sp(i).distancias(x, y) > 10 ^ -9 Then
                            sp(i).distancias(x, y) = calcB(sp1(i).nOTUs, x, y)
                        End If
                    Next
                Next
            Next

        End If

        For ia = 1 To Ns

            For ib = 1 To Ns
                If arr.Contains(ia) And arr.Contains(ib) = False Then
                    sp(nmax).distancias(ia, ib) = 1
                ElseIf arr.Contains(ia) = False And arr.Contains(ib) = True Then
                    sp(nmax).distancias(ia, ib) = 1
                End If
            Next

        Next


      
        Dim elm As Integer = sp(0).distancias.GetLength(0) - 1


        Dim RankLength = (elm ^ 2 - elm) / 2
        Dim ranks(sp.Length - 1) As vectomat
        For i = 0 To sp.Length - 1
            ReDim ranks(i).matv(RankLength - 1)
        Next
        Dim Tie As Long = 0
        calculateranks(sp, ranks, RankLength, Tie)
        Dim W, S As Double
        Dim p As Double = ranks.Length
        Dim n As Double = ranks(0).matv.Length
        S = calcW(ranks, Tie)
        W = (12 * S - (3 * p ^ 2 * n * (n + 1) ^ 2)) / (p ^ 2 * (n ^ 3 - n) - p * Tie)

        Dim COUNT As Integer = 1

        Dim matr(0) As dsx
        Dim rnd As Random

        sem = 1 'System.DateTime.Now.Millisecond
        rnd = New Random(sem)

        For i = 2 To nper
            Dim Wp, Spe As Double
            Dim ranksp(ranks.Length - 1) As vectomat

            For m = 0 To ranks.Length - 1

                ranksp(m).matv = permuteranks1(ranks(m).matv, elm)
                sem = rnd.Next
            Next
            Spe = calcW(ranksp, Tie)

            If Spe >= S Then 'Wp + (Wp * 0.002) >= W Then
                COUNT = COUNT + 1

            End If
            ProgressBar1.Increment(1)
            Application.DoEvents()

        Next
        Dim hyp(sp.Length - 1) As Single

        If post = True Then
            Dim count1 As Integer = 1
            For mat = 0 To sp.Length - 1
                count1 = 1
                For i = 2 To nper

                    Dim Wp As Single
                    Dim ranksp(ranks.Length - 1) As vectomat
                    For j = 0 To ranksp.Length - 1
                        If j <> mat Then
                            ranksp(j).matv = ranks(j).matv.Clone
                        End If
                    Next
                    ranksp(mat).matv = permuteranks1(ranks(mat).matv, elm)
                    sem = rnd.Next
                    Dim Spe As Double
                    Spe = calcW(ranksp, Tie)
                    If Spe >= S Then 'Wp + (Wp * 0.002) >= W Then
                        count1 = count1 + 1

                    End If
                    ProgressBar1.Increment(1)
                    Application.DoEvents()

                Next
                hyp(mat) = count1 / nper
            Next
            Dim str As String
            str = "H0: all fragments incongruent:" & vbNewLine & "Kendall's W=" & W & vbNewLine & "p value=" & COUNT / nper
            For x = 0 To hyp.Length - 1
                str = str & vbNewLine & "H0:" & names(x) & " incongruent with all, p value=" & hyp(x)
            Next
            MsgBox(str)

            Exit Function
        End If

        MsgBox("H0: all fragments incongruent:" & vbNewLine & "Kendall's W=" & W & vbNewLine & "p value=" & COUNT / nper)

    End Function

    Function calculateranks(ByVal sp() As dsx, ByRef ranks() As vectomat, ByVal l As Integer, ByRef Tie As Long) As vectomat()
        Dim a As Integer = sp(0).distancias.GetLength(0) - 1
        For i = 0 To sp.Length - 1
            Dim el As Integer = 0
            Dim arr(l - 1) As Double

            For j = 1 To a
                For k = j + 1 To a
                    arr(el) = sp(i).distancias(j, k)
                    el = el + 1
                Next
            Next
            Dim arr1() As Double
            arr1 = arr.Clone
            Array.Sort(arr1)
            ranks(i).matv = definerank(arr1, arr, Tie)




        Next



    End Function
    Function definerank(ByVal x() As Double, ByVal y() As Double, ByRef Tie As Long) As Single()

        Dim ranks(x.Length - 1) As Single
        Dim assrank(x.Length - 1) As Boolean
        assrank.Clear(assrank, 0, assrank.Length)
        For i = 0 To x.Length - 1
            Dim T3 As Integer = 0
            Dim rank As Double = i + 1
            Dim count As Integer = 1
            Dim j As Integer
            For j = 1 To 1000000
                If i + j < x.Length Then
                    If x(i) = x(i + j) Then
                        count = count + 1
                        rank = rank + (i + 1 + j)
                    Else
                        Exit For

                    End If
                Else : Exit For
                End If
            Next
            If count > 1 Then
                Tie = Tie + count ^ 3 - count
            End If
            rank = rank / count
            For g = 0 To y.Length - 1
                If y(g) = x(i) Then
                    If assrank(g) = False Then
                        ranks(g) = rank
                        assrank(g) = True
                    End If
                End If
            Next
            i = i + j - 1
        Next



        Return ranks
    End Function
    Function calcW(ByVal rank() As vectomat, ByVal T As Double) As Double
        Dim W As Double
        Dim p = rank.Length
        Dim n = rank(0).matv.Length
        Dim S As Double
        For i = 0 To n - 1
            S = S + sumranks2(i, rank)
        Next



      
        Return S
    End Function
    Function sumranks2(ByVal index As Integer, ByVal rank() As vectomat) As Double
        Dim a As Double
        For i = 0 To rank.Length - 1
            a = a + rank(i).matv(index)
        Next

        Return a ^ 2
    End Function
    Function permuteranks(ByVal rankes As vectomat(), ByVal index As Integer) As vectomat()
        Dim p As Integer = rankes.Length
        Dim n As Integer = rankes(0).matv.Length
        Dim ranks(rankes.Length - 1) As vectomat

        If index = -1 Then

            For i = 0 To p - 1
                ranks(i).matv = rankes(i).matv.Clone
                Dim v As Integer = n
                Do While v > 1
                    v = v - 1
                    Dim interc As Single
                    Dim j As Integer = Math.Round(Rnd() * v)
                    interc = ranks(i).matv(j)
                    ranks(i).matv(j) = ranks(i).matv(v)
                    ranks(i).matv(v) = interc

                Loop

            Next
        Else
            ranks(index).matv = rankes(index).matv.Clone
            Dim v As Integer = n
            Do While v > 1
                v = v - 1
                Dim interc As Single
                Dim j As Integer = Math.Round(Rnd() * v)
                interc = ranks(index).matv(j)
                ranks(index).matv(j) = ranks(index).matv(v)
                ranks(index).matv(v) = interc
            Loop
        End If
        Return ranks
    End Function
    Private sem As Integer
    Function permuteranks1(ByVal ranks As Single(), ByVal ncol As Integer) As Single()


        Dim rankes() As Single = ranks.Clone
        Dim rand As Random

        rand = New Random(sem)
        Dim v As Integer = ncol
        Do While v > 1

            Dim interc As Single

            Dim j As Integer = rand.Next(1, v)
            For f = 1 To j - 1



                Dim index As Integer = Math.Round((f - 1) * (ncol - f / 2) + j - f - 1)

                Dim index2 As Integer = Math.Round((f - 1) * (ncol - f / 2) + v - f - 1)
                interc = rankes(index)
                rankes(index) = rankes(index2)
                rankes(index2) = interc
            Next

            For c = j + 1 To v - 1
                Dim index As Integer = Math.Round((j - 1) * (ncol - j / 2) + c - j - 1)

                Dim index2 As Integer = Math.Round((c - 1) * (ncol - c / 2) + v - c - 1)
                interc = rankes(index)
                rankes(index) = rankes(index2)
                rankes(index2) = interc


            Next
            For c = v + 1 To ncol
                Dim index As Integer = Math.Round((j - 1) * (ncol - j / 2) + c - j - 1)

                Dim index2 As Integer = Math.Round((j - 1) * (ncol - j / 2) + v - j - 1)
                interc = rankes(index)
                rankes(index) = rankes(index2)
                rankes(index2) = interc
            Next
            v = v - 1

        Loop

        Return rankes
    End Function
    Function permuteranksx(ByVal ranks As Integer(), ByVal N As Integer) As Integer()


        Dim rankes() As Integer = ranks.Clone

        'Dim sem As Integer = System.DateTime.Now.Millisecond ' rnd() * 100000

        Dim v As Integer = N - 1

        Do While v > 0

            Dim interc As Single
            Dim j As Integer = Math.Round(Rnd() * (v))
            interc = rankes(j)
            rankes(j) = rankes(v)
            rankes(v) = interc
            v = v - 1
        Loop

        Return rankes
    End Function


    






    Private Sub NonultrametricCADMTestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonultrametricCADMTestToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Basic CADM")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Congruence Among Distance Matrices", .page = 6}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Permuting"
            CancelButton1.Visible = True
            blockmenusytab()
            pdis = False
            Dim posteriori As Boolean
            Dim bb As MsgBoxResult
            bb = MsgBox("Do you want to calculate a posteriori tests", MsgBoxStyle.YesNo, "Basic CADM")
            If bb = MsgBoxResult.Yes Then
                posteriori = True
            End If
            CADM(a.nrep, True, posteriori, False)
            pdis = True
        End If
        restore()

    End Sub

    Private Sub TabPage8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage8.Click

    End Sub

    Private Sub TestSpecificClusterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseClick

    End Sub

    Private Sub CopyAsTextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyAsTextToolStripMenuItem.Click
        Dim arr() As Array
        Dim txt As String
        Dim a As Integer = DataGridView2.SelectedCells.Count



        txt = alignmf(DataGridView2.SelectedCells.Item(0).RowIndex + 1, 1)
        Clipboard.SetText(txt)
       

       


    End Sub

    Private Sub LinkToADatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkToADatabaseToolStripMenuItem.Click
        Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://pubmlst.org/perl/mlstdbnet/mlstdbnet.pl?page=alleles&format=FASTA&locus=ANX4&file=af_profiles.xml")

        Dim response As System.Net.HttpWebResponse = request.GetResponse()

        ' Check if the response is OK (status code 200)
        If response.StatusCode = System.Net.HttpStatusCode.OK Then

            ' Parse the contents from the response to a stream object
            Dim stream As System.IO.Stream = response.GetResponseStream()
            ' Create a reader for the stream object
            Dim reader As New System.IO.StreamReader(stream)
            ' Read from the stream object using the reader, put the contents in a string
            Dim contents As String = reader.ReadToEnd()
            ' Create a new, empty XML document
            Dim document As New System.Xml.XmlDocument()

            ' Load the contents into the XML document
            document.LoadXml(contents)

            ' Now you have a XmlDocument object that contains the XML from the remote site, you can
            ' use the objects and methods in the System.Xml namespace to read the document

        Else
            ' If the call to the remote site fails, you'll have to handle this. There can be many reasons, ie. the 
            ' remote site does not respond (code 404) or your username and password were incorrect (code 401)
            '
            ' See the codes in the System.Net.HttpStatusCode enumerator 

            Throw New Exception("Could not retrieve document from the URL, response code: " & response.StatusCode)

        End If

    End Sub

    Private Sub DataGridView6_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView6.RowsRemoved
        'Stop
    End Sub

    Private Sub DataGridView6_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DataGridView6.UserDeletingRow

    End Sub

    Private Sub DataGridView6_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView6.VisibleChanged

    End Sub

    '''
    Function calculatex(ByVal result() As Single, ByVal index As Integer, ByVal sp As splits) As Single()

        Dim results() As Single = result


        For z = sp.otumat1.GetLength(0) - 1 To sp.nOTUs.GetLength(0) - 2



            If sp.notus1(z, 1) <> Nothing And sp.notus1(z, 0) > 10 ^ -9 Then

                Dim vals() As String = Nothing


                vals = sp.notus1(z, 1).Split("/")
                Dim a As Single = vals(index)

                results(results.Length - 1) = a
                If z <> sp.nOTUs.GetLength(0) - 2 Then
                    Array.Resize(results, results.Length + 1)
                End If
            End If

        Next
        Return results
    End Function
    Private Function NJtree2(ByVal nlsupp As Boolean, ByVal incong As Boolean, ByVal nrep As Integer, ByVal bionj As Boolean, ByVal facs As Boolean, ByVal facsplt As Boolean, ByVal detincong As Boolean) As splits
        Dim vs As Boolean

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        textito = Nothing
        Dim arch As String = Nothing
        ProgressBar2.Value = 0


        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0
        'SaveFileDialog1.ShowDialog()
        'Dim pathtosave As String = SaveFileDialog1.FileName


        i = 1


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1 - nmax


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)


                fi = fi + 1
            End If
        Next



        i = 1
        Dim sp As splits




        sp = Module1.NJ(conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, bionj, facs, facsplt)





        Application.DoEvents()
        Dim supnames() As String

        If nrep > 0 Then
            vs = True
            ReDim supnames(0)
            supnames(supnames.Length - 1) = "Bootstrap"

        End If
       
        If nlsupp = True Then
            Dim bar As Boolean
            If nrep > 0 Or facs = True Then
                bar = True
            End If
            vs = True

            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "#loci"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "#loci"
            End If
            Dim sp2 As splits
            sp2 = consensus(0, True, bionj)
            For i = sp.otumat1.GetLength(0) To (sp.nOTUs.GetLength(0) - 1)
                For j = 1 To sp2.nOTUs.GetLength(0) - 1
                    If sp.nOTUs(i, 2) = sp2.nOTUs(j, 2) And sp2.notus1(j, 0) <> 0 Then

                        If sp.notus1(i, 0) > 10 ^ -9 Or sp.notus1(i, 0) < -10 ^ -9 Then
                            Dim vrc As String = Nothing
                            If bar = True Then

                                vrc = sp.notus1(i, 1) & "/"
                            End If

                            sp.notus1(i, 1) = vrc & sp2.notus1(j, 1)
                        End If
                        Exit For
                    End If
                    If j = sp2.nOTUs.GetLength(0) - 1 Then
                        If sp.notus1(i, 0) > 10 ^ -9 Or sp.notus1(i, 0) < -10 ^ -9 Then
                            Dim vrc As String = Nothing
                            If bar = True Then

                                vrc = sp.notus1(i, 1) & "/"

                            End If
                            sp.notus1(i, 1) = vrc & "0"
                        End If
                    End If

                Next
            Next
        End If

        If incong = True Then
            compatibilityx(sp, a, avstates)

            vs = True
            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "incong"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "incong"
            End If

        End If

        ' writefiles(sp, ToolStripComboBox1.SelectedIndex + 1, True, False)
        'Dim nwkform As New treeviewer1 With {.nwktree = textito, .Text = "Tree Viewer"}

        Return sp

        restore()

        'RichTextBox1.SaveFile(pathtosave, RichTextBoxStreamType.PlainText)
    End Function
    Sub compatibilityx(ByVal sp As splits, ByVal a() As String, ByVal avstates As Boolean)
        Dim spi(a.Length - 1) As splits
        Dim newa As New newicker
        Dim N As Integer = sp.otumat1.GetLength(0) - 1


        For i = N + 1 To sp.nOTUs.GetLength(0) - 1
            sp.nOTUs(i, 0) = newa.rewritesplits(sp.nOTUs(i, 1), N)

        Next
        For i = 1 To a.Length - 1
            Dim an(1) As String
            an(1) = a(i)
            spi(i) = NJ(conca(an, 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)
            For ii = N + 1 To sp.nOTUs.GetLength(0) - 1
                spi(i).nOTUs(ii, 0) = newa.rewritesplits(spi(i).nOTUs(ii, 1), N)

            Next
        Next
        For j = sp.otumat1.GetLength(0) To sp.nOTUs.GetLength(0) - 2
            If sp.notus1(j, 0) > 10 ^ -9 Then

                Dim incomp As Integer = 0
                For k = 1 To spi.Length - 1
                    For m = spi(1).otumat1.GetLength(0) To spi(k).nOTUs.GetLength(0) - 2
                        If spi(k).notus1(m, 0) > 10 ^ -9 Then
                            If checkcomp(sp.nOTUs(j, 0), spi(k).nOTUs(m, 0), sp.otumat1.GetLength(0) - 1) = False Then
                                incomp = incomp + 1
                                Exit For
                            End If
                        End If

                    Next
                Next

                sp.notus1(j, 1) = sp.notus1(j, 1) & "/" & incomp
            End If
        Next

    End Sub ' calculates the number of imcompatible splits for each branch of the concatenated tree
    Sub SAVEBORRAR(ByVal NAME As String)
        If GridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim texto As String
        Dim c As String = TextBox13.Text
        Dim t As Char = Char.ConvertFromUtf32(Keys.Tab)



        texto = "MLSTest Project" & c
        For i = 0 To GridView1.RowCount - 1

            texto = texto & GridView1.Item(0, i).Value & t
            texto = texto & GridView1.Item(1, i).Value & t
            texto = texto & GridView1.Item(2, i).Value & c

        Next
        texto = texto & ";" & c & c
        For i = 0 To DataGridView7.RowCount - 1
            If DataGridView7.Item(1, i).Value = True Then
                texto = texto & "+"
            Else : texto = texto & "-"
            End If

        Next
        texto = texto & c & ";" & c & c
        If DataGridView6.RowCount <> 0 Then
            For i = 0 To DataGridView6.RowCount - 1

                texto = texto & DataGridView6.Item(0, i).Value & t
                texto = texto & DataGridView6.Item(1, i).Value & t
                texto = texto & DataGridView6.Item(2, i).Value & t
                texto = texto & DataGridView6.Item(3, i).Value & c
            Next
        End If
        texto = texto & ";" & c & c & hethand & c & c & ";" & c & c


        texto = texto & ToolStripComboBox1.SelectedIndex & c & ";" & c & c
        For i = 1 To tests.Length - 1
            texto = texto & tests(i) & c
        Next
        texto = texto & ";" & c & c

        Dim xx As String

        Application.DoEvents()

      
        Dim save As New StreamWriter("D:\" & NAME & ".MLS")
        save.WriteLine(texto)
        save.Close()
        nameofproject = SaveFileDialog1.FileName

       
    End Sub
    Function permuteseqs(ByVal otumat1(,) As String) As String(,)
        Dim len As Integer = otumat1.GetLength(0) - 1
        Dim i As Integer = len

        Do While i > 2
            i = i - 1
            Dim j As Integer = Math.Round(Rnd() * (i - 2)) + 1
            Dim xseq As String = otumat1(j, 1)
            otumat1(j, 1) = otumat1(i, 1)
            otumat1(i, 1) = xseq
        Loop
        Return otumat1
    End Function
    '''

    Private Sub MCSTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MCSTToolStripMenuItem.Click
        GoTo l23
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Topological incongruence")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
l23:
        ProgressBar2.Visible = True

        ToolStripStatusLabel4.Visible = True
        CancelButton1.Visible = True
        blockmenusytab()

        Dim result(0) As Integer
        For i = 0 To GridView1.RowCount - 1
            GridView1.Item(2, i).Value = True
        Next
        result(0) = MCST(False)

        For i = 0 To GridView1.RowCount - 2

            For j = i + 1 To GridView1.RowCount - 1
                For k = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, k).Value = False
                Next

                GridView1.Item(2, i).Value = True
                GridView1.Item(2, j).Value = True
                Array.Resize(result, result.Length + 1)
                result(result.Length - 1) = MCST(False)

            Next
        Next


        restore()
    End Sub

    Private Sub TestSpecificGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestSpecificGroupToolStripMenuItem.Click
        If checksel() = 0 Then
            Dim aa As MsgBoxResult
            aa = MsgBox("None locus has been fixed. Do you want to use all loci in the analysis?", MsgBoxStyle.YesNo, "Basic CADM")
            If aa = MsgBoxResult.Yes Then
                For i = 0 To GridView1.RowCount - 1
                    GridView1.Item(2, i).Value = True
                Next
                ToolStripMenuItem10.Text = "Unselect all Loci"
            Else
                Exit Sub
            End If
        End If
        Dim a As New Dialog2 With {.Text = "Congruence Among Distance Matrices", .page = 6}
        a.ShowDialog()
        If a.DialogResult = Windows.Forms.DialogResult.OK Then
            ProgressBar1.Visible = True
            ToolStripStatusLabel3.Text = "Permuting"
            CancelButton1.Visible = True
            blockmenusytab()

           
            CADMtest(a.nrep, True, False)

        End If
        restore()

    End Sub
    Private Sub NJtreeborrar(ByVal perfi As Boolean, ByVal nlsupp As Boolean, ByVal incong As Boolean, ByVal nrep As Integer, ByVal bionj As Boolean, ByVal facs As Boolean, ByVal facsplt As Boolean, ByVal detincong As Boolean)
        Dim vs As Boolean

        Dim avstates As Boolean
        If hethand <> 1 Then
            avstates = True
        Else
            avstates = False
        End If
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        textito = Nothing
        Dim arch As String = Nothing
        ProgressBar2.Value = 0


        Dim a() As String
        Dim fi As Integer = 1

        Dim t As Integer = 0
        'SaveFileDialog1.ShowDialog()
        'Dim pathtosave As String = SaveFileDialog1.FileName


        i = 1


        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1 - nmax


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)


                fi = fi + 1
            End If
        Next



        i = 1
        Dim sp As splits
        If perfi = False Then



            sp = NJborrar(conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, nrep, bionj, facs, facsplt)


        Else
            Dim perfiles(,) = Perfal.perf(a, TextBox12.Text, locusname, True)
            otumat1 = conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False)
            sp = Module1.NJperfal(perfiles, otumat1, arch, nrep)

        End If


        Application.DoEvents()
        Dim supnames() As String

        If nrep > 0 Then
            vs = True
            ReDim supnames(0)
            supnames(supnames.Length - 1) = "Bootstrap"

        End If
        If facs = True Then
            vs = True
            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "1-njCS"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "1-njCS"
            End If
        End If
        If nlsupp = True Then
            Dim bar As Boolean
            If nrep > 0 Or facs = True Then
                bar = True
            End If
            vs = True

            If supnames Is Nothing Then
                ReDim supnames(0)
                supnames(supnames.Length - 1) = "#loci"
            Else
                Array.Resize(supnames, supnames.Length + 1)
                supnames(supnames.Length - 1) = "#loci"
            End If
            Dim sp2 As splits
            sp2 = consensus(0, True, bionj)
            For i = sp.otumat1.GetLength(0) To (sp.nOTUs.GetLength(0) - 1)
                For j = 1 To sp2.nOTUs.GetLength(0) - 1
                    If sp.nOTUs(i, 2) = sp2.nOTUs(j, 2) And sp2.notus1(j, 0) <> 0 Then

                        If sp.notus1(i, 0) > 10 ^ -9 Or sp.notus1(i, 0) < -10 ^ -9 Then
                            Dim vrc As String = Nothing
                            If bar = True Then

                                vrc = sp.notus1(i, 1) & "/"
                            End If

                            sp.notus1(i, 1) = vrc & sp2.notus1(j, 1)
                        End If
                        Exit For
                    End If
                    If j = sp2.nOTUs.GetLength(0) - 1 Then
                        If sp.notus1(i, 0) > 10 ^ -9 Or sp.notus1(i, 0) < -10 ^ -9 Then
                            Dim vrc As String = Nothing
                            If bar = True Then

                                vrc = sp.notus1(i, 1) & "/"

                            End If
                            sp.notus1(i, 1) = vrc & "0"
                        End If
                    End If

                Next
            Next
        End If

        If incong = True Then
            compatibilityborrar(sp, a, avstates)

            vs = True
            ReDim supnames(0)
            supnames(0) = "incongruent loci"

        End If
        If detincong = True Then
            relaxedIncong(sp, 0, True, False)
            Dim x As Integer
            ReDim supnames(0)
            vs = True
            For g = 0 To GridView1.RowCount - 1
                If GridView1.Item(2, g).Value = True Then
                    Array.Resize(supnames, x + 1)
                    supnames(x) = GridView1.Item(0, x).Value
                    x = x + 1
                End If
            Next

        End If
        ' writefiles(sp, ToolStripComboBox1.SelectedIndex + 1, True, False)
        'Dim nwkform As New treeviewer1 With {.nwktree = textito, .Text = "Tree Viewer"}


        Dim nwkform As New treeviewer2 With {.sptree = sp, .Text = "Tree Viewer", ._viewsupport = vs, ._supportnames = supnames} '
        nwkform.Show()
        restore()

        'RichTextBox1.SaveFile(pathtosave, RichTextBoxStreamType.PlainText)
    End Sub
    Sub compatibilityborrar(ByVal sp As splits, ByVal a() As String, ByVal avstates As Boolean)

        Dim newa As New newicker
        Dim N As Integer = sp.otumat1.GetLength(0) - 1
        Dim locus(GridView1.Rows.Count) As String
        Dim i, nmax As Integer
        For i = 1 To locus.Length - 1
            locus(i) = GridView1.Item(1, i - 1).Value
        Next
        textito = Nothing
        Dim arch As String = Nothing
        ProgressBar2.Value = 0



        Dim fi As Integer = 1

        Dim t As Integer = 0
        'SaveFileDialog1.ShowDialog()
        'Dim pathtosave As String = SaveFileDialog1.FileName


        i = 1
        For xx = 0 To GridView1.RowCount - 1
            GridView1.Item(2, xx).Value = False

        Next
        Dim rand As New Random

        For r = 1 To 7
            Dim tt As Boolean = False

            Dim b As Integer

            tt = True
            Do While tt = True
                b = rand.Next(0, GridView1.RowCount - 1)
                If GridView1.Item(2, b).Value = False Then
                    'With DataGridView7.Item(0, b).Value.ToString
                    'If .EndsWith("[1]") Or .EndsWith("[2]") Or .EndsWith("[3]") Or .EndsWith("[4]") Then
                    GridView1.Item(2, b).Value = True

                    tt = False
                    'End If
                    ' End With
                End If
            Loop


           

        Next



        Dim ch(GridView1.Rows.Count) As Boolean
        For i = 1 To ch.Length - 1
            ch(i) = GridView1.Item(2, i - 1).Value
            If ch(i) = True Then
                nmax = nmax + 1
            End If
        Next
        ProgressBar2.Maximum = ch.Length - 1 - nmax


        For i = 1 To ch.Length - 1
            If ch(i) = True Then
                Array.Resize(a, fi + 1)
                a(fi) = locus(i)


                fi = fi + 1
            End If
        Next
        Dim spi(a.Length - 1) As splits


        i = 1

        For i = N + 1 To sp.nOTUs.GetLength(0) - 1
            sp.nOTUs(i, 0) = newa.rewritesplits(sp.nOTUs(i, 1), N)

        Next
        For i = 1 To a.Length - 1

            spi(i) = NJ(conca(a, fi - 1, TextBox12.Text, TextBox13.Text, avstates, False), Module1.splitestmaker, 0, False, False, False)
            For ii = N + 1 To sp.nOTUs.GetLength(0) - 1
                spi(i).nOTUs(ii, 0) = newa.rewritesplits(spi(i).nOTUs(ii, 1), N)

            Next
        Next
        For j = sp.otumat1.GetLength(0) To sp.nOTUs.GetLength(0) - 2
            If sp.notus1(j, 0) > 10 ^ -9 Then

                Dim incomp As Integer = 0
                For k = 1 To spi.Length - 1
                    For m = spi(1).otumat1.GetLength(0) To spi(k).nOTUs.GetLength(0) - 2
                        If spi(k).notus1(m, 0) > 10 ^ -9 Then
                            If checkcomp(sp.nOTUs(j, 0), spi(k).nOTUs(m, 0), sp.otumat1.GetLength(0) - 1) = False Then
                                incomp = incomp + 1
                                Exit For
                            End If
                        End If

                    Next
                Next

                sp.notus1(j, 1) = incomp
            End If
        Next

    End Sub
    Function NJborrar(ByVal otumat1(,) As String, ByVal splitest(,) As String, ByVal nrep As Integer, ByVal bionj As Boolean, ByVal support As Boolean, ByVal suppforsplitest As Boolean) As splits


        nseq = otumat1.GetLength(0) - 2

        lseq = otumat1(1, 1).ToString.Length
        Dim a(1) As Integer


        lseqred = otumat1(1, 1).ToString.Length
        Dim sp3 As dsx
        sp3 = NJgo(otumat1, bionj, False) 'calculates the distance matrix


        Dim sp4 As splits
        sp4 = NJproc(nseq + 1, sp3.distancias, otumat1, False, nrep, sp3.Vij, bionj) 'Makes the splits

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
    Public Function getfirstFilename()
        Return GridView1.Item(1, 0).Value
    End Function
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub AnalizeExternalTreesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalizeExternalTreesToolStripMenuItem.Click
        Dialog3.Show()
    End Sub

    Private Sub BatchExternalexeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BatchExternalexeToolStripMenuItem.Click
        Dialog4.ShowDialog()

    End Sub

    Private Sub SelectTheseLociToolStripMenuItem_MouseUp(sender As Object, e As MouseEventArgs) Handles SelectTheseLociToolStripMenuItem.MouseUp

    End Sub
End Class
