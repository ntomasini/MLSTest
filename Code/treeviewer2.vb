Imports System.Drawing.Graphics
Imports System.Math
Public Structure nodeap
    Dim bvisible As Boolean
    Dim xb As Single
    Dim yb As Single
    Dim bcolor As Brush
    Dim pen As Pen
    Dim fontb As Font
    Dim fontg As Font
    Dim gpen As Pen

End Structure
Public Structure taxon
    Dim tvisible As Boolean
    Dim tcolor As Brush
    Dim tfont As Font
End Structure
Public Structure treemat
    Dim nodo() As String
    Dim splitsize() As Integer
    Dim x(,) As Double
    Dim y(,) As Double
End Structure
Public Class treeviewer2

    Private _sptree As splits
    Private _nwktree As String
    Private _GRUBURST As reburst
    Private _xypoints(,) As Double
    Private treematrix As treemat
    Private UPGMA As Boolean
    Private copyT As taxon
    Private copyN As nodeap
    Private group As Integer
    Private color1, color2, color3, color4, color5, color6 As Color
    Private viewsupport As Boolean
    Private support() As Single
    Private supportnames() As String
    'position of tree on page
    Private seqburst As Boolean
    Private showit() As Boolean


    Sub New()

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        _nwktree = ""
        _GRUBURST = Nothing
        _xypoints = Nothing
        UPGMA = False
        ReDim isplitests(Module1.splitestmaker.GetLength(0) - 2)
        color1 = Color.Black
        color2 = Color.Black
        color3 = Color.Black
        color4 = Color.Black
        color5 = Color.Black
        color6 = Color.Black


        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Public Property _treemat() As treemat
        Get
            Return treematrix
        End Get
        Set(ByVal value As treemat)
            treematrix = value
        End Set
    End Property
    Public Property _group() As Integer
        Get
            Return group
        End Get
        Set(ByVal value As Integer)

            group = value
        End Set
    End Property
    Public Property _showit() As Boolean()
        Get
            Return showit
        End Get
        Set(ByVal value As Boolean())

            showit = value
        End Set
    End Property

    Public Property _drawnodenames() As Boolean
        Get
            Return drawnodenames
        End Get
        Set(ByVal GRAPH As Boolean)
            drawnodenames = GRAPH
        End Set
    End Property
    Public Property _viewsupport() As Boolean
        Get
            Return viewsupport
        End Get
        Set(ByVal GRAPH As Boolean)
            viewsupport = GRAPH
        End Set
    End Property
    Public Property _isplitests() As Single()
        Get
            Return isplitests
        End Get
        Set(ByVal value() As Single)
            isplitests = value
        End Set
    End Property
    Public Property _nodes() As nodeap()
        Get
            Return nodes
        End Get
        Set(ByVal value() As nodeap)
            nodes = value
        End Set
    End Property
    Public Property _taxa() As taxon()
        Get
            Return taxa
        End Get
        Set(ByVal value() As taxon)
            taxa = value
        End Set
    End Property
    Public Property xypoints() As Double(,)
        Get
            Return _xypoints
        End Get
        Set(ByVal GRAPH As Double(,))
            _xypoints = GRAPH
        End Set
    End Property
    Public Property _upgma() As Boolean
        Get
            Return UPGMA
        End Get
        Set(ByVal AAS As Boolean)
            UPGMA = AAS
        End Set
    End Property
    Public Property GRUBURST() As reburst
        Get
            Return _GRUBURST
        End Get
        Set(ByVal GRAPH As reburst)
            _GRUBURST = GRAPH
        End Set
    End Property
    Public Property _seqburst() As Boolean
        Get
            Return seqburst
        End Get
        Set(ByVal tree As Boolean)
            seqburst = tree
        End Set
    End Property
    Public Property sptree() As splits
        Get
            Return _sptree
        End Get
        Set(ByVal tree As splits)
            _sptree = tree
        End Set
    End Property
    Dim minx, miny, maxx, maxy As Single
    Private Sub PictureBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        PictureBox1.Focus()
        If Panel2.Visible = False Then
            Panel2.Visible = True
        Else
            'Panel2.Visible = False
        End If


        If CheckBox4.Checked = True Then
            Dim ss As Point

            ss = PictureBox1.PointToClient(Cursor.Position)


            minx = (ss.X - pixFromLeft) / x_factor
            miny = (ss.Y - 10 - pixFromTop) / y_factor
            maxx = (ss.X - pixFromLeft) / x_factor
            maxy = (ss.Y + 10 - pixFromTop) / y_factor

            With treematrix
                For i = 1 To treematrix.x.GetLength(0) - 4

                    If .x(i, 0) <= minx And .x(i, 1) > maxx And .y(i, 1) >= miny And .y(i, 1) < maxy Then
                        sptree = rerooting(treematrix, i)

                        treematrix = makeSbMat(sptree, i)
                        draw()

                        Exit For
                    End If
                Next i
            End With

        End If
        If _GRUBURST.gs IsNot Nothing Then
            Dim ss As Point

            ss = PictureBox1.PointToClient(Cursor.Position)
            selected = 0
            selgroupi = 0
            minx = (ss.X + 10)
            miny = (ss.Y - 10)
            maxx = (ss.X - 10)
            maxy = (ss.Y + 10)

            With _GRUBURST.gs(TextBox1.Text - 1)
                For i = 1 To .matg.GetLength(0) - 1
                    'If i = 9 Then Stop

                    If .matg1(i, 0) <= minx And .matg1(i, 0) > maxx And .matg1(i, 1) >= miny And .matg1(i, 1) < maxy Then

                        selgroupi = i

                        selected = 2


                    End If

                Next i
            End With


        End If

        If _sptree.nOTUs IsNot Nothing Then
            Dim ss As Point
            ss = PictureBox1.PointToClient(Cursor.Position)
            selected = 0
            selgroupi = 0
            minx = (ss.X + 2 - pixFromLeft) / x_factor
            miny = (ss.Y - 10 - pixFromTop) / y_factor
            maxx = (ss.X - 2 - pixFromLeft) / x_factor
            maxy = (ss.Y + 10 - pixFromTop) / y_factor

            With treematrix
                For i = 1 To treematrix.x.GetLength(0) - 4

                    If .x(i, 0) <= minx And .x(i, 1) > maxx And .y(i, 1) >= miny And .y(i, 1) < maxy Then

                        selgroupi = i

                        splitestindex = 0
                        selected = 2
                        Dim splitests(,) As String = Module1.splitestmaker
                        For j = 1 To splitests.GetLength(0) - 2
                            If splitests(j, 1) = sptree.nOTUs(i, 2) Then

                                splitestindex = j
                                Exit For

                            End If

                        Next

                    ElseIf .x(i, 0) <= minx And .y(i, 1) >= miny And .y(i, 1) < maxy Then
                        Try
                            selstrain = _sptree.nOTUs(i, 1)
                            StrainsToolStripMenuItem.Visible = True

                        Catch
                        End Try
                    End If
                Next i
            End With

        End If


        draw()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If taxa IsNot Nothing Then
            For i = 1 To _sptree.otumat1.GetLength(0) - 1
                If taxa(i).tfont.Size > 4 Then
                    taxa(i).tfont = New Font(taxa(i).tfont.Name, taxa(i).tfont.Size - 1, taxa(i).tfont.Style)
                End If
            Next
        Else
            fontsize = fontsize - 1
        End If
        draw()

    End Sub

    Private Sub RepeatButton1_Click()
        moveU()
        'RepeatButton1.Interval = RepeatButton1.Interval / 1.8
    End Sub

    Private Sub RepeatButton2_Click()

        moveD()
        'RepeatButton2.Interval = RepeatButton2.Interval / 1.8
    End Sub

    Private Sub RepeatButton3_Click()



        y_factor1 = y_factor1 + 0.05
        pixFromTop = pixFromTop - 1
        draw()
        'RepeatButton3.Interval = RepeatButton3.Interval / 1
    End Sub

    Private Sub RepeatButton4_Click()


        y_factor1 = y_factor1 - 0.05
        pixFromTop = pixFromTop + 1
        draw()

        'RepeatButton4.Interval = RepeatButton4.Interval - 1
    End Sub

    Private Sub RepeatButton6_Click()

        scalexy = scalexy + 10
        scalexy1 = scalexy1 + 10
        If _GRUBURST.gs IsNot Nothing Then selected = 0
        If selected = 0 Then

            maxWidth = maxWidth + 0.05
        ElseIf selected = 1 Then
            nodes(selgroupi).gpen.Width = nodes(selgroupi).gpen.Width + 1
        ElseIf selected = 2 Then
            For i = 1 To treematrix.nodo.GetLength(0) - 3
                If treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
                    nodes(i).pen.Width = nodes(i).pen.Width + 0.2
                End If

            Next

        End If
        draw()

        ' RepeatButton6.Interval = RepeatButton6.Interval / 1.5
    End Sub

    Private Sub RepeatButton5_Click()

        scalexy = scalexy - 10
        scalexy1 = scalexy1 - 10
        If _GRUBURST.gs IsNot Nothing Then selected = 0
        If selected = 0 Then
            maxWidth = maxWidth - 0.05
        ElseIf selected = 1 Then
            nodes(selgroupi).gpen.Width = nodes(selgroupi).gpen.Width - 0.2
        ElseIf selected = 2 Then
            If _GRUBURST.gs Is Nothing Then
                For i = 1 To treematrix.nodo.GetLength(0) - 3
                    If treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
                        nodes(i).pen.Width = nodes(i).pen.Width - 0.2
                    End If

                Next
            End If
        End If
        draw()
        'RepeatButton5.Interval = RepeatButton5.Interval / 1.5

    End Sub

    Private Sub RepeatButton7_Click()


        moveL()
        'RepeatButton7.Interval = RepeatButton7.Interval / 1.8
    End Sub

    Private Sub RepeatButton8_Click()
        moveR()
        'RepeatButton8.Interval = RepeatButton8.Interval / 1.8
    End Sub


    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CheckBox1.Checked = False Then

            viewlabels = False

        Else

            viewlabels = True

        End If
        draw()
    End Sub

    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click



    End Sub

    Public sized As Boolean = True

    Private Sub Panel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel1.Click

    End Sub


    Private Sub Panel1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        sized = False




    End Sub
    Sub draw()
        If sptree.nOTUs IsNot Nothing Then
            Panel1.CreateGraphics.Clear(Color.White)
            Dim bit As Bitmap = New Bitmap(Panel1.Width, Panel1.Height)
            bit = drawTree(bit, treematrix)

            Dim point As New PointF(1, 1)


            PictureBox1.BackgroundImage = bit
        ElseIf _GRUBURST.gs IsNot Nothing Then
            Panel1.CreateGraphics.Clear(Color.White)
            Dim bit As Bitmap = New Bitmap(Panel1.Width, Panel1.Height)
            group = TextBox1.Text - 1
            If group > _GRUBURST.gs.Length - 2 Or group < 0 Then

                group = _GRUBURST.gs.Length - 2
                TextBox1.Text = group + 1

            End If

            bit = draweburst(bit, group)

            PictureBox1.BackgroundImage = bit
        Else
            Dim bit As Bitmap = New Bitmap(Panel1.Width, Panel1.Height)
            bit = drawmds(bit)
            PictureBox1.BackgroundImage = bit
        End If
    End Sub
    Public Function drawmds(ByVal bit As Bitmap) As Bitmap
        Dim myGraphics As Graphics = Graphics.FromImage(bit)
        myGraphics.Clear(Color.White)
        Dim xyp(_xypoints.GetLength(0) - 1, 3) As Double


        Dim minx As Double = findminormax(_xypoints, 1, True)
        Dim maxy As Double = findminormax(_xypoints, 2, False)
        Dim miny As Double = findminormax(_xypoints, 2, True)
        Dim maxx As Double = findminormax(_xypoints, 1, False)
        Dim maxyl As Double = absval(miny - maxy)
        Dim maxxl As Double = absval(minx - maxx)
        Dim scale As Double
        If maxxl > maxyl Then
            scale = (bit.Width + scalexy) / maxxl
        Else
            scale = (bit.Height + scalexy) / maxyl
        End If
        For i = 1 To _xypoints.GetLength(0) - 1
            xyp(i, 0) = _xypoints(i, 0)
            xyp(i, 1) = absval(minx - _xypoints(i, 1))
            xyp(i, 2) = -absval(miny - _xypoints(i, 2))
        Next
        Dim pen1 As New Pen(Color.Black, linewidth)
        For i = 1 To _xypoints.GetLength(0) - 1
            If _xypoints(i, 0) <> 0 Then
                Dim a As New Rectangle((xyp(i, 1) * scale) - (SIZEST / 2) + pixFromLeft, (xyp(i, 2) * scale) + pixFromTop - (SIZEST / 2), SIZEST, SIZEST)
                myGraphics.DrawEllipse(pen1, a)
                If viewlabels = True Then

                    Dim aFont As New System.Drawing.Font("Arial", fontsize, FontStyle.Regular)
                    myGraphics.DrawString(xyp(i, 0), aFont, Brushes.Black, (xyp(i, 1) * scale) + 3 + pixFromLeft, (xyp(i, 2) * scale) + pixFromTop - 3)
                End If

            End If
        Next
        Dim P1, P2, P3 As Point
        P1 = New Point(pixFromLeft - 0.05 * bit.Width, 0 + pixFromTop + 0.05 * bit.Height)
        P2 = New Point(pixFromLeft - 0.05 * bit.Width, -(maxy - miny) * scale + pixFromTop - 0.05 * bit.Height)
        P3 = New Point((maxx - minx) * scale + pixFromLeft, 0 + pixFromTop + 0.05 * bit.Height)

        myGraphics.DrawLine(pen1, P1, P2)
        myGraphics.DrawLine(pen1, P1, P3)

        Return bit
    End Function
    Function distancetomin(ByVal mm As Double, ByVal xy As Double) As Double
        Dim distan As Double

        distan = absval(mm - xy)



    End Function
    Function absval(ByVal n As Double) As Double
        If n < 0 Then
            n = n * -1
        End If
        Return n
    End Function
    Function findminormax(ByVal arr(,) As Double, ByVal vari As Integer, ByVal mintrue As Boolean) As Double
        Dim res As Double
        If mintrue = True Then
            res = 1000000
            For i = 1 To arr.GetLength(0) - 1
                If arr(i, vari) < res Then
                    res = arr(i, vari)
                End If
            Next
        End If
        If mintrue = False Then
            res = -1000000
            For i = 1 To arr.GetLength(0) - 1
                If arr(i, vari) > res Then
                    res = arr(i, vari)
                End If
            Next
        End If
        Return res
    End Function

    Public Function draweburst(ByVal bit As Bitmap, ByVal i As Integer) As Bitmap
        Dim myGraphics As Graphics = Graphics.FromImage(bit)
        myGraphics.Clear(Color.White)
        Dim center(1) As Single

        center(0) = (bit.Width / 2) + pixFromLeft
        center(1) = (bit.Height / 2) + pixFromTop
        Dim centralength As Double = (bit.Height + scalexy) / 3




        _GRUBURST.gs(i).centerx = center(0)
        _GRUBURST.gs(i).centery = center(1)

        ReDim _GRUBURST.gs(i).matg1(_GRUBURST.gs(i).matg.GetLength(0) - 1, _GRUBURST.gs(i).matg.GetLength(1) - 1)
        Dim fi As Integer = founderindex(_GRUBURST.gs(i).matg, _GRUBURST.gs(i).founder)
        If fi <> 0 Then
            _GRUBURST.gs(i).matg1(fi, 0) = _GRUBURST.gs(i).centerx
            _GRUBURST.gs(i).matg1(fi, 1) = _GRUBURST.gs(i).centery

        End If
        Dim ANGLE = 2 * PI / _GRUBURST.gs(i).initialsize
        Dim ANGLE2 = 2 * PI / (_GRUBURST.gs(i).matg.GetLength(0) - 1 - _GRUBURST.gs(i).initialsize)
        For j = 1 To _GRUBURST.gs(i).matg.GetLength(0) - 1
            Dim brlength As Single
            If j <= _GRUBURST.gs(i).initialsize Then
                brlength = centralength / 4
                If _GRUBURST.gs(i).matg1(j, 0) = 0 Then
                    _GRUBURST.gs(i).matg1(j, 0) = nodes(j).xb + _GRUBURST.gs(i).centerx + (Sin(ANGLE * (j - 1)) * brlength)
                    _GRUBURST.gs(i).matg1(j, 1) = nodes(j).yb + _GRUBURST.gs(i).centery - (Cos(ANGLE * (j - 1)) * brlength)

                End If

            Else
                brlength = centralength / 2
                If _GRUBURST.gs(i).matg1(j, 0) = 0 Then
                    _GRUBURST.gs(i).matg1(j, 0) = nodes(j).xb + _GRUBURST.gs(i).centerx + (Sin(ANGLE2 * (j - 1)) * brlength)
                    _GRUBURST.gs(i).matg1(j, 1) = nodes(j).yb + _GRUBURST.gs(i).centery - (Cos(ANGLE2 * (j - 1)) * brlength)

                End If
            End If


        Next




        Dim s As Integer = selgroupi
        Dim a1 As New Drawing.RectangleF(_GRUBURST.gs(i).matg1(s, 0) - (SIZEST * _GRUBURST.gs(i).matg(s, 1) / 2 + 7), _GRUBURST.gs(i).matg1(s, 1) - (SIZEST * _GRUBURST.gs(i).matg(s, 1) / 2 + 7), SIZEST * _GRUBURST.gs(i).matg(s, 1) + 7, SIZEST * _GRUBURST.gs(i).matg(s, 1) + 7)
        ' myGraphics.FillEllipse(Brushes.Yellow, a1)
        For j = 1 To _GRUBURST.gs(i).matg.GetLength(0) - 1
            Dim max As Integer

            Dim a As New Drawing.RectangleF(_GRUBURST.gs(i).matg1(j, 0) - (SIZEST * _GRUBURST.gs(i).matg(j, 1) / 2), _GRUBURST.gs(i).matg1(j, 1) - (SIZEST * _GRUBURST.gs(i).matg(j, 1) / 2), SIZEST * _GRUBURST.gs(i).matg(j, 1), SIZEST * _GRUBURST.gs(i).matg(j, 1))
            Dim aFont As New System.Drawing.Font("Arial", fontsize, FontStyle.Regular)



            If j = s Then
                myGraphics.FillEllipse(Brushes.LightGreen, a)
            Else
                myGraphics.FillEllipse(Brushes.Blue, a)
            End If
            If viewlabels = True Then
                myGraphics.DrawString(_GRUBURST.gs(i).matg(j, 0), aFont, Brushes.Black, _GRUBURST.gs(i).matg1(j, 0) - 15, _GRUBURST.gs(i).matg1(j, 1))

            End If

            'myGraphics.DrawLine(Pens.Blue, __GRUBURST.gs(i).centerx, __GRUBURST.gs(i).centery, __GRUBURST.gs(i).matg1(j, 0), __GRUBURST.gs(i).matg1(j, 1))

            If _GRUBURST.matx(_GRUBURST.gs(i).matg(j, 0), 2) <> " " Then
                For k = 1 To _GRUBURST.matx.GetLength(0) - 1
                    If _GRUBURST.matx(_GRUBURST.gs(i).matg(j, 0), 2).Contains(" " & k.ToString & " ") = True Then
                        Dim ind As Integer = founderindex(_GRUBURST.gs(i).matg, k)
                        Dim pen1 As New Pen(linecolor, linewidth)

                        If drawcirclegroup = True Then
                            Dim maxF As Integer = 0



                            pen1.Color = Color.Blue
                            myGraphics.DrawLine(pen1, _GRUBURST.gs(i).matg1(ind, 0), _GRUBURST.gs(i).matg1(ind, 1), _GRUBURST.gs(i).matg1(j, 0), _GRUBURST.gs(i).matg1(j, 1))



                        End If
                    End If
                Next
            End If
        Next

        Return bit
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If taxa IsNot Nothing Then
            For i = 1 To _sptree.otumat1.GetLength(0) - 1
                If taxa(i).tfont.Size < 40 Then
                    taxa(i).tfont = New Font(taxa(i).tfont.Name, taxa(i).tfont.Size + 1, taxa(i).tfont.Style)
                End If
            Next
        Else
            fontsize = fontsize + 1
        End If
        draw()

    End Sub

    Private Sub Button5_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click

        scaleBarLength = scaleBarLength / 10
        SIZEST = SIZEST - 1
        draw()
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click

        scaleBarLength = scaleBarLength * 10
        SIZEST = SIZEST + 1
        draw()
    End Sub




    Public ancho As Double = 1




    Private Sub treeviewer2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        fontsize = 10
        maxWidth = 500
        y_factor = 15
        Spacegroups = 75
        SIZEST = 2
        linewidth = 1
        fontsizeb = 7
        pixFromTop = 50
        scalexy = 1
        scalexy1 = 1
        Me.Visible = False
        If Panel1.Size.Width <> 0 And Panel1.Height <> 0 Then
            If sptree.nOTUs Is Nothing Then
                If GRUBURST.gs IsNot Nothing Then
                    pixFromTop = 0
                    pixFromLeft = 0
                    linecolor = Color.Blue
                    Label4.Text = "ST Length"
                    CheckBox2.Text = "View Lines"
                    TextBox1.Visible = True
                    ComboBox1.Visible = False
                    ReDim nodes(_GRUBURST.gs(0).matg.GetLength(0) - 1)
                    TextBox1.Text = 1
                    CheckBox2.Checked = True
                    Label8.Visible = True
                    Label8.Text = "Group"

                Else
                    pixFromTop = Panel1.Height - 20
                    pixFromLeft = 20
                    SIZEST = 2
                    Label4.Text = "Dot size"
                    CheckBox2.Visible = False
                    TextBox1.Visible = False
                    Label8.Visible = False
                    ComboBox1.Visible = False
                    CheckBox5.Visible = False
                End If

                Label5.Visible = False
                Button23.Visible = False
                Button24.Visible = False
            Else

                ReDim taxa(_sptree.otumat1.GetLength(0) - 1)
                For i = 1 To _sptree.otumat1.GetLength(0) - 1
                    taxa(i).tfont = New Font("Arial", 9, FontStyle.Regular)
                    taxa(i).tcolor = Brushes.Black
                    taxa(i).tvisible = True

                Next
                If nodes Is Nothing Then

                    ReDim nodes(_sptree.nOTUs.GetLength(0) + 1)
                End If
                For i = 1 To nodes.Length - 1

                    If nodes(i).pen Is Nothing Then
                        nodes(i).pen = New Pen(Color.Black, 1)
                    End If
                    nodes(i).fontb = New Font("Arial", 7, FontStyle.Regular)
                    nodes(i).gpen = New Pen(Color.Black, 1)
                    nodes(i).fontg = New Font("Arial", 10, FontStyle.Regular)
                    nodes(i).bcolor = Brushes.Black
                Next
                CheckBox2.Checked = viewsupport
                If seqburst = True Then
                    CheckBox2.Text = "Group Definition"
                    Label5.Visible = False
                    TextBox1.Visible = False
                    ComboBox1.Visible = False
                    Label8.Visible = False
                    treematrix = makeSbMat(sptree, Nothing)
                Else
                    ComboBox1.Visible = True
                    Dim count As Integer
                    If supportnames Is Nothing Then
                        count = -1
                    Else
                        count = supportnames.Length - 1
                    End If
                    If count >= 0 Then
                        ReDim support(count)
                        ReDim showit(count)
                        For i = 0 To count

                            ComboBox1.Items.Add(supportnames(i))
                            showit(i) = True
                            support(i) = -1
                        Next
                        TextBox1.Text = -1
                        ComboBox1.SelectedIndex = 0
                        Button17.Visible = True
                    Else
                        ComboBox1.Visible = False
                        TextBox1.Visible = False
                        Label8.Visible = False
                        CheckBox5.Visible = False
                        Button17.Visible = False
                    End If

                    If treematrix.nodo Is Nothing Then
                        treematrix = makeSbMat(sptree, 27)
                        If Form1.ToolStripComboBox1.SelectedIndex <> 0 And Form1.ToolStripComboBox1.SelectedIndex <> -1 Then
                            Dim aa As Integer
                            For i = 1 To sptree.nOTUs.GetLength(0) - 1
                                Dim AV As Integer = Form1.ToolStripComboBox1.SelectedIndex
                                If sptree.nOTUs(i, 1) = " " & Form1.ToolStripComboBox1.SelectedIndex & " " Then
                                    aa = i
                                    Exit For
                                End If
                            Next
                            If UPGMA = False Then
                                sptree = rerooting(treematrix, aa)
                                treematrix = makeSbMat(sptree, aa)
                            End If

                        End If
                    End If

                    CheckBox4.Visible = True
                    CheckBox3.Visible = True
                    Label9.Visible = True
                    Button7.Visible = True
                    Button8.Visible = True
                    PictureBox1.ContextMenuStrip = ContextMenuStrip2
                    Button6.Visible = True

                End If
                Button9.Visible = True


                maxWidth = 1
            End If
            draw()

        End If
        Me.Visible = True


    End Sub



    Private Sub treeviewer1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        If sized = False Then
            If Panel1.Size.Width <> 0 And Panel1.Height <> 0 Then

                If _GRUBURST.gs IsNot Nothing Then
                    pixFromTop = 0
                    pixFromLeft = 0
                End If
                draw()
                sized = False
            End If



        End If

    End Sub

    Private Sub treeviewer1_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin
        ancho = Panel1.Size.Width
    End Sub

    Private Sub treeviewer1_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        If sized = False Then
            If Panel1.Size.Width <> 0 And Panel1.Height <> 0 Then

                If _GRUBURST.gs IsNot Nothing Then
                    pixFromTop = 0
                    pixFromLeft = 0
                End If
                draw()
                sized = False
            End If



        End If
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim props1 As New props With {.gruburst = _GRUBURST, .xypoints = _xypoints, _
                                      .y_factor = _y_factor, .maxWidth = _MaxWidth, .pixFromLeft = _pixFromLeft, _
                                      .pixFromTop = _pixFromTop, .viewlabels = viewlabels, _
                                      ._linecolor = linecolor, ._iHeight = Panel1.Height, ._drawcg = drawcirclegroup, ._iwidth = Panel1.Width, _
                                       ._scalexy = scalexy, ._spacegroups = Spacegroups, ._isplitests = isplitests, _
._drawnodenames = drawnodenames, ._nodes = nodes, ._sbl = scaleBarLength, ._taxa = taxa, .fontsize = fontsize, ._linewidth = linewidth, ._sizeST = SIZEST}
        props1._group = group
        props1._showit = showit
        props1.treematx = treematrix
        props1.sptree = sptree
        props1.Show()
        props1.cutoffb = support


    End Sub

    Friend WithEvents botones As System.Windows.Forms.Button

    Private Sub CheckBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.Click
        If CheckBox3.Checked = False Then

            drawnodenames = False

        Else
            drawnodenames = True

        End If
        draw()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "."c Or e.KeyChar = "-"c Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If Me.Visible = True Then
            If TextBox1.Text = "" Then
                If ComboBox1.Visible = True Then
                    support(ComboBox1.SelectedIndex) = -1
                Else
                    TextBox1.Text = 1
                End If

            Else


                If _GRUBURST.gs IsNot Nothing And TextBox1.Text <> "0" Then
                    selgroupi = 0
                    If _GRUBURST.gs.Length > TextBox1.Text Then

                        ReDim nodes(_GRUBURST.gs(TextBox1.Text - 1).matg.GetLength(0) - 1)

                        draw()

                    Else
                        TextBox1.Text = 1
                    End If
                Else
                    If IsNumeric(TextBox1.Text) = False Then
                        support(ComboBox1.SelectedIndex) = -1
                    Else
                        support(ComboBox1.SelectedIndex) = TextBox1.Text
                    End If

                End If
            End If

            draw()
        End If
    End Sub


    Public x, y As Single

    Private Sub PictureBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.DoubleClick
        selected = 0
        selgroupi = 0
    End Sub
    Private press As Boolean
    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        x = MousePosition.X
        y = MousePosition.Y
        press = True


    End Sub
    Private posL As Integer
    Private posT As Integer

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Dim x1 = MousePosition.X
        Dim y1 = MousePosition.Y

        If press = True Then
            If selected = 2 Then
                'If nodes(selgroupi).xb < 1000 Then Stop

                If _GRUBURST.gs IsNot Nothing Then



                    nodes(selgroupi).xb = nodes(selgroupi).xb + (x1 - x)
                    nodes(selgroupi).yb = nodes(selgroupi).yb + (y1 - y)



                Else
                    If nodes(selgroupi).bvisible = True And _sptree.notus1(selgroupi, 1) <> Nothing Then
                        nodes(selgroupi).xb = nodes(selgroupi).xb + MousePosition.X - x
                        nodes(selgroupi).yb = nodes(selgroupi).yb + MousePosition.Y - y
                    End If
                End If


            ElseIf selected = 1 Then
                isplitests(splitestindex) = isplitests(splitestindex) + MousePosition.X - x
            Else
                pixFromLeft = pixFromLeft + MousePosition.X - x
                pixFromTop = pixFromTop + MousePosition.Y - y
            End If
            x = MousePosition.X
            y = MousePosition.Y

            If GRUBURST.gs Is Nothing Then
                draw()
            End If
        End If


    End Sub

    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        press = False
        If _GRUBURST.gs IsNot Nothing Then
            draw()
        End If
    End Sub

    '-------------------------------------


    'position of tree on page
    Private pixFromLeft As Double = 50  'pixels from left edge of page
    Private pixFromTop As Double = 50  'pixels from top of page

    'position and size of text boxes with names in
    'all in pixels
    Private viewlabels As Boolean = True


    ' Private tbheight As Double= 25      'text box height

    Private x_spacing As Double = 3      'distance between branch end and name box
    Private colordefuente As New SolidBrush(Color.Black)
    Private colordefuenteboo As New SolidBrush(Color.Black)
    Private drawcirclegroup As Boolean = True
    Private fontsize As Integer = 10
    Private fontsizeb As Integer = 7
    Private linewidth As Integer = 1
    Private linecolor As Color = Color.Black
    Private scalexy As Integer = 1
    Private scalexy1 As Integer = 1
    Private SIZEST As Integer = 2
    'scaling the tree
    Private x_factor As Double = 200    'to multiply idstances by
    Private y_factor As Double = 15 'vertical spacing between names
    Private y_factor1 As Double = 1
    Private scaleBarLength As Double = 0.001   'length of scale bar as distance value
    Private maxWidth As Double = 1       'max width of tree (inches)

    Private ishowsupport() As Boolean

    'options
    Private bootstrap As Boolean = True
    Private realDistances As Boolean = True
    Private drawnodenames As Boolean = False
    'Private cutoff As Single
    Private f As Font
    Private save As Boolean = False
    Private xyMAT(,) As Double
    Private nodeMAT(,) As Integer
    Private out As Integer
    Private Spacegroups As Single
    Private selected As Integer
    Private isplitests() As Single

    Private nodes() As nodeap
    Private taxa() As taxon
    Dim stack() As Integer, top As Integer
    Public Property _spacegroups() As Single
        Get
            Return Spacegroups
        End Get
        Set(ByVal value As Single)
            Spacegroups = value
        End Set
    End Property
    Public Property _fontsize() As Integer
        Get
            Return fontsize
        End Get
        Set(ByVal GRAPH As Integer)
            fontsize = GRAPH
        End Set
    End Property
    Public Property _viewlabels() As Boolean
        Get
            Return viewlabels
        End Get
        Set(ByVal GRAPH As Boolean)
            viewlabels = GRAPH
        End Set
    End Property
    Public Property _drawcg() As Boolean
        Get
            Return drawcirclegroup
        End Get
        Set(ByVal GRAPH As Boolean)
            drawcirclegroup = GRAPH
        End Set
    End Property
    Public Property _SIZEST() As Integer
        Get
            Return SIZEST
        End Get
        Set(ByVal GRAPH As Integer)
            SIZEST = GRAPH
        End Set
    End Property

    Public Property _sbl() As Double
        Get
            Return scaleBarLength
        End Get
        Set(ByVal GRAPH As Double)
            scaleBarLength = GRAPH
        End Set
    End Property
    Public Property _MaxWidth() As Single
        Get
            Return maxWidth
        End Get
        Set(ByVal GRAPH As Single)
            maxWidth = GRAPH
        End Set
    End Property
    Public Property _y_factor() As Single
        Get
            Return y_factor
        End Get
        Set(ByVal GRAPH As Single)
            y_factor = GRAPH
        End Set
    End Property
    Public Property _pixFromTop() As Single
        Get
            Return pixFromTop
        End Get
        Set(ByVal GRAPH As Single)
            pixFromTop = GRAPH
        End Set
    End Property
    Public Property _cutoffz() As Single()
        Get
            Return support
        End Get
        Set(ByVal GRAPH As Single())
            support = GRAPH
        End Set
    End Property
    Public Property _supportnames() As String()
        Get
            Return supportnames
        End Get
        Set(ByVal GRAPH As String())
            supportnames = GRAPH
        End Set
    End Property
    Public Property _fontsizeb() As Single
        Get
            Return fontsizeb
        End Get
        Set(ByVal GRAPH As Single)
            fontsizeb = GRAPH
        End Set
    End Property
    Public Property _linewidth() As Single
        Get
            Return linewidth
        End Get
        Set(ByVal GRAPH As Single)
            linewidth = GRAPH
        End Set
    End Property
    Public Property _scalexy() As Single
        Get
            Return scalexy
        End Get
        Set(ByVal GRAPH As Single)
            scalexy = GRAPH
        End Set
    End Property
    Public Property _pixFromLeft() As Single
        Get
            Return pixFromLeft
        End Get
        Set(ByVal GRAPH As Single)
            pixFromLeft = GRAPH
        End Set
    End Property
    Public Property _colordefuente() As SolidBrush
        Get
            Return colordefuente
        End Get
        Set(ByVal GRAPH As SolidBrush)
            colordefuente = GRAPH
        End Set
    End Property
    Public Property _colordefuenteboo() As SolidBrush
        Get
            Return colordefuenteboo
        End Get
        Set(ByVal GRAPH As SolidBrush)
            colordefuenteboo = GRAPH
        End Set
    End Property
    Public Property _linecolor() As Color
        Get
            Return linecolor
        End Get
        Set(ByVal GRAPH As Color)
            linecolor = GRAPH
        End Set
    End Property

    Function xscale(ByVal length As Double, ByVal x(,) As Double)
        Dim max As Double
        For i = 1 To x.GetLength(0) - 1
            If x(i, 1) > max Then
                max = x(i, 1)
            End If

        Next
        If max = 0 Then
            MsgBox("ups! There is no tree here, all branch lengths are zero", MsgBoxStyle.Critical)

            xscale = 0
            Exit Function
        End If

        xscale = (length * maxWidth - length / 2.468) / max
    End Function
    Function yscale(ByVal length As Double, ByVal y(,) As Double) As Double
        Dim max As Double
        For i = 1 To y.GetLength(0) - 1
            If y(i, 2) > max Then
                max = y(i, 2)
            End If

        Next

        yscale = (length * y_factor1 - length / 6.17) / max
    End Function

    Function CHECKOUT(ByVal SPLIT1 As String, ByVal split2 As String) As Boolean

        Dim s As Boolean
        If SPLIT1 = split2 Then
            s = False
            Return s
            Exit Function
        End If
        Dim a() As String = SPLIT1.Split(" "c)
        Dim b() As String = split2.Split(" "c)

        For i = 1 To a.Length - 2



            If b.IndexOf(b, a(i)) <> -1 Then

                s = True
                Exit For
            End If

        Next
        Return s
    End Function
    Function compareexp(ByVal button As Object, ByVal valueS As Object, ByVal valueT As Object) As Object
        Select Case button.text
            Case Is = "="
                If valueS = valueT Then
                    Return valueS
                End If
            Case Is = ">="
                If valueS >= valueT Then
                    Return valueS
                End If
            Case Is = "<="
                If valueS <= valueT Then
                    Return valueS
                End If
        End Select
        If compareexp = Nothing Then Return -1000
    End Function
    Public Function drawTree(ByVal bit As Bitmap, ByVal mat_tree As treemat) As Bitmap
        Dim count0, count1, count2, count3, countT As Integer


        x_factor = xscale(bit.Width, mat_tree.x)

        y_factor = yscale(bit.Height, mat_tree.y)

        Dim myGraphics As Graphics = Graphics.FromImage(bit)
        myGraphics.Clear(Color.White)
        If drawnodenames = True Then
            Dim splitest(,) As String = Module1.splitestmaker
            If isplitests.Length < splitest.GetLength(0) Then
                Array.Resize(isplitests, splitest.GetLength(0) - 1)
            End If
            For v = 1 To splitest.GetLength(0) - 2
                Dim index As Integer = 0
                Dim ax() As String = splitest(v, 0).Split(" "c)
                Dim makegroup As Boolean = False
                For i = 1 To sptree.nOTUs.GetLength(0) - 1

                    If sptree.nOTUs(i, 2) = splitest(v, 1) Then
                        makegroup = True
                        index = i
                        Exit For
                    End If
                Next

                If makegroup = True Then
                    Dim min As Single = 1000000
                    Dim max As Single = 0
                    Dim maxX As Single = 0
                    For i = 1 To ax.Length - 2
                        For j = 1 To sptree.nOTUs.GetLength(0) - 1
                            If sptree.nOTUs(j, 1) = " " & ax(i) & " " Then
                                If mat_tree.y(j, 1) < min Then
                                    min = mat_tree.y(j, 1)
                                End If
                                If mat_tree.y(j, 1) > max Then
                                    max = mat_tree.y(j, 1)
                                End If
                                If mat_tree.x(j, 1) > maxX Then
                                    maxX = mat_tree.x(j, 1)
                                End If
                            End If
                        Next
                    Next

                    ''''

                    Dim xg As Single
                    Dim yg1 As Single
                    Dim yg2 As Single

                    yg1 = (min * y_factor) + pixFromTop - fontsize / 2
                    yg2 = (max * y_factor) + pixFromTop + fontsize / 2
                    xg = (maxX * x_factor) + pixFromLeft + Spacegroups + isplitests(v)
                    myGraphics.DrawLine(nodes(index).gpen, xg, yg1, xg, yg2)
                    Dim xbr As New SolidBrush(nodes(index).gpen.Color)
                    myGraphics.DrawString(Form1.DataGridView6.Item(0, v - 1).Value, nodes(index).fontg, xbr, xg + x_spacing, ((yg1 + yg2) / 2) - (fontsize))

                End If
            Next
        End If
        ''''


        'draw all the horizontal lines
        Dim s = selgroupi
        Dim y11 As Single
        Dim xl1 As Single
        Dim xr1 As Single

        y11 = (mat_tree.y(s, 1) * y_factor) + pixFromTop
        xl1 = (mat_tree.x(s, 0) * x_factor) + pixFromLeft
        xr1 = (mat_tree.x(s, 1) * x_factor) + pixFromLeft
        If s <> 0 Then
            Dim rect As RectangleF = New RectangleF(xl1 + 1, y11 - 3 - nodes(s).pen.Width, xr1 - xl1 - 2, 2 * (3 + nodes(s).pen.Width))

            myGraphics.FillRectangle(Brushes.LightGreen, rect)
        End If
        For j = 1 To mat_tree.x.GetLength(0) - 1
            Dim y As Single
            Dim xl As Single
            Dim xr As Single

            y = (mat_tree.y(j, 1) * y_factor) + pixFromTop
            xl = (mat_tree.x(j, 0) * x_factor) + pixFromLeft
            xr = (mat_tree.x(j, 1) * x_factor) + pixFromLeft
            myGraphics.DrawLine(nodes(j).pen, xr, y, xl, y)
            'write name
            If mat_tree.nodo.Contains(mat_tree.nodo(j) & "1") = False Then
                Dim idtax As Integer = sptree.nOTUs(j, 1)
                If viewlabels = True And taxa(idtax).tvisible = True Then
                    Dim aFont As Font = taxa(idtax).tfont



                    myGraphics.DrawString(sptree.otumat1(idtax, 0), aFont, taxa(idtax).tcolor, xr + x_spacing, y - taxa(idtax).tfont.Size)

                End If
            End If
        Next j



        For j = 1 To mat_tree.y.GetLength(0) - 1
            Dim x As Single
            Dim y1 As Single
            Dim y2 As Single

            y1 = (mat_tree.y(j, 0) * y_factor) + pixFromTop
            y2 = (mat_tree.y(j, 2) * y_factor) + pixFromTop
            x = (mat_tree.x(j, 1) * x_factor) + pixFromLeft
            myGraphics.DrawLine(nodes(j).pen, x, y1, x, y2)


        Next j


        ' myGraphics.DrawRectangle(Pens.Green, xl1, y11 - 4, xr1 - xl1, 8)
        'myGraphics.DrawLine(Pens.Red, xr1, y11, xl1, y11)








        'write bootstrap values

        For j = 0 To mat_tree.x.GetLength(0) - 3
            If nodes(j).bvisible = True Then
                If sptree.notus1(j, 0) > 0 Then
                    If sptree.notus1(j, 1) <> Nothing Then
                        Dim vals() As String
                        vals = sptree.notus1(j, 1).Split("/")


                        If evalsupport(vals, support) Then
                            Dim value As String = Nothing

                            For i = 0 To vals.Length - 1
                                If value Is Nothing Then
                                    If showit(i) = True Then value = vals(i)

                                Else
                                    If showit(i) = True Then value = value & "/" & vals(i)
                                End If
                            Next
                            Dim x As Double, y As Double
                            If value <> Nothing Then
                                Dim a As Single = value.ToString.Count * 0.9
                                x = (mat_tree.x(j, 1) * x_factor - nodes(j).fontb.Size * a) + pixFromLeft + nodes(j).xb
                                y = (mat_tree.y(j, 1) * y_factor) - (nodes(j).fontb.Size * 1.3) + pixFromTop + nodes(j).yb

                                'Dim aFont As New System.Drawing.Font("Arial", fontsizeb, FontStyle.Regular)


                                myGraphics.DrawString(value, nodes(j).fontb, nodes(j).bcolor, x, y)
                            End If
                        End If
                    End If
                End If
            End If
        Next j



        'draw scale Bar
        If realDistances Then
            Dim y As Single = bit.Height - bit.Height / 60
            Dim xl As Single, xr As Single


            xl = 5
            xr = (scaleBarLength * x_factor) + xl
            Dim aFont As New System.Drawing.Font("Arial", fontsizeb, FontStyle.Regular)
            Dim pen1 As New Pen(Color.Black, nodes(1).gpen.Width)

            myGraphics.DrawLine(pen1, xl, y, xr, y)
            myGraphics.DrawString(scaleBarLength, nodes(1).fontb, Brushes.Black, xr + x_spacing, y)
        End If



        Return bit


    End Function
    Function evalsupport(ByVal vals() As String, ByVal cutoff() As Single) As Boolean
        Dim l As Integer = vals.Length
        Dim v(l - 1) As Boolean
        For i = 0 To l - 1
            If vals(i) = "" Then
                v(i) = True
            Else
                Try
                    If vals(i) > cutoff(i) Then
                        v(i) = True
                    End If
                Catch
                    v(i) = True
                End Try

            End If
        Next
        If v.TrueForAll(v, AddressOf tfal) = True Then
            Return True
        Else : Return False
        End If

    End Function
    Function tfal(ByVal s As Boolean) As Boolean
        If s = True Then Return True Else Return False
    End Function


    Sub exportoT(ByVal SP As splits, ByVal treemat As treemat)
        Dim dtNOTUS As New DataTable
        Dim DTNODES As New DataTable
        Dim DTOTUMAT As New DataTable

        For c = 1 To 4
            dtNOTUS.Columns.Add()

        Next
        For cc = 1 To 2

            DTOTUMAT.Columns.Add()
        Next
        For C = 1 To 8
            DTNODES.Columns.Add()
        Next


        Dim ds As New DataSet







        Dim i As Integer = SP.nOTUs.GetLength(0) - 1
        For g = 1 To i

            dtNOTUS.Rows.Add(SP.nOTUs(g, 1), SP.nOTUs(g, 2), SP.notus1(g, 0), SP.notus1(g, 1))


        Next
        i = treemat.nodo.Length - 1
        For g = 1 To i

            DTNODES.Rows.Add(treemat.nodo(g), treemat.x(g, 0), treemat.x(g, 1), treemat.y(g, 0), treemat.y(g, 1), treemat.y(g, 2), nodes(g).pen.Color.ToArgb, nodes(g).pen.Width)


        Next

        For i = 0 To SP.otumat1.GetLength(0) - 1
            DTOTUMAT.Rows.Add(SP.otumat1(i, 0), SP.otumat1(i, 1))


        Next

        ds.Tables.Add(dtNOTUS)
        ds.Tables.Add(DTNODES)
        ds.Tables.Add(DTOTUMAT)

        Dim xx As String





        Application.DoEvents()
        SaveFileDialog1.Title = "Export TREE"
        xx = "Format " & "xml" & "|" & "*." & "xml"
        SaveFileDialog1.Filter = xx
        SaveFileDialog1.DefaultExt = "xml"
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName <> "" Then
            ds.WriteXml(SaveFileDialog1.FileName)

            ds.Dispose()
        End If




    End Sub









    Private Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim nwk As newicker
        nwk = New newicker
        Dim texto As String

        Dim r As MsgBoxResult
        r = MsgBox("Do you want to save as standard Newick format? if you click on NO the tree will be saved as a formatted MLSTest tree", MsgBoxStyle.YesNo, "Save...")
        If r = MsgBoxResult.No Then
            exportoT(_sptree, treematrix)
        Else
            texto = nwk.makenwk(sptree.notus1, sptree.nOTUs, sptree.otumat1.GetLength(0) - 1, sptree.otumat1, treematrix.nodo)
            SaveFileDialog1.Title = "Export TREE"
            Dim xx As String = "Format " & "nwk" & "|" & "*." & "nwk"
            SaveFileDialog1.Filter = xx
            SaveFileDialog1.DefaultExt = "nwk"
            SaveFileDialog1.FileName = ""
            SaveFileDialog1.ShowDialog()
            If SaveFileDialog1.FileName <> "" Then
                Dim a As New System.IO.StreamWriter(SaveFileDialog1.FileName)
                a.WriteLine(texto)
                a.Close()

            End If
        End If


        ''




    End Sub





    Public Function maketreeMat(ByVal sp As splits, ByVal outsplit As Integer) As treemat
        Dim nodo(sp.nOTUs.GetLength(0) + 1) As String
        Dim splindex(sp.nOTUs.GetLength(0) - 1) As Integer
        Dim checkedsp(sp.nOTUs.GetLength(0) - 1) As Boolean
        Dim checkedsp1(sp.nOTUs.GetLength(0) - 1) As Boolean
        Dim notus(sp.nOTUs.GetLength(0) - 1) As String
        For i = 1 To sp.nOTUs.GetLength(0) - 1
            notus(i) = sp.nOTUs(i, 2)
        Next
        Dim splitsize(sp.nOTUs.GetLength(0) - 1) As Integer
        Dim maxsize As Integer = 1000
        Dim raiz1, raiz2 As String

        Dim realgroups(sp.notus1.GetLength(0) - 1) As String
        raiz1 = makeroot(sp.nOTUs(outsplit, 2), sp.otumat1.GetLength(0) - 1, True)
        raiz2 = makeroot(sp.nOTUs(outsplit, 2), sp.otumat1.GetLength(0) - 1, False)

        Dim rep1 As String = represent(raiz1)
        Dim rep2 As String = " " & represent(raiz2)
        For i = 1 To nodo.GetLength(0) - 4

            If i <> outsplit Then

                If checkcutsite(raiz1, sp.nOTUs(i, 2)) = True Then


                    nodo(i) = "1"
                    splitsize(i) = findsplitsize(sp.nOTUs(i, 2), sp.otumat1.GetLength(0) - 1, rep2)
                    realgroups(i) = realsize(sp.nOTUs(i, 2), splitsize(i), rep2)
                Else
                    nodo(i) = "2"
                    splitsize(i) = findsplitsize(sp.nOTUs(i, 2), sp.otumat1.GetLength(0) - 1, rep1)
                    realgroups(i) = realsize(sp.nOTUs(i, 2), splitsize(i), rep1)
                    If sp.nOTUs(i, 2) = sp.nOTUs(outsplit, 2) Then
                        checkedsp1(i) = True

                    End If
                End If

            End If

        Next
        Dim idx As Integer = 1

        Do

            Dim ind As Integer = findmax(splitsize, maxsize, idx)

            If ind = 0 Or splitsize(ind) < maxsize Then
                maxsize = maxsize - 1
                idx = 1
                ind = findmax(splitsize, maxsize, idx)
                checkedsp.Clear(checkedsp, 0, nodo.Length - 2)

            Else

                Dim id As String = nodo(ind)




                If checkedsp1(ind) = False Then
                    If checkedsp(ind) = False Then



                        nodo(ind) = nodo(ind) & "1"

                        Dim max2 As Integer = 0
                        For i = 1 To nodo.GetLength(0) - 4

                            If i <> ind And splitsize(i) <= maxsize And nodo(i) = id Then


                                If checkcutsite(realgroups(ind), sp.nOTUs(i, 2)) = True Then
                                    nodo(i) = nodo(i) & "1"

                                Else
                                    nodo(i) = nodo(i) & "2"
                                    If splitsize(i) > splitsize(max2) Then max2 = i

                                End If


                                checkedsp(i) = True
                            End If
                        Next

                        checkedsp1(max2) = True



                    End If
                End If

                maxsize = splitsize(ind)
                idx = ind + 1
            End If




        Loop While maxsize > 0
        nodo(outsplit) = "1"
        nodo(nodo.Length - 3) = "2"
        sptree.nOTUs(nodo.Length - 3, 0) = raiz2
        sp.notus1(outsplit, 0) = sp.notus1(outsplit, 0) / 2
        sp.notus1(nodo.Length - 3, 0) = sp.notus1(outsplit, 0)
        Dim x(nodo.Length - 1, 1) As Double
        Dim y(nodo.Length - 1, 2) As Double
        Dim ed As Boolean = False
        Dim n As Integer = 1

        Do
            Dim xp As Double
            ed = True
            For i = 1 To nodo.Length - 1
                If nodo(i) <> Nothing Then
                    If nodo(i).Length = n Then
                        If n = 1 Then

                            x(i, 0) = 0
                            x(i, 1) = sp.notus1(i, 0)
                        Else
                            xp = x(Array.IndexOf(nodo, nodo(i).Substring(0, nodo(i).Length - 1)), 1)
                            x(i, 0) = xp
                            x(i, 1) = xp + sp.notus1(i, 0)

                        End If

                        ed = False
                    End If
                End If
            Next
            n = n + 1

        Loop While ed = False
        Dim count As Integer = 0
        Dim idi As String = "1"
        Do
            Dim s As Integer = Array.IndexOf(nodo, idi)
            If nodo.Contains(idi & "1") = False Then
                y(s, 0) = count
                y(s, 1) = count
                y(s, 2) = count
                count = count + 1
            End If
            idi = nextnodo(nodo, idi)

        Loop While idi <> "end"

        For sz = 2 To splitsize.Max
            For i = 1 To nodo.Length - 3
                If splitsize(i) = sz Then
                    idi = nodo(i)
                    If nodo.Contains(idi & "1") = True Then
                        Dim idi1 As Integer = nodo.IndexOf(nodo, idi & "1")
                        Dim idi2 As Integer = nodo.IndexOf(nodo, idi & "2")
                        y(i, 0) = Math.Min(y(idi1, 1), y(idi2, 1))

                        y(i, 2) = Math.Max(y(idi1, 1), y(idi2, 1))
                        y(i, 1) = (y(i, 0) + y(i, 2)) / 2

                    End If
                End If
            Next
        Next
        Dim idio As String = nodo(outsplit)
        Dim idio1 As Integer = nodo.IndexOf(nodo, idio & "1")
        Dim idio2 As Integer = nodo.IndexOf(nodo, idio & "2")
        y(outsplit, 0) = Math.Min(y(idio1, 1), y(idio2, 1))

        y(outsplit, 2) = Math.Max(y(idio1, 1), y(idio2, 1))
        y(outsplit, 1) = (y(outsplit, 0) + y(outsplit, 2)) / 2
        y(nodo.GetLength(0) - 2, 0) = y(Array.IndexOf(nodo, "1"), 1)
        y(nodo.GetLength(0) - 2, 2) = y(Array.IndexOf(nodo, "2"), 1)

        Dim treemat As treemat
        treemat.nodo = nodo
        treemat.x = x
        treemat.y = y
        treemat.splitsize = splitsize
        Return treemat




    End Function
    Function checkbin(ByVal nodo() As String, ByVal nod1 As Integer, ByVal notus(,) As String)
        Dim a As Boolean
        For i = 1 To nod1 - 1

            If nodo(i) = nodo(nod1) Then
                If checkcutsite(notus(nod1, 1), notus(i, 2)) Then
                    a = True
                Else
                    a = False
                    Exit For
                End If

            End If
        Next
        Return a
    End Function
    Function correctnbt(ByVal nodo() As String, ByVal realgroups() As String, ByVal nod As String, ByVal notus() As String)
        Dim newgroup As String

        For i = 1 To nodo.Length - 2
            If nodo(i) = nod Then



                Dim ind1 As Integer = Array.IndexOf(nodo, nod.Substring(0, nod.Length - 1))
                If checkcutsite(realgroups(ind1), sptree.nOTUs(i, 2)) = True Then
                    nodo(i) = nodo(i) & "1"
                Else
                    nodo(i) = nodo(i) & "2"
                    newgroup = newgroup & realgroups(i)

                End If


            End If

        Next
        Array.Resize(realgroups, realgroups.Length)
        realgroups(realgroups.Length - 1) = newgroup
        Return nodo
    End Function
    Function nextnodo(ByVal nodo() As String, ByVal idi As String) As String

        If nodo.Contains(idi & "1") = True Then
            idi = idi & "1"
        Else
            Dim again As Boolean = False
            Do
                If idi.Substring(idi.Length - 1, 1) = "1" Then
                    If idi.Length = 1 Then
                        If nodo.Contains("21") = True Or nodo.Contains("2") = True Then
                            idi = "2"
                        Else
                            idi = "end"
                            Exit Do
                        End If

                    Else
                        idi = idi.Substring(0, idi.Length - 1) & "2"
                    End If
                    again = False
                Else
                    If idi = "2" Then
                        idi = "end"
                        Exit Do
                    Else
                        idi = idi.Substring(0, idi.Length - 1)
                        again = True
                    End If


                End If
            Loop While again = True
        End If
        Return idi

    End Function
    Function checknomax(ByVal nodoarr() As String, ByVal splitsize() As Integer, ByVal nodo As String, ByVal index As Integer) As Boolean
        Dim size As Integer
        For i = 0 To nodoarr.Length - 1
            If nodoarr(i) = nodo Then
                If splitsize(i) > size Then
                    size = splitsize(i)
                End If
            End If
        Next
        If splitsize(index) = size Then
            Return False
        Else
            Return True
        End If
    End Function
    Function realsize(ByVal splitr As String, ByVal size As Integer, ByVal repR As String) As String
        Dim spaces As Integer = countspaces(splitr, " ") + 1
        If splitr.First <> " " Then splitr = " " & splitr
        If splitr.Contains(repR & " ") = True Then
            splitr = makeroot(splitr, spaces + size, False)

        End If
        Return splitr
    End Function
    Function represent(ByVal raiz As String) As String
        Dim fin As Integer
        For i = 1 To raiz.Length - 1
            If raiz.Chars(i) = " " Then
                fin = i
                Exit For
            End If
        Next
        Dim rep As String
        rep = raiz.Substring(1, fin - 1)
        Return rep
    End Function
    Function makeroot(ByVal splitr As String, ByVal n As Integer, ByVal std As Boolean) As String
        Dim split As String
        If std = True Then
            split = " " & splitr
        Else
            For i = 2 To n
                If splitr.Contains(" " & i & " ") = False Then
                    split = split & " " & i
                End If
            Next
            split = split & " "
        End If
        Return split
    End Function

    Function checkcutsite(ByVal splitr As String, ByVal split1 As String) As Boolean


        Dim spaces(countspaces(splitr, " ") + 1) As Integer
        split1 = " " & split1
        Dim index As Integer = 0
        For i = 0 To splitr.Length - 1
            If splitr.Substring(i, 1) = " " Then
                spaces(index) = i
                index = index + 1
            End If

        Next
        Dim contain As Boolean = False
        Dim inicio As Boolean = False
        Dim cut As Boolean = False
        If split1.Contains(" " & splitr.Substring(spaces(0) + 1, spaces(1) - spaces(0) - 1) & " ") = True Then inicio = True Else inicio = False
        For j = 0 To spaces.Length - 2
            If split1.Contains(" " & splitr.Substring(spaces(j) + 1, spaces(j + 1) - spaces(j) - 1) & " ") Then contain = True Else contain = False
            Dim b As String = splitr.Substring(spaces(j) + 1, spaces(j + 1) - spaces(j) - 1)
            If inicio = True Then
                If contain = False Then
                    'Dim a As String = splitr.Substring(spaces(j) + 1, spaces(j + 1) - spaces(j) - 1)
                    cut = True
                    Exit For
                End If
            Else
                If contain = True Then
                    cut = True
                    Exit For
                End If
            End If

        Next

        Return cut

    End Function


    Function makeSbMat(ByVal sp As splits, ByVal outsplit As Integer) As treemat
        Dim nodo(sp.nOTUs.GetLength(0) + 1) As String
        Dim raiz As String
        For i = 0 To sp.nOTUs.GetLength(0) - 1
            Dim j As Integer = sp.nOTUs.GetLength(0) - 1 - i
            Dim nodoj As String
            nodoj = nodo(j)
            Dim Term As Boolean = True


            For m = 1 To j - 1
                If nodo(m) = nodoj Then
                    Term = False
                    Exit For
                End If
            Next


            If Term = False Then
                If nodo(j) <> Nothing Then
                    If nodoj.Substring(nodoj.Length - 1, 1) = "2" Then
                        If checkbin(nodo, j, sp.nOTUs) = False Then
                            nodo(j) = nodo(j) & "1"
                        Else
                            Term = True
                        End If
                    Else
                        nodo(j) = nodo(j) & "1"
                    End If
                Else
                    nodo(j) = nodo(j) & "1"
                End If
            End If

            If Term = False Then
                For k = 1 To j - 1

                    If nodo(k) = nodoj And sp.nOTUs(j, 1) <> Nothing Then
                        If checkcutsite(sp.nOTUs(j, 1), sp.nOTUs(k, 2)) = True Then
                            nodo(k) = nodo(k) & "1"

                        Else
                            nodo(k) = nodo(k) & "2"
                        End If
                    End If
                Next


            End If
        Next






        Dim x(nodo.Length - 1, 1) As Double
        Dim y(nodo.Length - 1, 2) As Double
        Dim ed As Boolean = False
        Dim n As Integer = 1

        Do
            Dim xp As Double
            ed = True
            For i = 1 To nodo.Length - 1
                If nodo(i) <> Nothing Then
                    If nodo(i).Length = n Then
                        If n = 1 Then

                            x(i, 0) = 0
                            x(i, 1) = sp.notus1(i, 0)
                        Else
                            Dim nd As String
                            Dim t As Integer
                            Dim s As Integer = 1
                            Do

                                nd = nodo(i).Substring(0, nodo(i).Length - s)
                                t = Array.IndexOf(nodo, nd)
                                s = s + 1
                            Loop While t = -1 And s < nodo(i).Length
                            If t = -1 Then t = 0
                            xp = x(t, 1)
                            x(i, 0) = xp
                            x(i, 1) = xp + sp.notus1(i, 0)

                        End If

                        ed = False
                    End If
                End If
            Next
            n = n + 1

        Loop While ed = False
        Dim count As Integer = 0
        Dim idi As String = "1"
        Dim terminales(nodo.Length - 1) As Boolean
        Do
            Dim s As Integer = Array.IndexOf(nodo, idi)
            If nodo.Contains(idi & "1") = False Then
                y(s, 0) = count
                y(s, 1) = count
                y(s, 2) = count
                count = count + 1
                terminales(s) = True
            End If
            idi = nextnodo(nodo, idi)

        Loop While idi <> "end"

        For i = 1 To nodo.Length - 2
            
            idi = nodo(i)
            If nodo.Contains(idi & "1") = True Then
                Dim idi1 As Integer = nodo.IndexOf(nodo, idi & "1")
                Dim nd As String = nodo(i) & "2"

                Dim s As Double = 0
                If nodo.Contains(idi & "2") Then
                    s = y(Array.IndexOf(nodo, idi & "2"), 1)
                Else

                    Do While nodo.Contains(nd & "1") = True
                        If nodo.Contains(nd & "2") = False Then
                            nd = nd & "2"
                        Else
                            s = y(Array.IndexOf(nodo, nd & "2"), 1)
                            Exit Do
                        End If

                    Loop


                End If
                y(i, 0) = Math.Min(y(idi1, 1), s)

                y(i, 2) = Math.Max(y(idi1, 1), s)
                y(i, 1) = (y(i, 0) + y(i, 2)) / 2
                'If y(i, 2) = 24 Then Stop
            End If

        Next

        'Dim idio As String = nodo(outsplit)



        Dim treemat As treemat
        treemat.nodo = nodo
        treemat.x = x
        treemat.y = y
        'treemat.splitsize = splitsize
        Return treemat












    End Function

    Function findsplitsize(ByVal split As String, ByVal n As Integer, ByVal repR As String) As Integer
        Dim count As Integer
        For i = 0 To split.Length - 2
            If split.Substring(i, 1) = " " Then
                count = count + 1
            End If
        Next

        If split.Contains(repR & " ") = True Then
            Return n - count
        Else
            Return count
        End If
    End Function
    Function findmax(ByVal splitsize() As Integer, ByVal limit As Integer, ByVal idinic As Integer) As Integer
        Dim max As Integer = 0
        For i = idinic To splitsize.Length - 1
            If splitsize(i) > splitsize(max) Then
                If splitsize(i) <= limit Then
                    max = i
                End If
            End If
        Next
        Return max
    End Function
    Function countspaces(ByVal split As String, ByVal a As Char) As Integer
        Dim count As Integer
        For i = 1 To split.Length - 3
            If split.Substring(i, 1) = a Then
                count = count + 1
            End If
        Next
        Return count
    End Function
    Function countchar(ByVal split As String, ByVal a As Char) As Integer
        If split = Nothing Then
            Return -1
            Exit Function
        End If
        Dim count As Integer
        For i = 0 To split.Length - 1
            If split.Substring(i, 1) = a Then
                count = count + 1
            End If
        Next
        Return count
    End Function
    Private Function rerooting(ByVal mat As treemat, ByVal outnode As Integer) As splits
        Dim out As String = mat.nodo(outnode)
        Dim idof2 As Integer = Array.IndexOf(mat.nodo, "2")
        If idof2 <> -1 Then
            sptree.notus1(idof2, 0) = CSng(sptree.notus1(sptree.notus1.GetLength(0) - 1, 0)) + CSng(sptree.notus1(Array.IndexOf(mat.nodo, "2"), 0))
        End If
        Dim nodesQ(nodes.Length - 1) As nodeap

        sptree.nOTUs(sptree.nOTUs.GetLength(0) - 1, 1) = revertsplit(sptree.nOTUs(outnode, 1), sptree.otumat1.GetLength(0) - 1)
        sptree.nOTUs(sptree.nOTUs.GetLength(0) - 1, 2) = sptree.nOTUs(outnode, 2)
        sptree.notus1(sptree.nOTUs.GetLength(0) - 1, 0) = sptree.notus1(outnode, 0) / 2
        sptree.notus1(sptree.nOTUs.GetLength(0) - 1, 1) = sptree.notus1(outnode, 1)

        sptree.notus1(outnode, 0) = sptree.notus1(outnode, 0) / 2
        Do While out.Length > 1

            out = out.Substring(0, out.Length - 1)
            Dim index As Integer = Array.IndexOf(mat.nodo, out)
            If index <> -1 And index <> sptree.nOTUs.GetLength(0) - 1 Then
                sptree.nOTUs(index, 1) = revertsplit(sptree.nOTUs(index, 1), sptree.otumat1.GetLength(0) - 1)

            End If




        Loop

        Dim notus(sptree.nOTUs.GetLength(0) - 1, sptree.nOTUs.GetLength(1) - 1) As String
        Dim notus1(sptree.nOTUs.GetLength(0) - 1, sptree.nOTUs.GetLength(1) - 1) As String


        Dim sz As Integer = 1
        Dim id As Integer = 1

        Do While id <= notus.GetLength(0) - 1

            For i = 1 To sptree.nOTUs.GetLength(0) - 1
                If splitsize(sptree.nOTUs(i, 1)) = sz Then
                    notus(id, 1) = sptree.nOTUs(i, 1)
                    notus(id, 2) = sptree.nOTUs(i, 2)
                    notus1(id, 0) = sptree.notus1(i, 0)
                    notus1(id, 1) = sptree.notus1(i, 1)
                    nodesQ(id) = nodeclone(nodes(i))
                    id = id + 1

                End If
            Next
            sz = sz + 1
        Loop
        nodesQ(nodesQ.Length - 3) = nodeclone(nodes(outnode))
        nodesQ(nodesQ.Length - 1) = nodeclone(nodes(nodesQ.Length - 1))
        nodesQ(nodesQ.Length - 2) = nodeclone(nodes(nodesQ.Length - 2))
        Dim sptree1 As splits
        sptree1.nOTUs = notus
        sptree1.notus1 = notus1
        sptree1.otumat1 = sptree.otumat1
        For i = 0 To nodes.Length - 1
            nodes(i) = nodesQ(i)
        Next
        Return sptree1
    End Function
    Function nodeclone(ByVal nodei As nodeap) As nodeap
        Dim a As nodeap
        a.bcolor = nodei.bcolor
        a.bvisible = nodei.bvisible
        a.fontb = nodei.fontb
        a.fontg = nodei.fontg
        a.gpen = New Pen(nodei.gpen.Color, nodei.gpen.Width)

        a.pen = New Pen(nodei.pen.Color, nodei.pen.Width)

        a.xb = nodei.xb
        a.yb = nodei.yb
        Return a
    End Function
    Function revertsplit(ByVal split As String, ByVal n As Integer) As String
        Dim nsplit As String = " "
        For i = 1 To n
            If split.Contains(" " & i & " ") = False Then
                nsplit = nsplit & i & " "
            End If
        Next
        Return nsplit
    End Function
    Function splitsize(ByVal split1 As String) As Integer
        Dim spaces As Integer = countspaces(split1, " ") + 1


        Return spaces
    End Function



    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged

    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Spacegroups = Spacegroups - 2
        draw()

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Spacegroups = Spacegroups + 2
        draw()

    End Sub
    Private selgroupi As Integer
    Private splitestindex As Integer
    Private selstrain As Integer



    Private Sub ReRootToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub IToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub



    Private Sub UnselectAsGroupToTestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        drawnodenames = CheckBox3.Checked

        draw()

    End Sub

    Private Sub SToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub SelectGroupBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub HideSupportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseWheel
        Dim i As Integer = e.Delta
        If i > 0 Then
            scalexy = scalexy + 10
            scalexy1 = scalexy1 + 10
            If _GRUBURST.gs IsNot Nothing Then selected = 0

            If selected = 0 Then
                If My.Computer.Keyboard.CtrlKeyDown = False Then
                    y_factor1 = y_factor1 - 0.05
                    pixFromTop = pixFromTop - 1
                Else
                    maxWidth = maxWidth + 0.05
                End If
            ElseIf selected = 1 Then
                nodes(selgroupi).gpen.Width = nodes(selgroupi).gpen.Width + 1
            ElseIf selected = 2 Then
                For j = 1 To treematrix.nodo.GetLength(0) - 3
                    If treematrix.nodo(j).StartsWith(treematrix.nodo(selgroupi)) Then
                        nodes(j).pen.Width = nodes(j).pen.Width + 1
                    End If

                Next

            End If

        Else
            scalexy = scalexy - 10
            scalexy1 = scalexy1 - 10

            If selected = 0 Then
                If My.Computer.Keyboard.CtrlKeyDown = False Then
                    y_factor1 = y_factor1 + 0.05
                    pixFromTop = pixFromTop + 1
                Else
                    maxWidth = maxWidth - 0.05
                End If
                'maxWidth = maxWidth + 0.05
            ElseIf selected = 1 Then
                nodes(selgroupi).gpen.Width = nodes(selgroupi).gpen.Width - 0.2
            ElseIf selected = 2 Then
                For i = 1 To treematrix.nodo.GetLength(0) - 3
                    If treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
                        nodes(i).pen.Width = nodes(i).pen.Width - 0.2
                    End If

                Next

            End If
        End If
        draw()
    End Sub

    Private Sub PictureBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles PictureBox1.PreviewKeyDown
        Select Case e.KeyCode
            Case Keys.Right
                moveR()
            Case Keys.Left
                moveL()
            Case Keys.Up
                moveU()
            Case Keys.Down
                moveD()
            Case Keys.Control

        End Select







    End Sub
    Sub moveR()
        If selected = 1 Then
            isplitests(splitestindex) = isplitests(splitestindex) + 4
        ElseIf selected = 0 Then
            pixFromLeft = pixFromLeft + 10
        ElseIf selected = 2 Then

            nodes(selgroupi).xb = nodes(selgroupi).xb + 4
        End If
        draw()
    End Sub
    Sub moveL()
        If selected = 1 Then
            isplitests(splitestindex) = isplitests(splitestindex) - 4
        ElseIf selected = 0 Then
            pixFromLeft = pixFromLeft - 10
        ElseIf selected = 2 Then

            nodes(selgroupi).xb = nodes(selgroupi).xb - 4
        End If
        draw()
    End Sub
    Sub moveU()
        If selected = 2 Then

            nodes(selgroupi).yb = nodes(selgroupi).yb - 4
        ElseIf selected = 0 Then
            pixFromTop = pixFromTop - 10
        End If
        draw()
    End Sub
    Sub moveD()
        If selected = 2 Then

            nodes(selgroupi).yb = nodes(selgroupi).yb + 4
        ElseIf selected = 0 Then
            pixFromTop = pixFromTop + 10
        End If
        draw()
    End Sub




    Private Sub BranchColorToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub ContextMenuStrip2_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        EditBranchSupportToolStripMenuItem.Visible = False
        FlipToolStripMenuItem.Visible = False
        ReRootToolStripMenuItem1.Visible = False
        SupportToolStripMenuItem.Visible = False
        BranchToolStripMenuItem.Visible = False
        StrainsToolStripMenuItem.Visible = False
        GroupToolStripMenuItem.Visible = False
        SelectAsGroupToTestToolStripMenuItem.Visible = False
        UnselectAsGroupToTestToolStripMenuItem1.Visible = False
        SelectBarForPositionToolStripMenuItem.Visible = False
        SelectBarForSizeToolStripMenuItem.Visible = False
        ColorToolStripMenuItem2.Visible = False
        FontToolStripMenuItem2.Visible = False
        CopToolStripMenuItem.Visible = False
        PasteStyleToolStripMenuItem.Visible = False

        Dim ss As Point
        ss = PictureBox1.PointToClient(Cursor.Position)
        selected = 0
        selgroupi = 0
        minx = (ss.X + 2 - pixFromLeft) / x_factor
        miny = (ss.Y - 10 - pixFromTop) / y_factor
        maxx = (ss.X - 2 - pixFromLeft) / x_factor
        maxy = (ss.Y + 10 - pixFromTop) / y_factor

        With treematrix
            For i = 1 To treematrix.x.GetLength(0) - 4
                'If i = 9 Then Stop
                If .x(i, 0) <= minx And .x(i, 1) > maxx And .y(i, 1) >= miny And .y(i, 1) < maxy Then
                    ContextMenuStrip2.Visible = True
                    selgroupi = i
                    ToolStripTextBox1.Text = _sptree.notus1(selgroupi, 1)
                    splitestindex = 0
                    Dim splitests(,) As String = Module1.splitestmaker
                    For j = 1 To splitests.GetLength(0) - 2
                        If splitests(j, 1) = sptree.nOTUs(i, 2) Then

                            'SelectAsGroupToTestToolStripMenuItem.Visible = False
                            'UnselectAsGroupToTestToolStripMenuItem1.Visible = True

                            splitestindex = j
                            Exit For

                        End If

                    Next
                    If splitestindex <> 0 Then
                        UnselectAsGroupToTestToolStripMenuItem1.Visible = True
                        SelectBarForPositionToolStripMenuItem.Visible = True
                        SelectBarForSizeToolStripMenuItem.Visible = True
                        ColorToolStripMenuItem2.Visible = True
                        FontToolStripMenuItem2.Visible = True
                    Else
                        SelectAsGroupToTestToolStripMenuItem.Visible = True

                    End If
                    ReRootToolStripMenuItem1.Visible = True
                    SupportToolStripMenuItem.Visible = True
                    BranchToolStripMenuItem.Visible = True
                    CopToolStripMenuItem.Visible = True
                    FlipToolStripMenuItem.Visible = True
                    EditBranchSupportToolStripMenuItem.Visible = True
                    If copyN.pen IsNot Nothing Then
                        PasteStyleToolStripMenuItem.Visible = True
                    End If
                    GroupToolStripMenuItem.Visible = True
                    If _sptree.otumat1(1, 1) <> Nothing And UPGMA = False Then
                        CalculateACSToolStripMenuItem.Enabled = True
                    Else
                        CalculateACSToolStripMenuItem.Enabled = False
                    End If
                    Exit For
                ElseIf .x(i, 0) <= minx And .y(i, 1) >= miny And .y(i, 1) < maxy Then
                    Try
                        selstrain = _sptree.nOTUs(i, 1)
                        StrainsToolStripMenuItem.Visible = True
                        CopToolStripMenuItem.Visible = True
                        If copyT.tfont IsNot Nothing Then
                            PasteStyleToolStripMenuItem.Visible = True
                        End If
                        Exit For
                    Catch
                    End Try
                End If
            Next i
        End With

    End Sub

    Private Sub ReRootToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReRootToolStripMenuItem1.Click
        sptree = rerooting(treematrix, selgroupi)


        treematrix = makeSbMat(sptree, selgroupi)
        draw()
    End Sub

    Private Sub WriteGroupNameAndClickHereToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WriteGroupNameAndClickHereToolStripMenuItem.Click

        Dim split As String

        Form1.DataGridView6.Rows.Add()
        Form1.DataGridView6.Item(0, Form1.DataGridView6.RowCount - 1).Value = ToolStripTextBox2.Text
        split = sptree.nOTUs(selgroupi, 1)

        Dim ax() As String = sptree.nOTUs(selgroupi, 1).Split(" "c)
        Dim strains As String
        For i = 1 To ax.Length - 2



            strains = strains & sptree.otumat1(ax(i), 0) & ", "

        Next
        Dim rows As Integer = Form1.DataGridView6.RowCount - 1
        Form1.DataGridView6.Item(3, rows).Value = split
        Form1.DataGridView6.Item(2, rows).Value = sptree.nOTUs(selgroupi, 2)
        Form1.DataGridView6.Item(1, rows).Value = strains

        draw()

    End Sub

    Private Sub UnselectAsGroupToTestToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnselectAsGroupToTestToolStripMenuItem1.Click
        Dim splitests(,) As String = Module1.splitestmaker
        Dim cellpos As Integer
        For j = 1 To splitests.GetLength(0) - 2
            If splitests(j, 1) = sptree.nOTUs(selgroupi, 2) Then
                cellpos = j - 1

                Exit For

            End If

        Next

        Form1.DataGridView6.Rows.RemoveAt(cellpos)
        draw()

    End Sub

    Private Sub SelectForSizeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectForSizeToolStripMenuItem.Click

        selected = 2
    End Sub

    Private Sub SelectForPositionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectForPositionToolStripMenuItem.Click

        selected = 2
    End Sub

    Private Sub LineColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        nodes(selgroupi).pen.Color = ColorDialog1.Color
        'For i = 1 To treematrix.nodo.GetLength(0) - 3
        'If treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
        'nodes(i).pen.Color = ColorDialog1.Color
        ' End If

        'Next
        draw()
    End Sub

    Private Sub ShowhideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowhideToolStripMenuItem.Click

        If nodes(selgroupi).bvisible = True Then
            nodes(selgroupi).bvisible = False
        Else
            nodes(selgroupi).bvisible = True
        End If
        draw()
    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        FontDialog1.Font = nodes(selgroupi).fontb
        FontDialog1.ShowDialog()
        nodes(selgroupi).fontb = FontDialog1.Font
        draw()
    End Sub

    Private Sub ColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        nodes(selgroupi).bcolor = New SolidBrush(ColorDialog1.Color)
        draw()
    End Sub

    Private Sub SelectBarForPositionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectBarForPositionToolStripMenuItem.Click
        selected = 1
    End Sub

    Private Sub FontToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem1.Click
        FontDialog1.Font = taxa(selstrain).tfont
        FontDialog1.ShowDialog()

        taxa(selstrain).tfont = FontDialog1.Font
        draw()

    End Sub

    Private Sub ColorToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColorToolStripMenuItem1.Click



        ColorDialog1.ShowDialog()
        taxa(selstrain).tcolor = New SolidBrush(ColorDialog1.Color)
        draw()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        FontDialog1.ShowDialog()
        If taxa IsNot Nothing Then
            For i = 1 To _sptree.otumat1.GetLength(0) - 1
                taxa(i).tfont = FontDialog1.Font
            Next
        End If
        draw()
    End Sub

    Private Sub ShowhideTaxonToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowhideTaxonToolStripMenuItem.Click
        If taxa(selstrain).tvisible = False Then
            taxa(selstrain).tvisible = True
        Else
            taxa(selstrain).tvisible = False
        End If
        draw()
    End Sub


    Private Sub SelectBarForSizeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectBarForSizeToolStripMenuItem.Click
        selected = 1
    End Sub

    Private Sub FontToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem2.Click
        FontDialog1.Font = nodes(selgroupi).fontg
        FontDialog1.ShowDialog()
        nodes(selgroupi).fontg = FontDialog1.Font
        draw()
    End Sub

    Private Sub ColorToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColorToolStripMenuItem2.Click
        ColorDialog1.Color = nodes(selgroupi).gpen.Color
        ColorDialog1.ShowDialog()
        nodes(selgroupi).gpen.Color = ColorDialog1.Color
        'For i = 1 To treematrix.nodo.GetLength(0) - 3
        'If treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
        'nodes(i).pen.Color = ColorDialog1.Color
        ' End If

        'Next
        draw()
    End Sub

    Private Sub CopToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopToolStripMenuItem.Click
        If selgroupi <> 0 Then
            copyN = nodes(selgroupi)
        ElseIf selstrain <> 0 Then
            copyT = taxa(selstrain)

        End If
    End Sub

    Private Sub PasteStyleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteStyleToolStripMenuItem.Click
        If selgroupi <> 0 Then

            nodes(selgroupi).gpen = New Pen(copyN.gpen.Color, copyN.gpen.Width)

            nodes(selgroupi).bcolor = copyN.bcolor
            nodes(selgroupi).bvisible = copyN.bvisible
            nodes(selgroupi).fontb = copyN.fontb
            nodes(selgroupi).fontg = copyN.fontg
            For i = 1 To treematrix.nodo.GetLength(0) - 3
                If treematrix.nodo(i) = (treematrix.nodo(selgroupi)) Then 'treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
                    nodes(i).pen = New Pen(copyN.pen.Color, copyN.pen.Width)
                End If

            Next

        ElseIf selstrain <> 0 Then
            taxa(selstrain) = copyT

        End If
        draw()
    End Sub

    Private Sub CalculateACSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculateACSToolStripMenuItem.Click
        If selgroupi <> 0 Then
            If _sptree.otumat1(1, 1) <> Nothing Then
                If supportnames Is Nothing Then
                    ReDim supportnames(1)
                End If

                CheckBox2.Checked = True
                If supportnames(0) Is Nothing Then
                    ReDim supportnames(1)
                    supportnames(0) = "1-njCS"
                    ReDim showit(1)
                    showit(0) = True
                    ReDim support(1)
                    support(0) = -1
                    ComboBox1.Visible = True
                    ComboBox1.Items.Add("1-njCS")
                    ComboBox1.SelectedIndex = 0
                    Label8.Visible = True
                    TextBox1.Visible = True
                    CheckBox5.Visible = True
                    _sptree.notus1(selgroupi, 1) = Math.Round(cladesupport(_sptree.nOTUs(selgroupi, 2), _sptree, Module1.NJgo(_sptree.otumat1, False, False), _sptree.otumat1) * 1000) / 1000
                Else
                    If supportnames.Contains("1-njCS") Then
                        Dim index As Integer = Array.IndexOf(supportnames, "1-njCS")
                        Dim vals() As String
                        If sptree.notus1(selgroupi, 1) <> Nothing Then
                            vals = sptree.notus1(selgroupi, 1).Split("/")
                        Else
                            ReDim vals(0)

                        End If

                        If index > vals.Length - 1 Then
                            Array.Resize(vals, vals.Length + 1)

                        End If
                        vals(index) = Math.Round(cladesupport(_sptree.nOTUs(selgroupi, 2), _sptree, Module1.NJgo(_sptree.otumat1, False, False), _sptree.otumat1) * 1000) / 1000

                        Dim value As String = Nothing

                        For i = 0 To vals.Length - 1
                            If value = Nothing Then
                                value = vals(i)
                            Else
                                value = value & "/" & vals(i)
                            End If
                        Next
                        _sptree.notus1(selgroupi, 1) = value
                    Else
                        _sptree.notus1(selgroupi, 1) = _sptree.notus1(selgroupi, 1) & "/" & Math.Round(cladesupport(_sptree.nOTUs(selgroupi, 2), _sptree, Module1.NJgo(_sptree.otumat1, False, False), _sptree.otumat1) * 1000) / 1000
                        Array.Resize(supportnames, supportnames.Length + 1)
                        Array.Resize(support, support.Length + 1)
                        Array.Resize(showit, showit.Length + 1)
                        supportnames(supportnames.Length - 1) = "1-njCS"
                        showit(supportnames.Length - 1) = True
                        support(supportnames.Length - 1) = -1
                        ComboBox1.Items.Add("1-njCS")
                        ComboBox1.Visible = True
                    End If
                End If
            End If
        End If
        draw()
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        If nodes IsNot Nothing Then
            For i = 0 To nodes.Length - 1
                nodes(i).bvisible = CheckBox2.Checked
            Next
        End If

        drawcirclegroup = CheckBox2.Checked

        If Me.Visible = True Then
            draw()
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Panel3.Visible = True
        ComboBox2.Items.Clear()
        For i = 0 To supportnames.Length - 1
            ComboBox2.Items.Add(supportnames(i))
            ComboBox2.SelectedIndex = 0
        Next
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        coloring()
        Panel3.Visible = False
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click

        Select Case sender.text
            Case "<="
                sender.text = "="
            Case "="
                sender.text = ">="
            Case ">="
                sender.text = "<="
        End Select
        coloring()

    End Sub
    Sub coloring()
        Dim count1, count2, count3, count4, count5, count0, countt As Integer


        color1 = Button16.BackColor
        color2 = Button18.BackColor
        color3 = Button19.BackColor
        color4 = Button20.BackColor
        color5 = Button21.BackColor
        color6 = Button22.BackColor
        Dim i As Integer = 0
        For z = 1 To nodes.Length - 4
            i = nodes.Length - z - 3

            If splitsize(_sptree.nOTUs(i, 1)) > 1 And _sptree.notus1(i, 1) <> Nothing And _sptree.notus1(i, 0) > 10 ^ -9 Then
                countt = countt + 1
                Dim vals() As String = Nothing

                vals = sptree.notus1(i, 1).Split("/")
                Dim a As Single = vals(ComboBox2.SelectedIndex)

                Select Case a
                    Case Is = compareexp(Button11, a, TextBox2.Text)
                        nodes(i).pen.Color = color1
                        If i <> nodes.Length - 3 Then count0 = count0 + 1
                    Case Is = compareexp(Button12, a, TextBox3.Text)
                        nodes(i).pen.Color = color2
                        If i <> nodes.Length - 3 Then count1 = count1 + 1
                    Case Is = compareexp(Button13, a, TextBox4.Text)
                        nodes(i).pen.Color = color3
                        If i <> nodes.Length - 3 Then count2 = count2 + 1
                    Case Is = compareexp(Button14, a, TextBox5.Text)


                        nodes(i).pen.Color = color4
                        If i <> nodes.Length - 3 Then count3 = count3 + 1
                    Case Is = compareexp(Button15, a, TextBox6.Text)
                        nodes(i).pen.Color = color5
                        If i <> nodes.Length - 3 Then count4 = count4 + 1
                    Case Else
                        nodes(i).pen.Color = Color.Black
                End Select
            ElseIf splitsize(_sptree.nOTUs(i, 1)) = 1 Then

                nodes(i).pen.Color = color6
            ElseIf _sptree.notus1(i, 0) < 10 ^ -9 Then
                'Stop
                For h = i + 1 To treematrix.nodo.Length - 1
                    If treematrix.nodo(h) = treematrix.nodo(i).Substring(0, treematrix.nodo(i).Length - 1) Then
                        'Stop

                        nodes(i).pen.Color = nodes(h).pen.Color

                    End If
                Next
            End If

        Next
        'countt = countt - 1
        Label11.Text = Math.Round(count0 / countt * 1000) / 10 & "%"
        Label12.Text = Math.Round(count1 / countt * 1000) / 10 & "%"
        Label13.Text = Math.Round(count2 / countt * 1000) / 10 & "%"
        Label14.Text = Math.Round(count3 / countt * 1000) / 10 & "%"
        Label15.Text = Math.Round(count4 / countt * 1000) / 10 & "%"

        draw()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click, Button18.Click, Button19.Click, Button20.Click, Button21.Click, Button22.Click
        ColorDialog1.ShowDialog()
        sender.BackColor = ColorDialog1.Color



    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.Visible = True Then
            viewlabels = CheckBox1.Checked
            draw()
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, TextBox6.TextChanged
        If IsNumeric(TextBox2.Text) And IsNumeric(TextBox3.Text) And IsNumeric(TextBox4.Text) And IsNumeric(TextBox5.Text) And IsNumeric(TextBox6.Text) Then
            If nodes IsNot Nothing Then
                coloring()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.Visible Then
            TextBox1.Text = support(ComboBox1.SelectedIndex)
            CheckBox5.Checked = showit(ComboBox1.SelectedIndex)
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        coloring()
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If Me.Visible = True Then
            showit(ComboBox1.SelectedIndex) = CheckBox5.CheckState
            draw()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Select Case Timer1.Tag
            Case "23"
                RepeatButton4_Click()
            Case "24"
                RepeatButton3_Click()
            Case "25"
                RepeatButton5_Click()
            Case "26"
                RepeatButton6_Click()
            Case "27"
                RepeatButton1_Click()
            Case "28"
                RepeatButton2_Click()
            Case "29"
                RepeatButton7_Click()
            Case "30"
                RepeatButton8_Click()

        End Select


    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click, Button30.Click, Button29.Click, Button28.Click, Button27.Click

    End Sub

    Private Sub Button23_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button23.MouseDown, Button24.MouseDown, Button25.MouseDown, Button26.MouseDown, Button30.MouseDown, Button29.MouseDown, Button28.MouseDown, Button27.MouseDown
        Timer1.Tag = sender.tag
        Timer1.Start()


    End Sub

    Private Sub Button23_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button23.MouseUp, Button24.MouseUp, Button25.MouseUp, Button26.MouseUp, Button30.MouseUp, Button29.MouseUp, Button28.MouseUp, Button27.MouseUp
        Timer1_Tick(Nothing, Nothing)
        Timer1.Stop()

    End Sub

    Private Sub BranchesColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BranchesColorToolStripMenuItem.Click
        ColorDialog1.Color = nodes(selgroupi).gpen.Color
        ColorDialog1.ShowDialog()
        nodes(selgroupi).gpen.Color = ColorDialog1.Color
        For i = 1 To treematrix.nodo.GetLength(0) - 3
            If treematrix.nodo(i).StartsWith(treematrix.nodo(selgroupi)) Then
                nodes(i).pen.Color = ColorDialog1.Color
            End If

        Next
        draw()
    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        If selgroupi > 0 Then
            _sptree.notus1(selgroupi, 1) = ToolStripTextBox1.Text
            Dim a() As String = ToolStripTextBox1.Text.Split("/"c)
            If showit Is Nothing Then
                ReDim showit(0)
                ReDim support(0)
                ComboBox1.Visible = True
                TextBox1.Visible = True
                CheckBox5.Visible = True
                showit(showit.Length - 1) = True
                ComboBox1.Items.Add("Branch value")
                support(a.Length - 1) = -1
                CheckBox2.Checked = True
                ComboBox1.SelectedIndex = 0
            End If

            If a.Length > showit.Length Then
                showit.Resize(showit, a.Length)
                showit(showit.Length - 1) = True
                ComboBox1.Items.Add("Branch value")
                Array.Resize(support, a.Length)
                support(a.Length - 1) = -1
            End If
        End If
    End Sub

    Private Sub MeanConsensusSupportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MeanConsensusSupportToolStripMenuItem.Click

    

       
    End Sub
  
    Private Sub FlipToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FlipToolStripMenuItem.Click
        If selgroupi > 0 Then
            flipbranch(selgroupi)
        End If
    End Sub
    Sub flipbranch(ByVal idbranch As Integer)
        With treematrix
            Dim unomayordos As Boolean
            Dim uno As Integer
            Dim dos As Integer
            Dim idi As String
            idi = .nodo(idbranch)
            Dim count1, count2 As Integer
            For i = 1 To _sptree.otumat1.GetLength(0) - 1
                If .nodo(i).StartsWith(idi & "1") = True Then
                    count1 = count1 + 1
                    uno = .y(i, 0)
                ElseIf .nodo(i).StartsWith(idi & "2") Then
                    count2 = count2 + 1
                    dos = .y(i, 0)
                End If

            Next
            If uno > dos Then
                count2 = -count2
            Else
                count1 = -count1
            End If
            For i = 1 To _sptree.otumat1.GetLength(0) - 1

                If .nodo(i).StartsWith(idi & "1") = True Then
                    .y(i, 0) = .y(i, 0) + count2
                    .y(i, 1) = .y(i, 1) + count2
                    .y(i, 2) = .y(i, 2) + count2
                ElseIf .nodo(i).StartsWith(idi & "2") Then
                    .y(i, 0) = .y(i, 0) + count1
                    .y(i, 1) = .y(i, 1) + count1
                    .y(i, 2) = .y(i, 2) + count1
                End If

            Next

            For i = 1 To .nodo.Length - 2

                idi = .nodo(i)
                If .nodo.Contains(idi & "1") = True Then
                    Dim idi1 As Integer = .nodo.IndexOf(.nodo, idi & "1")
                    Dim nd As String = .nodo(i) & "2"

                    Dim s As Double = 0
                    If .nodo.Contains(idi & "2") Then
                        s = .y(Array.IndexOf(.nodo, idi & "2"), 1)
                    Else

                        Do While .nodo.Contains(nd & "1") = True
                            If .nodo.Contains(nd & "2") = False Then
                                nd = nd & "2"
                            Else
                                s = .y(Array.IndexOf(.nodo, nd & "2"), 1)
                                Exit Do
                            End If

                        Loop


                    End If
                    .y(i, 0) = Math.Min(.y(idi1, 1), s)

                    .y(i, 2) = Math.Max(.y(idi1, 1), s)
                    .y(i, 1) = (.y(i, 0) + .y(i, 2)) / 2
                    'If y(i, 2) = 24 Then Stop
                End If

            Next
        End With
        draw()
    End Sub
End Class
