Imports System.Windows.Forms

Public Class props
    Private _sptree As splits
    Private _treemat As treemat
    Private _support As Single()
    Private _GRUBURST As reburst
    Private _xypoints(,) As Double
    Private _maxWidth As Single
    Private _y_factor As Single
    Private _pixFromTop As Single
    Private _pixFromLeft As Single
    Private _fontsize As Integer
    Private _viewlabels As Boolean
    Private _drawcirclegroup As Boolean
    Private _colordefuenteboo As New SolidBrush(Color.Black)
    Private colordefuente As New SolidBrush(Color.Black)
    Private iwidth As Single
    Private iHeight As Single
    Private tbwidth As Single
    Private scalexy
    Private linecolor As Color = Color.Blue
    Private spacegroups As Single
    Private isplitests() As Single
    Private nodes As nodeap()
    Dim drawnodenames As Boolean
    Private taxa() As taxon
    Private linewidth As Single
    Private SIZEST As Single
    Private group As Integer
    Private scalebarlength As Double
    Private dr
    Private showit() As Boolean


    Sub New()
        InitializeComponent()

        _GRUBURST = Nothing
        _xypoints = Nothing
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Public Property _taxa() As taxon()
        Get
            Return taxa
        End Get
        Set(ByVal value() As taxon)
            taxa = value
        End Set
    End Property
    Public Property _sbl() As Double
        Get
            Return scalebarlength
        End Get
        Set(ByVal GRAPH As Double)
            scalebarlength = GRAPH
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
    Public Property _linewidth() As Single
        Get
            Return linewidth
        End Get
        Set(ByVal value As Single)
            linewidth = value
        End Set
    End Property
    Public Property _sizeST() As Single
        Get
            Return SIZEST
        End Get
        Set(ByVal value As Single)
            SIZEST = value
        End Set
    End Property
    Public Property _spacegroups() As Single
        Get
            Return spacegroups
        End Get
        Set(ByVal value As Single)
            spacegroups = value
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
    Public Property _scalexy() As Single
        Get
            Return scalexy
        End Get
        Set(ByVal GRAPH As Single)
            scalexy = GRAPH
        End Set
    End Property
    Public Property _tbwidth() As Single
        Get
            Return tbwidth
        End Get
        Set(ByVal GRAPH As Single)
            tbwidth = GRAPH
        End Set
    End Property

    Public Property _iHeight() As Single
        Get
            Return iHeight
        End Get
        Set(ByVal GRAPH As Single)
            iHeight = GRAPH
        End Set
    End Property
    Public Property _iwidth() As Single
        Get
            Return iwidth
        End Get
        Set(ByVal GRAPH As Single)
            iwidth = GRAPH
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
    Public Property colordefuenteboo() As SolidBrush
        Get
            Return _colordefuenteboo
        End Get
        Set(ByVal GRAPH As SolidBrush)
            _colordefuenteboo = GRAPH
        End Set
    End Property
    
    Public Property viewlabels() As Boolean
        Get
            Return _viewlabels
        End Get
        Set(ByVal GRAPH As Boolean)
            _viewlabels = GRAPH
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

    Public Property fontsize() As Integer
        Get
            Return _fontsize
        End Get
        Set(ByVal GRAPH As Integer)
            _fontsize = GRAPH
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
    Public Property xypoints() As Double(,)
        Get
            Return _xypoints
        End Get
        Set(ByVal GRAPH As Double(,))
            _xypoints = GRAPH
        End Set
    End Property
    Public Property maxWidth() As Single
        Get
            Return _maxWidth
        End Get
        Set(ByVal GRAPH As Single)
            _maxWidth = GRAPH
        End Set
    End Property
    Public Property y_factor() As Single
        Get
            Return _y_factor
        End Get
        Set(ByVal GRAPH As Single)
            _y_factor = GRAPH
        End Set
    End Property

    Public Property pixFromTop() As Single
        Get
            Return _pixFromTop
        End Get
        Set(ByVal GRAPH As Single)
            _pixFromTop = GRAPH
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
    Public Property treematx() As treemat
        Get
            Return _treemat
        End Get
        Set(ByVal tree As treemat)
            _treemat = tree
        End Set
    End Property
    Public Property cutoffb() As Single()
        Get
            Return _support
        End Get
        Set(ByVal cut As Single())
           
            _support = cut
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
    Public Property pixFromLeft() As String
        Get
            Return _pixFromLeft
        End Get
        Set(ByVal cut As String)
            _pixFromLeft = cut
        End Set
    End Property
    Public Property _drawcg() As Boolean
        Get
            Return _drawcirclegroup
        End Get
        Set(ByVal GRAPH As Boolean)
            _drawcirclegroup = GRAPH
        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim treeviewer1 As New treeviewer2

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        '  Try

        treeviewer1.SaveFileDialog1.FileName = ""
        Dim resol As Integer = TextBox2.Text
        Dim maxwidthpix As Integer = TextBox1.Text / 25.4 * resol
        Dim HIGHT As Integer
        If _treemat.nodo IsNot Nothing Then
            HIGHT = CInt(Int(maxwidthpix))
        Else
            HIGHT = CInt(Int(maxwidthpix * (iHeight / iwidth)))
        End If
       
        Dim bit As Bitmap = New Bitmap(CInt(Int(maxwidthpix)), HIGHT, Imaging.PixelFormat.Format24bppRgb)


        'treeviewer1._MaxWidth = (maxWidth * (bit.Width / iwidth)) - (tbwidth * (bit.Width / iwidth))
       
        If _treemat.nodo IsNot Nothing Then


            treeviewer1.sptree = _sptree
            treeviewer1._cutoffz = _support
            treeviewer1._sbl = scalebarlength

            treeviewer1._spacegroups = spacegroups * (bit.Height / iHeight)

            treeviewer1._y_factor = y_factor * (bit.Height / iHeight)

            treeviewer1._pixFromTop = pixFromTop * (bit.Height / iHeight)
            treeviewer1._pixFromLeft = pixFromLeft * (bit.Width / iwidth)
            treeviewer1._viewlabels = viewlabels
            treeviewer1._drawnodenames = drawnodenames
            treeviewer1._showit = showit
            Dim ispclone(isplitests.Length - 1) As Single
            For i = 1 To isplitests.Length - 1
                ispclone(i) = isplitests(i) * (bit.Width / iwidth)
            Next
            Dim nodesclone(nodes.Length - 1) As nodeap
            For i = 1 To nodes.Length - 1
                nodesclone(i).xb = nodesclone(i).xb * (bit.Width / iwidth)
                nodesclone(i).yb = nodesclone(i).yb * (bit.Height / iHeight)
                nodesclone(i).fontb = New Font(nodes(i).fontb.Name, nodes(i).fontb.Size * (bit.Height / iHeight), nodes(i).fontb.Style)
                nodesclone(i).pen = New Pen(nodes(i).pen.Color, nodes(i).pen.Width * (bit.Height / iHeight))
                nodesclone(i).bvisible = nodes(i).bvisible
                nodesclone(i).bcolor = nodes(i).bcolor
                nodesclone(i).gpen = New Pen(nodes(i).gpen.Color, nodes(i).gpen.Width * (bit.Height / iHeight))
                nodesclone(i).fontg = New Font(nodes(i).fontb.Name, nodes(i).fontg.Size * (bit.Height / iHeight), nodes(i).fontg.Style)

            Next
            Dim taxaclone(taxa.Length - 1) As taxon
            For i = 1 To taxa.Length - 1
                taxaclone(i).tcolor = taxa(i).tcolor
                taxaclone(i).tfont = New Font(taxa(i).tfont.Name, taxa(i).tfont.Size * (bit.Height / iHeight), taxa(i).tfont.Style)
                taxaclone(i).tvisible = taxa(i).tvisible
            Next
            treeviewer1._nodes = nodesclone
            treeviewer1._isplitests = ispclone
            treeviewer1._taxa = taxaclone

            bit = treeviewer1.drawTree(bit, _treemat)

        Else
            treeviewer1._scalexy = scalexy * (bit.Width / iwidth)
            treeviewer1._fontsize = fontsize * (bit.Height / iHeight)
            treeviewer1._linewidth = linewidth * (bit.Height / iHeight)
            treeviewer1._SIZEST = SIZEST * (bit.Width / iwidth)

            If GRUBURST.gs IsNot Nothing Then
                treeviewer1.GRUBURST = _GRUBURST
                Dim nodesclone(nodes.Length - 1) As nodeap
                For i = 1 To nodes.Length - 1
                    nodesclone(i).xb = nodes(i).xb * (bit.Width / iwidth)
                    nodesclone(i).yb = nodes(i).yb * (bit.Height / iHeight)
                   

                Next
                treeviewer1._nodes = nodesclone
                treeviewer1._drawcg = _drawcirclegroup
                If group > _GRUBURST.gs.Length - 2 Or group < 0 Then

                    group = _GRUBURST.gs.Length - 2


                End If

                bit = treeviewer1.draweburst(bit, group)
            Else
                treeviewer1.xypoints = _xypoints
                treeviewer1._pixFromTop = pixFromTop * (bit.Height / iHeight)
                treeviewer1._pixFromLeft = pixFromLeft * (bit.Width / iwidth)
                treeviewer1._viewlabels = _viewlabels

                bit = treeviewer1.drawmds(bit)
            End If
        End If
        Dim a As String = "Format " & ComboBox1.Text & "|" & "*." & ext1(ComboBox1.Text)
        treeviewer1.SaveFileDialog1.Filter = a
        treeviewer1.SaveFileDialog1.DefaultExt = ext1(ComboBox1.Text)

        treeviewer1.SaveFileDialog1.ShowDialog()
        If treeviewer1.SaveFileDialog1.FileName <> "" Then
            bit.SetResolution(resol, resol)

            bit.Save(treeviewer1.SaveFileDialog1.FileName, ext(ComboBox1.Text))
        End If
       
       

        bit.Dispose()
        Me.Close()
        '  Catch
        ' End Try
    End Sub
    Function ext(ByVal a As String) As Imaging.ImageFormat
        Dim ex As Imaging.ImageFormat
        Select Case a
            Case "Jpeg"
                ex = Imaging.ImageFormat.Jpeg
            Case "Tiff"
                ex = Imaging.ImageFormat.Tiff
            Case "Gif"
                ex = Imaging.ImageFormat.Gif
            Case "BMP"
                ex = Imaging.ImageFormat.Bmp
            Case "PNG"
                ex = Imaging.ImageFormat.Png
        End Select
        Return ex
    End Function
    Function ext1(ByVal a As String) As String
        Dim ex As String
        Select Case a
            Case "Jpeg"
                ex = "jpg"
            Case "Tiff"
                ex = "tif"
            Case "Gif"
                ex = "gif"
            Case "BMP"
                ex = "bmp"
            Case "PNG"
                ex = "png"
        End Select
        Return ex
    End Function
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    
 

 

  
End Class
