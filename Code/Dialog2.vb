Imports System.Windows.Forms

Public Class Dialog2
    Private _gdef As Integer
    Private _page As Integer
    Private _nloc As Integer
    Private _nrep As Integer
    Private cutoff As Integer
    Private viewpanel1, viewboots As Boolean
    Private savetrees As Boolean
    Private supp As Boolean
    Private hethand As Integer
    Private bionj As Boolean
    Private viewsupport As Boolean
    Private templeton As Boolean


    Sub New()

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        _gdef = Nothing
        _page = Nothing
        _nloc = Nothing
        _nrep = 0
        viewboots = True
        supp = False
        viewpanel1 = False
        savetrees = False
        bionj = False

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Public Property temple() As Boolean
        Get
            Return templeton
        End Get
        Set(ByVal value As Boolean)
            templeton = value
        End Set
    End Property
    Public Property _bionj() As Boolean
        Get
            Return bionj
        End Get
        Set(ByVal a As Boolean)
            bionj = a
        End Set
    End Property
    Public Property _hethand() As Integer
        Get
            Return hethand
        End Get
        Set(ByVal a As Integer)
            hethand = a
        End Set
    End Property
    Public Property _supp() As Boolean
        Get
            Return supp
        End Get
        Set(ByVal a As Boolean)
            supp = a
        End Set
    End Property
    Public Property _cutoff() As Integer
        Get
            Return cutoff
        End Get
        Set(ByVal a As Integer)
            cutoff = a
        End Set
    End Property

    Public Property _savetrees() As Boolean
        Get
            Return savetrees
        End Get
        Set(ByVal a As Boolean)
            savetrees = a
        End Set
    End Property
    Public Property viewnmin() As Boolean
        Get
            Return viewpanel1
        End Get
        Set(ByVal a As Boolean)
            viewpanel1 = a
        End Set
    End Property
    Public Property viewboo() As Boolean
        Get
            Return viewboots
        End Get
        Set(ByVal a As Boolean)
            viewboots = a
        End Set
    End Property
    Public Property gdef() As Integer
        Get
            Return _gdef
        End Get
        Set(ByVal a As Integer)
            _gdef = a
        End Set
    End Property
    Public Property nrep() As Integer
        Get
            Return _nrep
        End Get
        Set(ByVal a As Integer)
            _nrep = a
        End Set
    End Property
    Public Property nloc() As Integer
        Get
            Return _nloc
        End Get
        Set(ByVal a As Integer)
            _nloc = a
        End Set
    End Property
    Public Property page() As Integer
        Get
            Return _page
        End Get
        Set(ByVal a As Integer)
            _page = a
        End Set
    End Property
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        If Panel1.Visible = True Then
            _gdef = NumericUpDown2.Value
        Else
            _gdef = NumericUpDown1.Value
        End If
        If TabControl1.SelectedIndex = 3 Then
            _gdef = NumericUpDown3.Value
        End If
        If _page = 7 Then
            _gdef = NumericUpDown4.Value
        End If
        If _page = 5 Then
            If TextBox1.Text <> Nothing Then
                _nrep = TextBox1.Text
            End If
        ElseIf _page = 6 Then
            _nrep = TextBox4.Text
        ElseIf _page = 4 Then
            _nrep = TextBox5.Text
        ElseIf TextBox2.Text <> "" Then
            _nrep = TextBox2.Text
        Else
            _nrep = 0
        End If

        If TextBox3.Text <> Nothing Then
            cutoff = TextBox3.Text
        Else
            cutoff = 0
        End If
        If TabControl1.SelectedTab Is TabPage7 Then
            savetrees = CheckBox8.Checked
        Else
            savetrees = CheckBox2.Checked
            If CheckBox6.Checked = True Then
                savetrees = True
            End If
        End If
        If _page = 4 Then
            supp = CheckBox9.Checked
        
        Else
            supp = CheckBox3.Checked
        End If
        viewpanel1 = CheckBox10.Checked
        hethand = ComboBox1.SelectedIndex
        If CheckBox4.Checked = True Or CheckBox5.Checked = True Then
            bionj = True
        Else
            bionj = False
        End If
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        TabControl1.SelectTab(page)
        NumericUpDown1.Maximum = _nloc
        NumericUpDown1.Value = _nloc
        NumericUpDown2.Maximum = _nloc
        NumericUpDown3.Maximum = _nloc
        NumericUpDown2.Value = _nloc
        ComboBox1.SelectedIndex = Form1.hethandpp
        If viewpanel1 = True Then
            Panel1.Visible = True
            TextBox3.Visible = True
            Label7.Visible = True

        Else
            Panel1.Visible = False
            TextBox3.Visible = False
            Label7.Visible = False
        End If
        If viewboots = True Then
            CheckBox1.Enabled = True
            CheckBox6.Enabled = True
        Else
            CheckBox1.Enabled = False
            CheckBox6.Enabled = False
        End If
        If supp = True Then
            CheckBox3.Visible = True
        Else
            CheckBox3.Visible = False
        End If

        If TextBox3.Text <> Nothing Then
            cutoff = TextBox3.Text
        Else
            cutoff = 0
        End If
        hethand = ComboBox1.SelectedIndex
        If bionj = True Then
            CheckBox4.Visible = True
            CheckBox10.Visible = True
        End If
        If viewpanel1 = False Then
            RadioButton1.Visible = False
            RadioButton2.Visible = False
            Label8.Visible = True
            TextBox4.Visible = True
        End If
    End Sub

 

    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        If CheckBox1.Checked = True Then
            TextBox2.Visible = True
            TextBox2.Focus()
            TextBox2.Text = "100"
            Label4.Visible = True
            'CheckBox3.Checked = False
            'CheckBox10.Checked = False
        Else
            TextBox2.Visible = False
            Label4.Visible = False
            TextBox2.Text = Nothing
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    

    
    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
     
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            Label9.Visible = True
            TextBox5.Visible = True
            TextBox5.Text = 100
            CheckBox9.Checked = False

        Else
            Label9.Visible = False
            TextBox5.Visible = False
            TextBox5.Text = 0
        End If

    End Sub

  

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            CheckBox7.Checked = False
        End If
    End Sub

  

    
   
    
    
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If sender.checked = True Then
            Label8.Visible = False
            CheckBox8.Visible = False
            TextBox4.Visible = False
            templeton = True
        Else
            Label8.Visible = True
            CheckBox8.Visible = True
            TextBox4.Visible = True
            templeton = False
        End If
    End Sub
End Class
