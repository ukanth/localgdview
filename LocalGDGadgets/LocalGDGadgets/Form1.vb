Imports ICSharpCode.SharpZipLib
Imports System.Xml


Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim dra As IO.FileInfo
    Dim filePath As String

    Dim fileSize(250) As String
    Dim fileName(250) As String
    Dim fullPath(250) As String
    Dim fileDate(250) As String

    Dim a() As String

    Dim grdTemp As DataGrid

    Dim x As Integer

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gName As System.Windows.Forms.TextBox
    Friend WithEvents aName As System.Windows.Forms.TextBox
    Friend WithEvents gSize As System.Windows.Forms.TextBox
    Friend WithEvents gDate As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.gName = New System.Windows.Forms.TextBox
        Me.aName = New System.Windows.Forms.TextBox
        Me.gSize = New System.Windows.Forms.TextBox
        Me.gDate = New System.Windows.Forms.TextBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Location = New System.Drawing.Point(8, 8)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(296, 394)
        Me.ListBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(368, 272)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(168, 32)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Open Gadget"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Location = New System.Drawing.Point(312, 112)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 24)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Gadget Name "
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(312, 192)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 24)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Gadget Size"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(312, 232)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 24)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Stored Date"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(312, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 24)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Author Name"
        '
        'gName
        '
        Me.gName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.gName.Location = New System.Drawing.Point(400, 112)
        Me.gName.Name = "gName"
        Me.gName.ReadOnly = True
        Me.gName.Size = New System.Drawing.Size(176, 14)
        Me.gName.TabIndex = 10
        Me.gName.Text = ""
        '
        'aName
        '
        Me.aName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.aName.Location = New System.Drawing.Point(400, 152)
        Me.aName.Name = "aName"
        Me.aName.ReadOnly = True
        Me.aName.Size = New System.Drawing.Size(176, 14)
        Me.aName.TabIndex = 11
        Me.aName.Text = ""
        '
        'gSize
        '
        Me.gSize.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.gSize.Location = New System.Drawing.Point(400, 192)
        Me.gSize.Name = "gSize"
        Me.gSize.ReadOnly = True
        Me.gSize.Size = New System.Drawing.Size(96, 14)
        Me.gSize.TabIndex = 12
        Me.gSize.Text = ""
        '
        'gDate
        '
        Me.gDate.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.gDate.Location = New System.Drawing.Point(400, 232)
        Me.gDate.Name = "gDate"
        Me.gDate.ReadOnly = True
        Me.gDate.Size = New System.Drawing.Size(128, 14)
        Me.gDate.TabIndex = 13
        Me.gDate.Text = ""
        '
        'PictureBox2
        '
        Me.PictureBox2.Location = New System.Drawing.Point(416, 56)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 14
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(464, 320)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(80, 80)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(312, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 24)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Gadget  Logo"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(320, 352)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 24)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Drag Gadget Here --->"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(586, 415)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.gDate)
        Me.Controls.Add(Me.gSize)
        Me.Controls.Add(Me.aName)
        Me.Controls.Add(Me.gName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "Local Google Desktop Gadgets by Umakanth"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PictureBox1.AllowDrop = True
        x = 0
        Try
            Dim di As New IO.DirectoryInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\My Google Gadgets")
            Dim diar1 As IO.FileInfo() = di.GetFiles()
            filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\My Google Gadgets\"
            For Each dra In diar1
                ListBox1.Items.Add(dra)
                fileSize(x) = dra.Length
                fullPath(x) = dra.FullName.ToString()
                fileName(x) = dra.Name.ToString()
                fileDate(x) = dra.CreationTime()
                x = x + 1
            Next
            ListBox1.SetSelected(0, True)
        Catch ex As Exception
            MessageBox.Show("Please check that you have installed Google Desktop on this machine", "Missing Google Desktop", MessageBoxButtons.OK)
            Exit Sub
        Finally
        End Try
        'list the names of all files in the specified directory
    End Sub


    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        deleteLogo("C:\temp\tmp\icon_large.png")
        deleteLogo("C:\temp\tmp\plugin_large.gif")
        'deleteLogo("C:\temp\tmp\gadget.gmanifest")
        Dim fz As New ICSharpCode.SharpZipLib.Zip.FastZip
        fz.ExtractZip(fullPath(ListBox1.SelectedIndex).ToString, "C:\temp\tmp", "icon_large.png")
        fz.ExtractZip(fullPath(ListBox1.SelectedIndex).ToString, "C:\temp\tmp", "plugin_large.gif")
        Button1.Enabled = True
        Dim size As Double = RoundDown(fileSize(ListBox1.SelectedIndex) / 1024, 2)
        Dim name() As String = Split(fileName(ListBox1.SelectedIndex).ToString, "-")
        If (name.Length > 1) Then
            gName.Text = name(0)
            Dim authorName() As String = name(1).Split(".")
            aName.Text = authorName(0)
        Else
            gName.Text = fileName(ListBox1.SelectedIndex)
            aName.Text = "Unknown"
        End If
        gSize.Text = size.ToString + " KB"
        gDate.Text = fileDate(ListBox1.SelectedIndex)
        If (isLogo("C:\Temp\tmp", "icon_large.png")) Then
            Dim MyBitmap As New Bitmap("C:\Temp\tmp\icon_large.png")
            PictureBox2.Image = New Bitmap(MyBitmap)
            MyBitmap.Dispose()
        ElseIf (isLogo("C:\Temp\tmp", "plugin_large.gif")) Then
            Dim MyBitmap As New Bitmap("C:\Temp\tmp\plugin_large.gif")
            PictureBox2.Image = New Bitmap(MyBitmap)
            MyBitmap.Dispose()
        End If
        'parseXML()
        'PictureBox2.Image.FromFile("C:\temp\tmp\icon_large.png")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Process.Start(fullPath(ListBox1.SelectedIndex))
        Catch ex As Exception
            MessageBox.Show("Please check that you have installed Google Desktop on this machine", "Missing Google Desktop", MessageBoxButtons.OK)
            Exit Sub
        Finally
        End Try
    End Sub

    Function alert(ByVal str As String)
        MessageBox.Show(str, "Debug", MessageBoxButtons.OK)
    End Function

    Function deleteLogo(ByVal FileToDelete As String) As Boolean
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If
    End Function

    Function isLogo(ByVal FolderPath, ByVal FileName) As Boolean
        Dim objFSO
        Dim objFolder, file
        Dim FileExists As Boolean

        objFSO = CreateObject("Scripting.FileSystemObject")
        objFolder = objFSO.GetFolder(FolderPath)
        For Each file In objFolder.files
            If LCase(file.name) = Trim(LCase(FileName)) Then
                FileExists = True
            Else
                FileExists = False
            End If
        Next
        objFSO = Nothing
        Return FileExists
    End Function

    Function RoundDown(ByVal number As Double, ByVal num_digits As Integer) As Double
        Return Math.Floor((10 ^ num_digits) * number) / (10 ^ num_digits)
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PictureBox2.Image = Image.FromFile("C:\Temp\tmp\defau.gif")
    End Sub

    Private Sub PictureBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PictureBox1.DragDrop
        Try
            Dim fz As New ICSharpCode.SharpZipLib.Zip.FastZip
            Dim dragFile As String
            dragFile = CType(e.Data.GetData(DataFormats.FileDrop), Array).GetValue(0).ToString()
            fz.ExtractZip(dragFile, "C:\temp\tmp", "icon_large.png")
            fz.ExtractZip(dragFile, "C:\temp\tmp", "plugin_large.gif")
            Button1.Enabled = True
            gName.Text = "-"
            aName.Text = "-"
            gSize.Text = "-"
            gDate.Text = "-"
            If (isLogo("C:\Temp\tmp", "icon_large.png")) Then
                Dim MyBitmap As New Bitmap("C:\Temp\tmp\icon_large.png")
                PictureBox2.Image = New Bitmap(MyBitmap)
                MyBitmap.Dispose()
            ElseIf (isLogo("C:\Temp\tmp", "plugin_large.gif")) Then
                Dim MyBitmap As New Bitmap("C:\Temp\tmp\plugin_large.gif")
                PictureBox2.Image = New Bitmap(MyBitmap)
                MyBitmap.Dispose()
            End If
        Catch ex As Exception
            MessageBox.Show("Only Support Google Desktop Gadgets(*.gg)", "Gadgets Drag Here", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub PictureBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PictureBox1.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        deleteLogo("C:\temp\tmp\icon_large.png")
        deleteLogo("C:\temp\tmp\plugin_large.gif")
    End Sub
End Class
