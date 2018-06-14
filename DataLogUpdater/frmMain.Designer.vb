<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDirectory = New System.Windows.Forms.TextBox()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.tvFolder = New System.Windows.Forms.TreeView()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblBrowse = New System.Windows.Forms.Label()
        Me.pnlPath = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblWork = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ToolStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pnlPath.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Folder Path"
        '
        'txtDirectory
        '
        Me.txtDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectory.Location = New System.Drawing.Point(117, 32)
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.ReadOnly = True
        Me.txtDirectory.Size = New System.Drawing.Size(635, 26)
        Me.txtDirectory.TabIndex = 1
        '
        'btnBrowse
        '
        Me.btnBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowse.Location = New System.Drawing.Point(758, 30)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(75, 29)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'tvFolder
        '
        Me.tvFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tvFolder.ImageIndex = 0
        Me.tvFolder.ImageList = Me.ImageList1
        Me.tvFolder.Location = New System.Drawing.Point(15, 95)
        Me.tvFolder.Name = "tvFolder"
        Me.tvFolder.SelectedImageIndex = 1
        Me.tvFolder.Size = New System.Drawing.Size(282, 320)
        Me.tvFolder.StateImageList = Me.ImageList1
        Me.tvFolder.TabIndex = 3
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "File-CSV-icon.png")
        Me.ImageList1.Images.SetKeyName(1, "Folder-icon.png")
        Me.ImageList1.Images.SetKeyName(2, "Folder-Open-icon.png")
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(845, 25)
        Me.ToolStrip1.TabIndex = 4
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(92, 22)
        Me.ToolStripButton1.Text = "Update Data"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 18)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Folder Contents"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.pnlPath)
        Me.Panel1.Location = New System.Drawing.Point(315, 95)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(518, 320)
        Me.Panel1.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(317, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(102, 18)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Progress Status"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(15, 426)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(818, 18)
        Me.ProgressBar1.TabIndex = 8
        '
        'lblStatus
        '
        Me.lblStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(17, 448)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(40, 14)
        Me.lblStatus.TabIndex = 9
        Me.lblStatus.Text = "Ready"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(18, 17)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(183, 26)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Browsing Directory:"
        '
        'lblBrowse
        '
        Me.lblBrowse.AutoSize = True
        Me.lblBrowse.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBrowse.Location = New System.Drawing.Point(261, 17)
        Me.lblBrowse.Name = "lblBrowse"
        Me.lblBrowse.Size = New System.Drawing.Size(18, 26)
        Me.lblBrowse.TabIndex = 1
        Me.lblBrowse.Text = "-"
        '
        'pnlPath
        '
        Me.pnlPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlPath.Controls.Add(Me.lblBrowse)
        Me.pnlPath.Controls.Add(Me.Label5)
        Me.pnlPath.Location = New System.Drawing.Point(3, 3)
        Me.pnlPath.Name = "pnlPath"
        Me.pnlPath.Size = New System.Drawing.Size(510, 61)
        Me.pnlPath.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.lblWork)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Location = New System.Drawing.Point(3, 70)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(510, 61)
        Me.Panel2.TabIndex = 1
        '
        'lblWork
        '
        Me.lblWork.AutoSize = True
        Me.lblWork.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWork.Location = New System.Drawing.Point(261, 17)
        Me.lblWork.Name = "lblWork"
        Me.lblWork.Size = New System.Drawing.Size(18, 26)
        Me.lblWork.TabIndex = 1
        Me.lblWork.Text = "-"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(18, 17)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(126, 26)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Working File:"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(845, 471)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.tvFolder)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.txtDirectory)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMain"
        Me.Text = "Data Log Updater"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.pnlPath.ResumeLayout(False)
        Me.pnlPath.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents txtDirectory As TextBox
    Friend WithEvents btnBrowse As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents tvFolder As TreeView
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents lblStatus As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents lblWork As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents pnlPath As Panel
    Friend WithEvents lblBrowse As Label
    Friend WithEvents Label5 As Label
End Class
