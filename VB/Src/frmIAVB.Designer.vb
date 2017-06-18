Namespace IAVB
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmIAVB
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        If bDebug Then Me.StartPosition = FormStartPosition.CenterScreen

    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
            If Not m_synth Is Nothing Then ' CA2213
                m_synth.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdSuivant As System.Windows.Forms.Button
    Public WithEvents cmdGo As System.Windows.Forms.Button
    Public WithEvents listAssert As System.Windows.Forms.ListBox
    Public WithEvents listIA As System.Windows.Forms.ListBox
    Public WithEvents textInput As System.Windows.Forms.TextBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIAVB))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdSuivant = New System.Windows.Forms.Button()
        Me.cmdGo = New System.Windows.Forms.Button()
        Me.listAssert = New System.Windows.Forms.ListBox()
        Me.listIA = New System.Windows.Forms.ListBox()
        Me.textInput = New System.Windows.Forms.TextBox()
        Me.cmdInstall = New System.Windows.Forms.Button()
        Me.listParole = New System.Windows.Forms.ListBox()
        Me.chkAuto = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'cmdSuivant
        '
        Me.cmdSuivant.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSuivant.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSuivant.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSuivant.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSuivant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSuivant.Location = New System.Drawing.Point(523, 232)
        Me.cmdSuivant.Name = "cmdSuivant"
        Me.cmdSuivant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSuivant.Size = New System.Drawing.Size(33, 33)
        Me.cmdSuivant.TabIndex = 4
        Me.cmdSuivant.Text = "->"
        Me.ToolTip1.SetToolTip(Me.cmdSuivant, "Exemple suivant")
        Me.cmdSuivant.UseVisualStyleBackColor = False
        '
        'cmdGo
        '
        Me.cmdGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdGo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGo.Location = New System.Drawing.Point(468, 232)
        Me.cmdGo.Name = "cmdGo"
        Me.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGo.Size = New System.Drawing.Size(33, 33)
        Me.cmdGo.TabIndex = 3
        Me.cmdGo.Text = "!"
        Me.ToolTip1.SetToolTip(Me.cmdGo, "Exécution (ou entrée)")
        Me.cmdGo.UseVisualStyleBackColor = False
        '
        'listAssert
        '
        Me.listAssert.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listAssert.BackColor = System.Drawing.SystemColors.Window
        Me.listAssert.Cursor = System.Windows.Forms.Cursors.Default
        Me.listAssert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listAssert.ForeColor = System.Drawing.SystemColors.WindowText
        Me.listAssert.ItemHeight = 16
        Me.listAssert.Location = New System.Drawing.Point(8, 296)
        Me.listAssert.Name = "listAssert"
        Me.listAssert.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.listAssert.Size = New System.Drawing.Size(563, 84)
        Me.listAssert.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.listAssert, "Exemples à tester (Click ou DblClick)")
        '
        'listIA
        '
        Me.listIA.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listIA.BackColor = System.Drawing.Color.White
        Me.listIA.Cursor = System.Windows.Forms.Cursors.Default
        Me.listIA.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listIA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.listIA.ItemHeight = 16
        Me.listIA.Location = New System.Drawing.Point(8, 88)
        Me.listIA.Name = "listIA"
        Me.listIA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.listIA.Size = New System.Drawing.Size(563, 116)
        Me.listIA.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.listIA, "IAVB")
        '
        'textInput
        '
        Me.textInput.AcceptsReturn = True
        Me.textInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.textInput.BackColor = System.Drawing.SystemColors.Window
        Me.textInput.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.textInput.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textInput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.textInput.Location = New System.Drawing.Point(8, 232)
        Me.textInput.MaxLength = 0
        Me.textInput.Name = "textInput"
        Me.textInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.textInput.Size = New System.Drawing.Size(436, 22)
        Me.textInput.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.textInput, "Zone de saisie")
        '
        'cmdInstall
        '
        Me.cmdInstall.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdInstall.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInstall.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInstall.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInstall.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInstall.Location = New System.Drawing.Point(484, 12)
        Me.cmdInstall.Name = "cmdInstall"
        Me.cmdInstall.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInstall.Size = New System.Drawing.Size(87, 33)
        Me.cmdInstall.TabIndex = 5
        Me.cmdInstall.Text = "Installation"
        Me.ToolTip1.SetToolTip(Me.cmdInstall, "Vérifier l'installation de MSAgent et des composants de synthèse vocal pour MSAge" & _
        "nt")
        Me.cmdInstall.UseVisualStyleBackColor = False
        '
        'listParole
        '
        Me.listParole.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listParole.BackColor = System.Drawing.Color.White
        Me.listParole.Cursor = System.Windows.Forms.Cursors.Default
        Me.listParole.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listParole.ForeColor = System.Drawing.SystemColors.WindowText
        Me.listParole.ItemHeight = 16
        Me.listParole.Location = New System.Drawing.Point(8, 12)
        Me.listParole.Name = "listParole"
        Me.listParole.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.listParole.Size = New System.Drawing.Size(324, 68)
        Me.listParole.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.listParole, "IAVB")
        '
        'chkAuto
        '
        Me.chkAuto.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAuto.AutoSize = True
        Me.chkAuto.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAuto.Location = New System.Drawing.Point(517, 270)
        Me.chkAuto.Name = "chkAuto"
        Me.chkAuto.Size = New System.Drawing.Size(54, 20)
        Me.chkAuto.TabIndex = 8
        Me.chkAuto.Text = "Auto"
        Me.ToolTip1.SetToolTip(Me.chkAuto, "Cocher pour passer automatiquement à l'exemple suivant")
        Me.chkAuto.UseVisualStyleBackColor = True
        '
        'frmIAVB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(580, 401)
        Me.Controls.Add(Me.chkAuto)
        Me.Controls.Add(Me.listParole)
        Me.Controls.Add(Me.cmdInstall)
        Me.Controls.Add(Me.cmdSuivant)
        Me.Controls.Add(Me.cmdGo)
        Me.Controls.Add(Me.listAssert)
        Me.Controls.Add(Me.listIA)
        Me.Controls.Add(Me.textInput)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(11, 30)
        Me.Name = "frmIAVB"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Intelligence Artificielle Vraiment Basique"
        Me.ResumeLayout(False)
        Me.PerformLayout()

End Sub
    Public WithEvents cmdInstall As System.Windows.Forms.Button
    Public WithEvents listParole As System.Windows.Forms.ListBox
    Friend WithEvents chkAuto As System.Windows.Forms.CheckBox
#End Region
End Class
End Namespace