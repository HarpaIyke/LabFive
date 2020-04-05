<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmTextEditor
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTextEditor))
        Me.mnuTopMenu = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtTextEditor = New System.Windows.Forms.TextBox()
        Me.opdOpen = New System.Windows.Forms.OpenFileDialog()
        Me.sfdSaveAs = New System.Windows.Forms.SaveFileDialog()
        Me.ttpTextEditor = New System.Windows.Forms.ToolTip(Me.components)
        Me.mnuTopMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuTopMenu
        '
        Me.mnuTopMenu.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.mnuTopMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit, Me.mnuHelp})
        Me.mnuTopMenu.Location = New System.Drawing.Point(0, 0)
        Me.mnuTopMenu.Name = "mnuTopMenu"
        Me.mnuTopMenu.Size = New System.Drawing.Size(800, 28)
        Me.mnuTopMenu.TabIndex = 0
        Me.mnuTopMenu.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuFileSave, Me.mnuFileSaveAs, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(44, 24)
        Me.mnuFile.Text = "&File"
        Me.mnuFile.ToolTipText = "Access the File submenu"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Name = "mnuFileNew"
        Me.mnuFileNew.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.mnuFileNew.Size = New System.Drawing.Size(216, 26)
        Me.mnuFileNew.Text = "&New"
        Me.mnuFileNew.ToolTipText = "Create a new file"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnuFileOpen.Size = New System.Drawing.Size(216, 26)
        Me.mnuFileOpen.Text = "&Open"
        Me.mnuFileOpen.ToolTipText = "Open an existing file"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.mnuFileSave.Size = New System.Drawing.Size(216, 26)
        Me.mnuFileSave.Text = "&Save"
        Me.mnuFileSave.ToolTipText = "Save the current file"
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Name = "mnuFileSaveAs"
        Me.mnuFileSaveAs.Size = New System.Drawing.Size(216, 26)
        Me.mnuFileSaveAs.Text = "Save &As..."
        Me.mnuFileSaveAs.ToolTipText = "Save a copy of this file"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(216, 26)
        Me.mnuFileExit.Text = "E&xit"
        Me.mnuFileExit.ToolTipText = "Exit the application"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuEditCopy, Me.mnuEditCut, Me.mnuEditPaste})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(47, 24)
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Name = "mnuEditCopy"
        Me.mnuEditCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuEditCopy.Size = New System.Drawing.Size(216, 26)
        Me.mnuEditCopy.Text = "&Copy"
        Me.mnuEditCopy.ToolTipText = "Copy the selected text"
        '
        'mnuEditCut
        '
        Me.mnuEditCut.Name = "mnuEditCut"
        Me.mnuEditCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuEditCut.Size = New System.Drawing.Size(216, 26)
        Me.mnuEditCut.Text = "Cu&t"
        Me.mnuEditCut.ToolTipText = "Cut the selected text"
        '
        'mnuEditPaste
        '
        Me.mnuEditPaste.Name = "mnuEditPaste"
        Me.mnuEditPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.mnuEditPaste.Size = New System.Drawing.Size(216, 26)
        Me.mnuEditPaste.Text = "&Paste"
        Me.mnuEditPaste.ToolTipText = "Paste any text that has been copied"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(53, 24)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnuHelpAbout.Size = New System.Drawing.Size(216, 26)
        Me.mnuHelpAbout.Text = "&About"
        Me.mnuHelpAbout.ToolTipText = "Show information about the creation of this application"
        '
        'txtTextEditor
        '
        Me.txtTextEditor.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtTextEditor.Font = New System.Drawing.Font("Consolas", 11.0!)
        Me.txtTextEditor.Location = New System.Drawing.Point(0, 28)
        Me.txtTextEditor.Multiline = True
        Me.txtTextEditor.Name = "txtTextEditor"
        Me.txtTextEditor.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtTextEditor.Size = New System.Drawing.Size(800, 422)
        Me.txtTextEditor.TabIndex = 1
        '
        'frmTextEditor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtTextEditor)
        Me.Controls.Add(Me.mnuTopMenu)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mnuTopMenu
        Me.Name = "frmTextEditor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Text Editor"
        Me.mnuTopMenu.ResumeLayout(False)
        Me.mnuTopMenu.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents mnuTopMenu As MenuStrip
    Friend WithEvents mnuFile As ToolStripMenuItem
    Friend WithEvents mnuFileNew As ToolStripMenuItem
    Friend WithEvents mnuFileOpen As ToolStripMenuItem
    Friend WithEvents mnuEdit As ToolStripMenuItem
    Friend WithEvents mnuEditCopy As ToolStripMenuItem
    Friend WithEvents txtTextEditor As TextBox
    Friend WithEvents opdOpen As OpenFileDialog
    Friend WithEvents mnuFileSave As ToolStripMenuItem
    Friend WithEvents mnuFileSaveAs As ToolStripMenuItem
    Friend WithEvents mnuFileExit As ToolStripMenuItem
    Friend WithEvents sfdSaveAs As SaveFileDialog
    Friend WithEvents mnuEditCut As ToolStripMenuItem
    Friend WithEvents mnuEditPaste As ToolStripMenuItem
    Friend WithEvents mnuHelp As ToolStripMenuItem
    Friend WithEvents mnuHelpAbout As ToolStripMenuItem
    Friend WithEvents ttpTextEditor As ToolTip
End Class
