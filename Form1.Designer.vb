<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.BtnExit = New System.Windows.Forms.Button()
        Me.BtnAddAccount = New System.Windows.Forms.Button()
        Me.BtnAddTrigger = New System.Windows.Forms.Button()
        Me.BtnListAccounts = New System.Windows.Forms.Button()
        Me.ListActions = New System.Windows.Forms.ListBox()
        Me.ListTriggers = New System.Windows.Forms.ListBox()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.ListAccounts = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblAllTriggers = New System.Windows.Forms.Label()
        Me.LblAllActions = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BtnExit
        '
        Me.BtnExit.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnExit.Location = New System.Drawing.Point(1016, 1018)
        Me.BtnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(150, 54)
        Me.BtnExit.TabIndex = 1
        Me.BtnExit.Text = "Exit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnAddAccount
        '
        Me.BtnAddAccount.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnAddAccount.Location = New System.Drawing.Point(348, 1020)
        Me.BtnAddAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnAddAccount.Name = "BtnAddAccount"
        Me.BtnAddAccount.Size = New System.Drawing.Size(150, 50)
        Me.BtnAddAccount.TabIndex = 2
        Me.BtnAddAccount.Text = "Add Account"
        Me.BtnAddAccount.UseVisualStyleBackColor = True
        '
        'BtnAddTrigger
        '
        Me.BtnAddTrigger.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnAddTrigger.Location = New System.Drawing.Point(676, 1020)
        Me.BtnAddTrigger.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnAddTrigger.Name = "BtnAddTrigger"
        Me.BtnAddTrigger.Size = New System.Drawing.Size(150, 50)
        Me.BtnAddTrigger.TabIndex = 3
        Me.BtnAddTrigger.Text = "Add Trigger"
        Me.BtnAddTrigger.UseVisualStyleBackColor = True
        '
        'BtnListAccounts
        '
        Me.BtnListAccounts.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnListAccounts.Location = New System.Drawing.Point(28, 1018)
        Me.BtnListAccounts.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnListAccounts.Name = "BtnListAccounts"
        Me.BtnListAccounts.Size = New System.Drawing.Size(150, 50)
        Me.BtnListAccounts.TabIndex = 4
        Me.BtnListAccounts.Text = "List Accounts"
        Me.BtnListAccounts.UseVisualStyleBackColor = True
        '
        'ListActions
        '
        Me.ListActions.ForeColor = System.Drawing.SystemColors.Highlight
        Me.ListActions.FormattingEnabled = True
        Me.ListActions.ItemHeight = 24
        Me.ListActions.Location = New System.Drawing.Point(28, 489)
        Me.ListActions.Margin = New System.Windows.Forms.Padding(4)
        Me.ListActions.Name = "ListActions"
        Me.ListActions.Size = New System.Drawing.Size(1141, 172)
        Me.ListActions.TabIndex = 6
        '
        'ListTriggers
        '
        Me.ListTriggers.ForeColor = System.Drawing.SystemColors.Highlight
        Me.ListTriggers.FormattingEnabled = True
        Me.ListTriggers.ItemHeight = 24
        Me.ListTriggers.Location = New System.Drawing.Point(32, 272)
        Me.ListTriggers.Margin = New System.Windows.Forms.Padding(4)
        Me.ListTriggers.Name = "ListTriggers"
        Me.ListTriggers.Size = New System.Drawing.Size(1141, 172)
        Me.ListTriggers.TabIndex = 7
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Location = New System.Drawing.Point(32, 688)
        Me.WebBrowser1.Margin = New System.Windows.Forms.Padding(4)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(21, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(1138, 322)
        Me.WebBrowser1.TabIndex = 8
        '
        'ListAccounts
        '
        Me.ListAccounts.ForeColor = System.Drawing.SystemColors.Highlight
        Me.ListAccounts.FormattingEnabled = True
        Me.ListAccounts.ItemHeight = 24
        Me.ListAccounts.Location = New System.Drawing.Point(37, 56)
        Me.ListAccounts.Margin = New System.Windows.Forms.Padding(4)
        Me.ListAccounts.Name = "ListAccounts"
        Me.ListAccounts.Size = New System.Drawing.Size(1126, 172)
        Me.ListAccounts.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label1.Location = New System.Drawing.Point(32, 28)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(143, 25)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "All Accounts"
        '
        'LblAllTriggers
        '
        Me.LblAllTriggers.AutoSize = True
        Me.LblAllTriggers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAllTriggers.ForeColor = System.Drawing.SystemColors.Highlight
        Me.LblAllTriggers.Location = New System.Drawing.Point(32, 244)
        Me.LblAllTriggers.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.LblAllTriggers.Name = "LblAllTriggers"
        Me.LblAllTriggers.Size = New System.Drawing.Size(133, 25)
        Me.LblAllTriggers.TabIndex = 12
        Me.LblAllTriggers.Text = "All Triggers"
        '
        'LblAllActions
        '
        Me.LblAllActions.AutoSize = True
        Me.LblAllActions.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAllActions.ForeColor = System.Drawing.SystemColors.Highlight
        Me.LblAllActions.Location = New System.Drawing.Point(26, 462)
        Me.LblAllActions.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.LblAllActions.Name = "LblAllActions"
        Me.LblAllActions.Size = New System.Drawing.Size(124, 25)
        Me.LblAllActions.TabIndex = 13
        Me.LblAllActions.Text = "All Actions"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1213, 1096)
        Me.Controls.Add(Me.LblAllActions)
        Me.Controls.Add(Me.LblAllTriggers)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListAccounts)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.ListTriggers)
        Me.Controls.Add(Me.ListActions)
        Me.Controls.Add(Me.BtnListAccounts)
        Me.Controls.Add(Me.BtnAddTrigger)
        Me.Controls.Add(Me.BtnAddAccount)
        Me.Controls.Add(Me.BtnExit)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form1"
        Me.Text = "ThinkAutomationReports"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnExit As Button
    Friend WithEvents BtnAddAccount As Button
    Friend WithEvents BtnAddTrigger As Button
    Friend WithEvents BtnListAccounts As Button
    Friend WithEvents ListActions As ListBox
    Friend WithEvents ListTriggers As ListBox
    Friend WithEvents WebBrowser1 As WebBrowser
    Friend WithEvents ListAccounts As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents LblAllTriggers As Label
    Friend WithEvents LblAllActions As Label
End Class
