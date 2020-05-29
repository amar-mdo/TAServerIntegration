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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BtnExit
        '
        Me.BtnExit.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnExit.Location = New System.Drawing.Point(739, 680)
        Me.BtnExit.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(109, 36)
        Me.BtnExit.TabIndex = 1
        Me.BtnExit.Text = "Exit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnAddAccount
        '
        Me.BtnAddAccount.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnAddAccount.Location = New System.Drawing.Point(359, 683)
        Me.BtnAddAccount.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnAddAccount.Name = "BtnAddAccount"
        Me.BtnAddAccount.Size = New System.Drawing.Size(109, 33)
        Me.BtnAddAccount.TabIndex = 2
        Me.BtnAddAccount.Text = "Add Account"
        Me.BtnAddAccount.UseVisualStyleBackColor = True
        '
        'BtnAddTrigger
        '
        Me.BtnAddTrigger.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnAddTrigger.Location = New System.Drawing.Point(485, 683)
        Me.BtnAddTrigger.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnAddTrigger.Name = "BtnAddTrigger"
        Me.BtnAddTrigger.Size = New System.Drawing.Size(109, 33)
        Me.BtnAddTrigger.TabIndex = 3
        Me.BtnAddTrigger.Text = "Add Trigger"
        Me.BtnAddTrigger.UseVisualStyleBackColor = True
        '
        'BtnListAccounts
        '
        Me.BtnListAccounts.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnListAccounts.Location = New System.Drawing.Point(20, 680)
        Me.BtnListAccounts.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnListAccounts.Name = "BtnListAccounts"
        Me.BtnListAccounts.Size = New System.Drawing.Size(109, 33)
        Me.BtnListAccounts.TabIndex = 4
        Me.BtnListAccounts.Text = "List Accounts"
        Me.BtnListAccounts.UseVisualStyleBackColor = True
        '
        'ListActions
        '
        Me.ListActions.ForeColor = System.Drawing.SystemColors.Highlight
        Me.ListActions.FormattingEnabled = True
        Me.ListActions.ItemHeight = 16
        Me.ListActions.Location = New System.Drawing.Point(20, 301)
        Me.ListActions.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ListActions.Name = "ListActions"
        Me.ListActions.Size = New System.Drawing.Size(830, 116)
        Me.ListActions.TabIndex = 6
        '
        'ListTriggers
        '
        Me.ListTriggers.ForeColor = System.Drawing.SystemColors.Highlight
        Me.ListTriggers.FormattingEnabled = True
        Me.ListTriggers.ItemHeight = 16
        Me.ListTriggers.Location = New System.Drawing.Point(23, 181)
        Me.ListTriggers.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ListTriggers.Name = "ListTriggers"
        Me.ListTriggers.Size = New System.Drawing.Size(830, 116)
        Me.ListTriggers.TabIndex = 7
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Location = New System.Drawing.Point(23, 430)
        Me.WebBrowser1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(15, 14)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(828, 234)
        Me.WebBrowser1.TabIndex = 8
        '
        'ListAccounts
        '
        Me.ListAccounts.ForeColor = System.Drawing.SystemColors.Highlight
        Me.ListAccounts.FormattingEnabled = True
        Me.ListAccounts.ItemHeight = 16
        Me.ListAccounts.Location = New System.Drawing.Point(27, 37)
        Me.ListAccounts.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ListAccounts.Name = "ListAccounts"
        Me.ListAccounts.Size = New System.Drawing.Size(820, 116)
        Me.ListAccounts.TabIndex = 9
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Button1.Location = New System.Drawing.Point(613, 681)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(108, 33)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Add Action"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(881, 727)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ListAccounts)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.ListTriggers)
        Me.Controls.Add(Me.ListActions)
        Me.Controls.Add(Me.BtnListAccounts)
        Me.Controls.Add(Me.BtnAddTrigger)
        Me.Controls.Add(Me.BtnAddAccount)
        Me.Controls.Add(Me.BtnExit)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "Form1"
        Me.Text = "ThinkAutomationReports"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnExit As Button
    Friend WithEvents BtnAddAccount As Button
    Friend WithEvents BtnAddTrigger As Button
    Friend WithEvents BtnListAccounts As Button
    Friend WithEvents ListActions As ListBox
    Friend WithEvents ListTriggers As ListBox
    Friend WithEvents WebBrowser1 As WebBrowser
    Friend WithEvents ListAccounts As ListBox
    Friend WithEvents Button1 As Button
End Class
