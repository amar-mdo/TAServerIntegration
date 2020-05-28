<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.BtnTestTAServer = New System.Windows.Forms.Button()
        Me.BtnExit = New System.Windows.Forms.Button()
        Me.BtnAddAccount = New System.Windows.Forms.Button()
        Me.BtnAddTrigger = New System.Windows.Forms.Button()
        Me.BtnListAccounts = New System.Windows.Forms.Button()
        Me.BtnListTriggers = New System.Windows.Forms.Button()
        Me.ListActions = New System.Windows.Forms.ListBox()
        Me.ListTriggers = New System.Windows.Forms.ListBox()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.ListAccounts = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'BtnTestTAServer
        '
        Me.BtnTestTAServer.Location = New System.Drawing.Point(27, 917)
        Me.BtnTestTAServer.Name = "BtnTestTAServer"
        Me.BtnTestTAServer.Size = New System.Drawing.Size(150, 50)
        Me.BtnTestTAServer.TabIndex = 0
        Me.BtnTestTAServer.Text = "Test TA Server"
        Me.BtnTestTAServer.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Location = New System.Drawing.Point(977, 914)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(114, 53)
        Me.BtnExit.TabIndex = 1
        Me.BtnExit.Text = "Exit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'BtnAddAccount
        '
        Me.BtnAddAccount.Location = New System.Drawing.Point(627, 917)
        Me.BtnAddAccount.Name = "BtnAddAccount"
        Me.BtnAddAccount.Size = New System.Drawing.Size(150, 50)
        Me.BtnAddAccount.TabIndex = 2
        Me.BtnAddAccount.Text = "Add Account"
        Me.BtnAddAccount.UseVisualStyleBackColor = True
        '
        'BtnAddTrigger
        '
        Me.BtnAddTrigger.Location = New System.Drawing.Point(803, 917)
        Me.BtnAddTrigger.Name = "BtnAddTrigger"
        Me.BtnAddTrigger.Size = New System.Drawing.Size(150, 50)
        Me.BtnAddTrigger.TabIndex = 3
        Me.BtnAddTrigger.Text = "Add Trigger"
        Me.BtnAddTrigger.UseVisualStyleBackColor = True
        '
        'BtnListAccounts
        '
        Me.BtnListAccounts.Location = New System.Drawing.Point(238, 918)
        Me.BtnListAccounts.Name = "BtnListAccounts"
        Me.BtnListAccounts.Size = New System.Drawing.Size(150, 50)
        Me.BtnListAccounts.TabIndex = 4
        Me.BtnListAccounts.Text = "List Accounts"
        Me.BtnListAccounts.UseVisualStyleBackColor = True
        '
        'BtnListTriggers
        '
        Me.BtnListTriggers.Location = New System.Drawing.Point(406, 917)
        Me.BtnListTriggers.Name = "BtnListTriggers"
        Me.BtnListTriggers.Size = New System.Drawing.Size(150, 50)
        Me.BtnListTriggers.TabIndex = 5
        Me.BtnListTriggers.Text = "List Triggers"
        Me.BtnListTriggers.UseVisualStyleBackColor = True
        '
        'ListActions
        '
        Me.ListActions.FormattingEnabled = True
        Me.ListActions.ItemHeight = 24
        Me.ListActions.Location = New System.Drawing.Point(30, 486)
        Me.ListActions.Name = "ListActions"
        Me.ListActions.Size = New System.Drawing.Size(1061, 172)
        Me.ListActions.TabIndex = 6
        '
        'ListTriggers
        '
        Me.ListTriggers.FormattingEnabled = True
        Me.ListTriggers.ItemHeight = 24
        Me.ListTriggers.Location = New System.Drawing.Point(31, 272)
        Me.ListTriggers.Name = "ListTriggers"
        Me.ListTriggers.Size = New System.Drawing.Size(1060, 196)
        Me.ListTriggers.TabIndex = 7
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Location = New System.Drawing.Point(31, 676)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(1060, 222)
        Me.WebBrowser1.TabIndex = 8
        '
        'ListAccounts
        '
        Me.ListAccounts.FormattingEnabled = True
        Me.ListAccounts.ItemHeight = 24
        Me.ListAccounts.Location = New System.Drawing.Point(36, 55)
        Me.ListAccounts.Name = "ListAccounts"
        Me.ListAccounts.Size = New System.Drawing.Size(1054, 172)
        Me.ListAccounts.TabIndex = 9
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1206, 1000)
        Me.Controls.Add(Me.ListAccounts)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.ListTriggers)
        Me.Controls.Add(Me.ListActions)
        Me.Controls.Add(Me.BtnListTriggers)
        Me.Controls.Add(Me.BtnListAccounts)
        Me.Controls.Add(Me.BtnAddTrigger)
        Me.Controls.Add(Me.BtnAddAccount)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnTestTAServer)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BtnTestTAServer As Button
    Friend WithEvents BtnExit As Button
    Friend WithEvents BtnAddAccount As Button
    Friend WithEvents BtnAddTrigger As Button
    Friend WithEvents BtnListAccounts As Button
    Friend WithEvents BtnListTriggers As Button
    Friend WithEvents ListActions As ListBox
    Friend WithEvents ListTriggers As ListBox
    Friend WithEvents WebBrowser1 As WebBrowser
    Friend WithEvents ListAccounts As ListBox
End Class
