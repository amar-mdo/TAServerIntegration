Imports System
Imports System.Net
Imports ThinkAutomationCore
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Data
Public Class Form1

    Dim fPath As String = "C:\MDO\ThinkAutomationTest\"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AutoSize = True
    End Sub

    Private Sub ListAccounts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListAccounts.DoubleClick

        If ListAccounts.SelectedItems.Count = 0 Then Exit Sub
        Dim SelectedAccount As String = ListAccounts.SelectedItems(0)

        If SelectedAccount.Length = 0 Then
            MsgBox("An Account Selection is required, by clicking the listed AccounName twice!")
            Exit Sub
        Else
            Dim Confirmation As String = MsgBox("Do you want the Account details for (" + SelectedAccount + ") written to a Text File?", MsgBoxStyle.YesNo, "TA Server Automation")

            ListTriggers.Items.Clear()
            ListActions.Items.Clear()
            WebBrowser1.DocumentText = ""
            LblAllTriggers.Text = "All Triggers"
            LblAllActions.Text = "All Actions"

            ListAccountTriggers(SelectedAccount)

            If Confirmation = vbYes Then
                WriteAccountDetailsToFile(SelectedAccount)
            End If

        End If

    End Sub
    Private Sub ListTriggers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListTriggers.DoubleClick

        If ListTriggers.SelectedItems.Count = 0 Then Exit Sub

        Dim SelectedAccount As String = ListTriggers.Items(0)
        Dim SelectedTrigger As String = ListTriggers.SelectedItems(0)

        ' First Clear the List and then fill the actions list

        ListActions.Items.Clear()

        ListTriggerActions(SelectedAccount, LTrim(SelectedTrigger))

    End Sub
    Private Sub WriteAccountDetailsToFile(AccountName As String)
        If AccountName.Length = 0 Then Exit Sub

        Dim fFileName As String = AccountName + ".txt"

        Dim sw As StreamWriter = Nothing
        Try
            Dim ThisAccount As ThinkAutomation.clsAccount
            Dim ThisTrigger As ThinkAutomation.clsAccountTrigger
            Dim blnAccountFound As Boolean

            sw = New StreamWriter(fPath + fFileName)

            blnAccountFound = False

            ' login to ThinkAutomation on 127.0.0.1:8855
            If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
                For Each ThisAccount In ThinkAutomation.Accounts ' Accounts collection will contain all ThinkAutomation accounts after login

                    If ThisAccount.Name = AccountName Then
                        blnAccountFound = True
                        sw.WriteLine(ThisAccount.XML)
                        ThisAccount.Triggers.Load()
                        ' display all triggers for this account
                        For Each ThisTrigger In ThisAccount.Triggers
                            sw.WriteLine(ThisTrigger.XML)
                        Next
                    End If

                Next

                If Not blnAccountFound Then
                    MsgBox("Could not find the Account. " & ThinkAutomation.SystemErrorLast)
                End If

            Else
                MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
            End If

        Finally
            If Not sw Is Nothing Then
                sw.Dispose()
            End If
        End Try

    End Sub
    Private Sub AddAccount(strAccount As String)
        If strAccount.Length = 0 Then Exit Sub

        Dim intAccountId As Integer
        Dim strIMAPAccount As String
        Dim strDropBoxSubject As String
        Dim strIMAPServer As String
        Dim strIMAPFolder As String
        Dim strIMAPMoveToFolder As String
        Dim intCheckEveryMins As Integer

        Dim connectionString As String
        Dim connection, UpdateConnection As SqlConnection
        Dim command, commandUpdate As SqlCommand
        Dim sql, UpdateSql As String

        connectionString = "Data Source=169.53.175.117;UID=exceluser;PWD=c4Ua8qPNJ4X804r;DATABASE=Perspective"
        sql = "select * from taaccounts where hastaaccountsetup = 0 and AccountName = '" + strAccount + "'"

        connection = New SqlConnection(connectionString)
        Try
            connection.Open()
            command = New SqlCommand(sql, connection)
            command.ExecuteNonQuery()
            'Create a Recordset object, populate it with the query results
            Dim objDR As SqlDataReader = command.ExecuteReader()

            If Not (objDR Is Nothing) AndAlso objDR.HasRows Then
                While objDR.Read
                    intAccountId = objDR("AccountId")
                    strIMAPAccount = objDR("IMAPAccount").ToString
                    strDropBoxSubject = objDR("DropBoxSubject").ToString
                    strIMAPServer = objDR("IMAPServer").ToString
                    strIMAPFolder = objDR("IMAPFolder").ToString
                    strIMAPMoveToFolder = objDR("IMAPMoveToFolder").ToString
                    intCheckEveryMins = objDR("CheckEveryMins")

                    If CreateTAServerAccount(strAccount, strIMAPAccount, strDropBoxSubject, strIMAPServer, strIMAPFolder, strIMAPMoveToFolder, intCheckEveryMins) Then
                        UpdateConnection = New SqlConnection(connectionString)
                        UpdateConnection.Open()
                        UpdateSql = "update taaccounts set HasTAAccountSetUp = 1 where AccountId = " + CStr(intAccountId)
                        commandUpdate = New SqlCommand(UpdateSql, UpdateConnection)
                        commandUpdate.ExecuteNonQuery()
                        commandUpdate.Dispose()
                        UpdateConnection.Close()
                        MsgBox("Added the Account: " + strAccount + " to the TA Server !!")
                    Else
                        MsgBox("Unable to Add the Account: " + strAccount + " to the TA Server !!")
                    End If
                End While
            Else
                MsgBox("Unable to Find the Given Account in the Database" + vbNewLine + "The Account Name provided is either incorrect or it has already been Added to the TA Server !!")
            End If

            command.Dispose()
            connection.Close()
        Catch ex As Exception
            MsgBox("Perspective Database ERROR! Cause: " + ex.Message)
        End Try


    End Sub
    Private Function CreateTAServerAccount(strAccount As String, strIMAPAccount As String, strDropBoxSubject As String, strIMAPServer As String, strIMAPFolder As String, strIMAPMoveToFolder As String, intCheckEveryMins As Integer) As Boolean
        Dim blnAccountAdded As Boolean = vbFalse
        Dim blnAccountAlreadyExists As Boolean = vbFalse
        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then

            'First Check If the Account Already Exists
            For Each ThisAccount In ThinkAutomation.Accounts
                If ThisAccount.Name = strAccount Then
                    blnAccountAlreadyExists = vbTrue
                End If
            Next

            ' Add the Account only if it Does not Exist in the TA Server
            If blnAccountAlreadyExists Then
                MsgBox("TA Server Account: " + strAccount + " already Exists cannot add Duplicate!!")
            Else
                Dim NewAccount As New ThinkAutomationCore.ThinkAutomation.clsAccount

                NewAccount.Name = strAccount
                NewAccount.IMAPAccount = strIMAPAccount
                NewAccount.DropBoxToAddress = strIMAPAccount
                NewAccount.DropBoxSubject = strDropBoxSubject
                NewAccount.IMAPServer = strIMAPServer
                NewAccount.IMAPFolder = strIMAPFolder
                NewAccount.IMAPMoveToFolder = strIMAPMoveToFolder
                NewAccount.CheckEveryMins = intCheckEveryMins


                ' add the new account to the collection and save it on the server
                If ThinkAutomationCore.ThinkAutomation.Accounts.AddAndSave(NewAccount) Then
                    Dim SelectedAccount As New ThinkAutomation.clsAccount
                    For Each ThisAccount In ThinkAutomation.Accounts
                        If ThisAccount.Name = strAccount Then

                            SelectedAccount = ThisAccount
                            SelectedAccount.IMAPAccount = strIMAPAccount
                            SelectedAccount.DropBoxToAddress = strIMAPAccount
                            SelectedAccount.DropBoxSubject = "Message From ThinkAutomation Client"
                            SelectedAccount.IMAPServer = strIMAPServer
                            SelectedAccount.IMAPFolder = strIMAPFolder
                            SelectedAccount.IMAPMoveToFolder = strIMAPMoveToFolder
                            SelectedAccount.IMAPPort = 993
                            SelectedAccount.CheckEveryMins = intCheckEveryMins


                            SelectedAccount.IMAP = vbTrue
                            SelectedAccount.IMAPLeaveMessageOn = vbTrue
                            SelectedAccount.IMAPMove = vbTrue
                            SelectedAccount.IMAPSSL = vbTrue
                            SelectedAccount.IMAPOAuth = vbTrue
                            SelectedAccount.IMAPNoPauseOnThrottle = vbTrue

                            Dim LastCheckedDate As DateTime = #2000-01-01 00:00:00#
                            Dim ServerLastEmailDate As DateTime = #2017-11-19 00:00:00#
                            SelectedAccount.LastChecked = LastCheckedDate
                            SelectedAccount.ServerLastEmail = ServerLastEmailDate
                            SelectedAccount.ServerPort = 110
                            SelectedAccount.StoreFull = vbTrue
                            SelectedAccount.MailServerPOP3 = vbTrue
                            SelectedAccount.DBMSGBody = vbTrue
                            SelectedAccount.HTTPFollow = vbTrue
                            SelectedAccount.HTTPClean = vbTrue
                            SelectedAccount.UnzipAttachments = vbTrue
                            SelectedAccount.Schedule = "030|YYYYYYY"
                            SelectedAccount.Dropbox = vbTrue
                            SelectedAccount.Archive = vbTrue
                            SelectedAccount.ArchiveDays = 3
                            SelectedAccount.MarkAsReadType = 2

                            SelectedAccount.Enabled = vbFalse
                            SelectedAccount.Save()
                        End If
                    Next
                    blnAccountAdded = True
                Else
                    MsgBox("Unable to Add Account: " + strAccount + " !!")
                End If
            End If
        Else
            MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
        Return blnAccountAdded

    End Function
    Private Sub AddTrigger(strAccount As String, strTriggerName As String)
        Dim ThisAccount As ThinkAutomation.clsAccount

        If strAccount.Length = 0 Or strTriggerName.Length = 0 Then Exit Sub

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then

            ' create a new trigger

            Dim NewTrigger As New ThinkAutomation.clsAccountTrigger
            NewTrigger.Name = strTriggerName
            NewTrigger.Enabled = vbTrue
            NewTrigger.HelperMessage = "#xD;#xA;#xD;#xA;"
            NewTrigger.HelperSubject = "Night Audit 10/26/17"
            NewTrigger.HelperTo = "lsimpkins@pacificahost.com,sloag@mydigitaloffice.ca"


            Dim IfBlock As New ThinkAutomation.clsActionLogical
            IfBlock.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.IfBlock

            ' build condition for if block
            Dim Condition As New ThinkAutomation.colConditionLines
            Dim MsgBodyIsNotBlank As New ThinkAutomation.clsConditionLine
            MsgBodyIsNotBlank.IfValue = "%msg_to%"
            MsgBodyIsNotBlank.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            MsgBodyIsNotBlank.Value = "lbbdt@mydigitaloffice.ca"
            Condition.Add(MsgBodyIsNotBlank)

            Dim SubjectContainsTest As New ThinkAutomation.clsConditionLine
            SubjectContainsTest.IsOr = True
            SubjectContainsTest.IfValue = "%msg_from%"
            SubjectContainsTest.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            SubjectContainsTest.Value = "infoDelivery@mi.mymicros.net"
            Condition.Add(SubjectContainsTest)

            'MsgBox("Adding Trigger Condition AS: " + Condition.XML)
            NewTrigger.Condition = Condition.XML
            'MsgBox("Added Trigger Condition AS: " + NewTrigger.Condition)

            'IfBlock.Condition = Condition.XML


            'IfBlock.Condition = "<Lines><Line>|%msg_to%|?|lbbdt@mydigitaloffice.ca</Line><Line>Or|%msg_from%|?|infoDelivery@mi.mymicros.net</Line></Lines>"
            Dim EndIfBlock As New ThinkAutomation.clsActionLogical
            EndIfBlock.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.EndIfBlock


            ' add the actions in the order we want them processed

            Dim IfBlock1 As New ThinkAutomation.clsActionLogical
            IfBlock1.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.IfBlock

            ' build condition for if block
            Dim Condition1 As New ThinkAutomation.colConditionLines
            Dim MsgBodyIsNotBlank1 As New ThinkAutomation.clsConditionLine
            MsgBodyIsNotBlank1.IfValue = "%msg_attachments%"
            MsgBodyIsNotBlank1.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            MsgBodyIsNotBlank1.Value = "xls"
            Condition1.Add(MsgBodyIsNotBlank1)

            'Dim SubjectContainsTest1 As New ThinkAutomation.clsConditionLine
            'SubjectContainsTest1.IsOr = True
            'SubjectContainsTest1.IfValue = "%msg_from%"
            'SubjectContainsTest1.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            'SubjectContainsTest1.Value = "infoDelivery@mi.mymicros.net"
            'Condition1.Add(SubjectContainsTest1)

            IfBlock1.Condition = Condition1.XML


            'IfBlock1.Condition = "<Lines><Line>|%msg_to%|?|lbbdt@mydigitaloffice.ca</Line><Line>Or|%msg_from%|?|infoDelivery@mi.mymicros.net</Line></Lines>"

            Dim EndIfBlock1 As New ThinkAutomation.clsActionLogical
            EndIfBlock1.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.EndIfBlock

            Dim AttachmentClass As New ThinkAutomation.clsActionAttach
            AttachmentClass.Condition = Condition1.XML
            AttachmentClass.Mask = "JFKQA*"
            AttachmentClass.SaveTo = "F:\TA - Attachments - MYP\3874\1170\OnQ\ForcastRoomRevenue"
            AttachmentClass.Overwrite = vbTrue

            ' add the actions in the order we want them processed
            NewTrigger.ActionList.Add(AttachmentClass)


            ' add the trigger to the selected account's collection and save it on the server
            Dim SelectedAccount As New ThinkAutomation.clsAccount
            For Each ThisAccount In ThinkAutomation.Accounts
                If ThisAccount.Name = strAccount Then
                    SelectedAccount = ThisAccount
                End If
            Next
            If SelectedAccount.Triggers.AddAndSave(NewTrigger) Then
                MsgBox("Trigger Added")
            End If
        Else
            MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub

    Private Sub AddMappTrigger(strAccount As String)
        If strAccount.Length = 0 Then Exit Sub

        Dim intAccountId As Integer
        Dim strIMAPAccount As String
        Dim strDropBoxSubject As String
        Dim strIMAPServer As String
        Dim strIMAPFolder As String
        Dim strIMAPMoveToFolder As String
        Dim intCheckEveryMins As Integer

        Dim connetionString As String
        Dim connection, UpdateConnection As SqlConnection
        Dim command, commandUpdate As SqlCommand
        Dim sql, UpdateSql As String

        connetionString = "Data Source=169.53.175.117;UID=exceluser;PWD=c4Ua8qPNJ4X804r;DATABASE=Perspective"
        sql = "select * from tatriggers where hastaaccountsetup = 0 and AccountName = '" + strAccount + "'"

        connection = New SqlConnection(connetionString)
        Try
            connection.Open()
            command = New SqlCommand(sql, connection)
            command.ExecuteNonQuery()
            'Create a Recordset object, populate it with the query results
            Dim objDR As SqlDataReader = command.ExecuteReader()

            If Not (objDR Is Nothing) AndAlso objDR.HasRows Then
                While objDR.Read
                    intAccountId = objDR("AccountId")
                    strIMAPAccount = objDR("IMAPAccount").ToString
                    strDropBoxSubject = objDR("DropBoxSubject").ToString
                    strIMAPServer = objDR("IMAPServer").ToString
                    strIMAPFolder = objDR("IMAPFolder").ToString
                    strIMAPMoveToFolder = objDR("IMAPMoveToFolder").ToString
                    intCheckEveryMins = objDR("CheckEveryMins")

                    If CreateTAServerAccount(strAccount, strIMAPAccount, strDropBoxSubject, strIMAPServer, strIMAPFolder, strIMAPMoveToFolder, intCheckEveryMins) Then
                        UpdateConnection = New SqlConnection(connetionString)
                        UpdateConnection.Open()
                        UpdateSql = "update taaccounts set HasTAAccountSetUp = 1 where AccountId = " + CStr(intAccountId)
                        commandUpdate = New SqlCommand(UpdateSql, UpdateConnection)
                        commandUpdate.ExecuteNonQuery()
                        commandUpdate.Dispose()
                        UpdateConnection.Close()
                        MsgBox("Added the Account: " + strAccount + " to the TA Server !!")
                    Else
                        MsgBox("Unable to Add the Account: " + strAccount + " to the TA Server !!")
                    End If
                End While
            Else
                MsgBox("Unable to Find the Given Account in the Database" + vbNewLine + "The Account Name provided is either incorrect or it has already been Added to the TA Server !!")
            End If

            command.Dispose()
            connection.Close()
        Catch ex As Exception
            MsgBox("Perspective Database ERROR! Cause: " + ex.Message)
        End Try
    End Sub
    Private Sub CreateTAServerAccountTrigger(strAccount As String, strTriggerName As String)
        Dim ThisAccount As ThinkAutomation.clsAccount

        If strAccount.Length = 0 Or strTriggerName.Length = 0 Then Exit Sub

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then

            ' create a new trigger

            Dim NewTrigger As New ThinkAutomation.clsAccountTrigger
            NewTrigger.Name = strTriggerName
            NewTrigger.Enabled = vbTrue
            NewTrigger.HelperMessage = "#xD;#xA;#xD;#xA;"
            NewTrigger.HelperSubject = "Night Audit 10/26/17"
            NewTrigger.HelperTo = "lsimpkins@pacificahost.com,sloag@mydigitaloffice.ca"


            Dim IfBlock As New ThinkAutomation.clsActionLogical
            IfBlock.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.IfBlock

            ' build condition for if block
            Dim Condition As New ThinkAutomation.colConditionLines
            Dim MsgBodyIsNotBlank As New ThinkAutomation.clsConditionLine
            MsgBodyIsNotBlank.IfValue = "%msg_to%"
            MsgBodyIsNotBlank.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            MsgBodyIsNotBlank.Value = "lbbdt@mydigitaloffice.ca"
            Condition.Add(MsgBodyIsNotBlank)

            Dim SubjectContainsTest As New ThinkAutomation.clsConditionLine
            SubjectContainsTest.IsOr = True
            SubjectContainsTest.IfValue = "%msg_from%"
            SubjectContainsTest.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            SubjectContainsTest.Value = "infoDelivery@mi.mymicros.net"
            Condition.Add(SubjectContainsTest)


            NewTrigger.Condition = Condition.XML



            ' add the actions in the order we want them processed

            Dim IfBlock1 As New ThinkAutomation.clsActionLogical
            IfBlock1.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.IfBlock

            ' build condition for if block
            Dim Condition1 As New ThinkAutomation.colConditionLines
            Dim MsgBodyIsNotBlank1 As New ThinkAutomation.clsConditionLine
            MsgBodyIsNotBlank1.IfValue = "%msg_attachments%"
            MsgBodyIsNotBlank1.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            MsgBodyIsNotBlank1.Value = "xls"
            Condition1.Add(MsgBodyIsNotBlank1)

            'Dim SubjectContainsTest1 As New ThinkAutomation.clsConditionLine
            'SubjectContainsTest1.IsOr = True
            'SubjectContainsTest1.IfValue = "%msg_from%"
            'SubjectContainsTest1.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            'SubjectContainsTest1.Value = "infoDelivery@mi.mymicros.net"
            'Condition1.Add(SubjectContainsTest1)

            IfBlock1.Condition = Condition1.XML


            'IfBlock1.Condition = "<Lines><Line>|%msg_to%|?|lbbdt@mydigitaloffice.ca</Line><Line>Or|%msg_from%|?|infoDelivery@mi.mymicros.net</Line></Lines>"

            Dim EndIfBlock1 As New ThinkAutomation.clsActionLogical
            EndIfBlock1.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.EndIfBlock

            Dim AttachmentClass As New ThinkAutomation.clsActionAttach
            AttachmentClass.Condition = Condition1.XML
            AttachmentClass.Mask = "JFKQA*"
            AttachmentClass.SaveTo = "F:\TA - Attachments - MYP\3874\1170\OnQ\ForcastRoomRevenue"
            AttachmentClass.Overwrite = vbTrue

            ' add the actions in the order we want them processed
            NewTrigger.ActionList.Add(AttachmentClass)


            ' add the trigger to the selected account's collection and save it on the server
            Dim SelectedAccount As New ThinkAutomation.clsAccount
            For Each ThisAccount In ThinkAutomation.Accounts
                If ThisAccount.Name = strAccount Then
                    SelectedAccount = ThisAccount
                End If
            Next
            If SelectedAccount.Triggers.AddAndSave(NewTrigger) Then
                MsgBox("Trigger Added")
            End If
        Else
            MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub


    Public Sub ListAllAccounts()
        Dim ThisAccount As ThinkAutomation.clsAccount

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then

            For Each ThisAccount In ThinkAutomation.Accounts ' Accounts collection will contain all ThinkAutomation accounts after login
                ' Add each account to the ListBox
                ListAccounts.Items.Add(ThisAccount.Name)
            Next
        Else
            MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub

    Public Sub ListAccountTriggers(AccountName As String)
        Dim ThisAccount As ThinkAutomation.clsAccount
        Dim ThisTrigger As ThinkAutomation.clsAccountTrigger
        Dim blnAccountFound As Boolean

        blnAccountFound = False

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
            For Each ThisAccount In ThinkAutomation.Accounts ' Accounts collection will contain all ThinkAutomation accounts after login

                If ThisAccount.Name = AccountName Then
                    blnAccountFound = True
                    LblAllTriggers.Text = "All Trigger for the Account: " + AccountName
                    ListTriggers.Items.Add(AccountName)
                    ThisAccount.Triggers.Load()
                    ' display all triggers for this account
                    For Each ThisTrigger In ThisAccount.Triggers
                        ListTriggers.Items.Add("     " & ThisTrigger.Name)
                    Next
                End If

            Next

            If Not blnAccountFound Then
                MsgBox("Could not find the Account. " & ThinkAutomation.SystemErrorLast)
            End If

        Else
            MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub


    Public Sub ListTriggerActions(AccountName As String, TriggerName As String)
        Dim ThisAccount As ThinkAutomation.clsAccount
        Dim ThisTrigger As ThinkAutomation.clsAccountTrigger
        Dim ThisAction As Object
        Dim ActionString As String

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
            For Each ThisAccount In ThinkAutomation.Accounts ' Accounts collection will contain all ThinkAutomation accounts after login

                If ThisAccount.Name = AccountName Then
                    ThisAccount.Triggers.Load()
                    ' display all triggers for this account
                    LblAllActions.Text = "All Actions for Trigger: " + TriggerName
                    For Each ThisTrigger In ThisAccount.Triggers
                        If ThisTrigger.Name = TriggerName Then
                            For Each ThisAction In ThisTrigger.ActionList

                                ActionString = ThisAction.ActionDescription
                                ListActions.Items.Add(ActionString)

                            Next
                            ' fill the webbrowser with the Trigger HTML view
                            WebBrowser1.DocumentText = ThisTrigger.HTML
                        End If
                    Next
                End If

            Next

        Else
            MsgBox("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub
    Private Sub BtnListAccounts_Click(sender As Object, e As EventArgs) Handles BtnListAccounts.Click
        ListAllAccounts()
    End Sub
    Private Sub BtnAddAccount_Click(sender As Object, e As EventArgs) Handles BtnAddAccount.Click
        Dim Name As String = InputBox("Enter The Account Name", "ThinkAutomation")

        If Name.Length = 0 Then Exit Sub Else AddAccount(Name)
    End Sub
    Private Sub BtnAddTrigger_Click(sender As Object, e As EventArgs) Handles BtnAddTrigger.Click
        If ListAccounts.SelectedItems.Count = 0 Then
            MsgBox("Please select the Account For which you want to add a Trigger!")
        Else
            Dim AccountName As String = LTrim(ListAccounts.SelectedItems(0))
            If AccountName.Length = 0 Then Exit Sub
            Dim TriggerName As String = InputBox("Enter the Name for the Trigger", "ThinkAutomation Server")
            If TriggerName.Length = 0 Then
                MsgBox("Name for the Trigger is Required !")
                Exit Sub
            Else
                MsgBox("Will be adding the Trigger: " + TriggerName + ", for the Account: " + AccountName)
                AddTrigger(AccountName, TriggerName)
            End If
        End If
    End Sub
    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        Application.Exit()
    End Sub
End Class
