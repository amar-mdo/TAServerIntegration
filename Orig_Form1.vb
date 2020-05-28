Imports System
Imports System.Net
Imports ThinkAutomationCore
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        Application.Exit()
    End Sub
    Private Sub ListTriggers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListTriggers.SelectedIndexChanged

        'If SelectedAccount Is Nothing Then Exit Sub

        If ListTriggers.SelectedItems.Count = 0 Then Exit Sub



        Dim SelectedItem As ListViewItem = ListTriggers.SelectedItems(0)



        'SelectedTrigger = SelectedAccount.Triggers(SelectedItem.Name)



        ' fill the actions list

        ListActions.Items.Clear()

        'Dim Action As Object

        'For Each Action In SelectedTrigger.ActionList

        'Dim Item As New ListViewItem

        'Item.Text = Action.ClassName

        'Item.Checked = Action.Enabled

        'Item.SubItems.Add(Action.ActionDescription)

        'ListActions.Items.Add(Item)

        'Next



        ' fill the webbrowser with the Trigger HTML view

        ' WebBrowser1.DocumentText = SelectedTrigger.HTML



    End Sub
    Private Sub AddAccount(strAccount As String)
        If strAccount.Length = 0 Then Exit Sub
        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
            ' logged in OK
            MsgBox("Logged In")

            Dim NewAccount As New ThinkAutomationCore.ThinkAutomation.clsAccount

            NewAccount.Name = strAccount
            MsgBox("Adding Account: " & strAccount)
            ' add the new account to the collection and save it on the server

            If ThinkAutomationCore.ThinkAutomation.Accounts.AddAndSave(NewAccount) Then MsgBox("Account Added") Else MsgBox("Unable to Add Account!")

        Else
            Console.WriteLine("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub
    Private Sub AddTrigger(strAccount As String)
        Dim ThisAccount As ThinkAutomation.clsAccount

        If strAccount.Length = 0 Then Exit Sub

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
            ' logged in OK
            MsgBox("Logged In")

            Dim Name As String = InputBox("Enter The Trigger Name", "ThinkAutomation")

            If Name.Length = 0 Then Exit Sub

            ' create a new trigger

            Dim NewTrigger As New ThinkAutomation.clsAccountTrigger
            NewTrigger.Name = Name
            NewTrigger.CheckBodyAlso = True
            NewTrigger.SubjectLineLIKE = "*TEST*"

            ' create some field extraction

            Dim Field As New ThinkAutomation.clsTriggerField
            Field.Name = "Subject"
            Field.ExtractBuiltIn = True
            Field.ExtractBuiltInField = "%msg_subject%"


            Dim Field2 As New ThinkAutomation.clsTriggerField
            Field2.Name = "TestField"
            Field2.LookFor = "ID"
            Field2.ThenLookFor = ":"
            Field2.Until_EndOfLine = True

            ' add the fields to the fieldstoprocess collection

            NewTrigger.FieldsToProcess.Add(Field)
            NewTrigger.FieldsToProcess.Add(Field2)


            ' create some default actions
            Dim Comment As New ThinkAutomation.clsActionComment
            Comment.Comment = "Actions Go Here"
            Dim Popup As New ThinkAutomation.clsActionPopup
            Popup.PopupMessage = "%Subject% %TestField%"


            Dim IfBlock As New ThinkAutomation.clsActionLogical
            IfBlock.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.IfBlock

            ' build condition for if block

            Dim Condition As New ThinkAutomation.colConditionLines
            Dim MsgBodyIsNotBlank As New ThinkAutomation.clsConditionLine
            MsgBodyIsNotBlank.IfValue = "%msg_body%"
            MsgBodyIsNotBlank.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.IsNotBlank

            Condition.Add(MsgBodyIsNotBlank)
            Dim SubjectContainsTest As New ThinkAutomation.clsConditionLine
            SubjectContainsTest.IfValue = "%Subject%"
            SubjectContainsTest.IsAnd = True

            SubjectContainsTest.IsOperator = ThinkAutomation.clsConditionLine.IsOperatorType.Contains
            SubjectContainsTest.Value = "test"
            Condition.Add(SubjectContainsTest)


            IfBlock.Condition = Condition.XML
            Dim EndIfBlock As New ThinkAutomation.clsActionLogical
            EndIfBlock.LogicalType = ThinkAutomation.clsActionLogical.LogicalTypeNo.EndIfBlock

            ' add the actions in the order we want them processed

            NewTrigger.ActionList.Add(Comment)
            NewTrigger.ActionList.Add(IfBlock)
            NewTrigger.ActionList.Add(Popup)
            NewTrigger.ActionList.Add(EndIfBlock)


            ' add the trigger to the selected account's collection and save it on the server
            Dim SelectedAccount As New ThinkAutomation.clsAccount
            For Each ThisAccount In ThinkAutomation.Accounts
                If ThisAccount.Name = strAccount Then
                    SelectedAccount = ThisAccount
                End If
            Next
            If SelectedAccount.Triggers.AddAndSave(NewTrigger) Then
                MsgBox("Trigger Added")
                ' add a message to test
                NewTrigger.AddMessageToProcess("stephen@parkersoft.co.uk", "test@test.com", "Test Subject", "Test Body" & vbCrLf & vbCrLf & "ID: 12345" & vbCrLf)
            End If
        Else
            Console.WriteLine("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub
    Private Sub BtnAddAccount_Click(sender As Object, e As EventArgs) Handles BtnAddAccount.Click
        Dim Name As String = InputBox("Enter The Account Name", "ThinkAutomation")

        If Name.Length = 0 Then Exit Sub Else AddAccount(Name)
        'If Name.Length = 0 Then Exit Sub Else MsgBox(Name)

        'Dim NewAccount As New ThinkAutomationCore.ThinkAutomation.clsAccount

        'NewAccount.Name = Name
        'MsgBox("Adding Account: " & Name)
        ' add the new account to the collection and save it on the server

        'If ThinkAutomationCore.ThinkAutomation.Accounts.AddAndSave(NewAccount) Then MsgBox("Account Added") Else MsgBox("Unable to Add Account!")

    End Sub

    Private Sub BtnTestTAServer_Click(sender As Object, e As EventArgs) Handles BtnTestTAServer.Click
        Dim MyTAServer1 As TAServerIntegration.MyTAServer.GetTAServer
        MsgBox("Initiating TA Server")
        MyTAServer1 = New MyTAServer.GetTAServer
        MsgBox("Initiated TA Server")

        MyTAServer1.examineTimeZone()
    End Sub
    Public Sub ListAccountsAndTriggers()
        Dim ThisAccount As ThinkAutomation.clsAccount
        Dim ThisTrigger As ThinkAutomation.clsAccountTrigger

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
            For Each ThisAccount In ThinkAutomation.Accounts ' Accounts collection will contain all ThinkAutomation accounts after login

                'Console.WriteLine("Account: " & ThisAccount.Name)
                ListAccounts.Items.Add(ThisAccount.Name)
                ' get the Triggers from the server since these are not automatically loaded at Login
                ThisAccount.Triggers.Load()
                ' display all triggers for this account

                For Each ThisTrigger In ThisAccount.Triggers
                    'Console.WriteLine("Trigger: " & ThisTrigger.Name & " Has " & ThisTrigger.ActionList.Count.ToString & " Actions")
                    ListAccounts.Items.Add("     " & "Trigger: " & ThisTrigger.Name & " Has " & ThisTrigger.ActionList.Count.ToString & " Actions")
                Next

            Next

        Else
            Console.WriteLine("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub
    Private Sub BtnListAccounts_Click(sender As Object, e As EventArgs) Handles BtnListAccounts.Click
        ListAccountsAndTriggers()
    End Sub

    Private Sub BtnAddTrigger_Click(sender As Object, e As EventArgs) Handles BtnAddTrigger.Click
        Dim Name As String = InputBox("Enter The Account Name", "ThinkAutomation")
        If Name.Length = 0 Then Exit Sub
        AddTrigger(Name)
    End Sub
End Class
