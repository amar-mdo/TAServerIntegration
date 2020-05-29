﻿Imports System
Imports System.Net
Imports ThinkAutomationCore
Imports System.IO
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        Application.Exit()
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

        Dim fPath As String = "C:\MDO\ThinkAutomationTest\"
        'Dim fFileName As String = AccountName + "-" + Now.ToString("s") + ".txt"
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

    ' For Adding Action
    Private Sub AddAction(strAccount As String)
        Dim ThisAccount As ThinkAutomation.clsAccount

        If strAccount.Length = 0 Then Exit Sub

        ' login to ThinkAutomation on 127.0.0.1:8855
        If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then
            ' logged in OK
            MsgBox("Logged In")

            Dim Name As String = InputBox("Enter The Action Condition", "ThinkAutomation")

            If Name.Length = 0 Then Exit Sub

            ' create a new trigger

            Dim NewAction As New ThinkAutomation.clsAccountTrigger
            NewAction.Name = Name
            NewAction.CheckBodyAlso = True
            NewAction.SubjectLineLIKE = "*TEST*"

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

            NewAction.FieldsToProcess.Add(Field)
            NewAction.FieldsToProcess.Add(Field2)


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

            NewAction.ActionList.Add(Comment)
            NewAction.ActionList.Add(IfBlock)
            NewAction.ActionList.Add(Popup)
            NewAction.ActionList.Add(EndIfBlock)


            ' add the trigger to the selected account's collection and save it on the server
            Dim SelectedAccount As New ThinkAutomation.clsAccount
            For Each ThisAccount In ThinkAutomation.Accounts
                If ThisAccount.Name = strAccount Then
                    SelectedAccount = ThisAccount
                End If
            Next
            If SelectedAccount.Triggers.AddAndSave(NewAction) Then
                MsgBox("Trigger Added")
                ' add a message to test
                NewAction.AddMessageToProcess("stephen@parkersoft.co.uk", "test@test.com", "Test Subject", "Test Body" & vbCrLf & vbCrLf & "ID: 12345" & vbCrLf)
            End If
        Else
            Console.WriteLine("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
        End If
    End Sub
    Private Sub BtnAddAccount_Click(sender As Object, e As EventArgs) Handles BtnAddAccount.Click
        Dim Name As String = InputBox("Enter The Account Name", "ThinkAutomation")

        If Name.Length = 0 Then Exit Sub Else AddAccount(Name)
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

    Private Sub BtnAddTrigger_Click(sender As Object, e As EventArgs) Handles BtnAddTrigger.Click
        Dim Name As String = InputBox("Enter The Account Name", "ThinkAutomation")
        If Name.Length = 0 Then Exit Sub
        AddTrigger(Name)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' MsgBox("Hello, you can add Action here!")
        Dim Name As String = InputBox("Enter The Account Name", "ThinkAutomation")
        If Name.Length = 0 Then Exit Sub
        Dim NameTrigger As String = InputBox("Enter the Trigger Name", "ThinkAutomation")
        If Name.Length = 0 Then Exit Sub
        AddAction(Name)

    End Sub

End Class
