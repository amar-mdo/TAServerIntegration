Imports System

Imports System.Net

Imports ThinkAutomationCore



Namespace MyTAServer

    Class GetTAServer

        Shared Sub Main(ByVal args As String())
        End Sub

        Public Sub ConnectTAServer()
            Dim ThisAccount As ThinkAutomation.clsAccount
            Dim ThisTrigger As ThinkAutomation.clsAccountTrigger

            ' login to ThinkAutomation on 127.0.0.1:8855

            If ThinkAutomation.Server.Login("127.0.0.1:8855", "Admin", "") Then

                ' logged in OK
                MsgBox("Logged In")

                For Each ThisAccount In ThinkAutomation.Accounts ' Accounts collection will contain all ThinkAutomation accounts after login

                    Console.WriteLine("Account: " & ThisAccount.Name)

                    ' get the Triggers from the server since these are not automatically loaded at Login

                    ThisAccount.Triggers.Load()

                    ' display all triggers for this account

                    For Each ThisTrigger In ThisAccount.Triggers
                        Console.WriteLine("Trigger: " & ThisTrigger.Name & " Has " & ThisTrigger.ActionList.Count.ToString & " Actions")
                    Next

                Next
            Else
                Console.WriteLine("Could not login. Is ThinkAutomation Server running? " & ThinkAutomation.SystemErrorLast)
            End If
        End Sub
        Public Sub examineTimeZone()
            Dim tz As System.TimeZone = System.TimeZone.CurrentTimeZone
            Dim s As String = "Current time zone is "
            s &= CStr(tz.GetUtcOffset(Now).Hours) & " hours and "
            s &= CStr(tz.GetUtcOffset(Now).Minutes) & " minutes "
            s &= "different from UTC (coordinated universal time)"
            s &= vbCrLf & "and is currently "
            If tz.IsDaylightSavingTime(Now) = False Then s &= "not "
            s &= "on ""summer time""."
            MsgBox(s)
        End Sub

    End Class

End Namespace
