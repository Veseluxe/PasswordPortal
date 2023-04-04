Imports System

Module Program

    Dim Password As String
    Dim PassCheckUp As Boolean
    Dim PassCheckNum As Boolean
    Dim PassCheckLen As Boolean
    Dim PassCheckSp As Boolean
    Dim ExPassCheck As Boolean

    Sub List()

        Console.WriteLine("1 | Enter a new password")
        Console.WriteLine("2 | Check their password")
        Console.WriteLine("3 | Change their password")
        Console.WriteLine("4 | Quit")
        Console.WriteLine()

    End Sub
    Sub Uppercase()

        PassCheckUp = False

        For Count = 0 To Password.Length - 1
            If Char.IsUpper(Password(Count)) Then
                PassCheckUp = True
            End If
        Next

        Call Number()

    End Sub
    Sub Number()

        PassCheckNum = False

        For Count = 0 To Password.Length - 1
            If Char.IsDigit(Password(Count)) Then
                PassCheckNum = True
            End If
        Next

        Call Length()

    End Sub
    Sub Length()

        If Password.Length >= 10 And Password.Length <= 20 Then
            PassCheckLen = True
        Else : PassCheckLen = False
        End If

        Call Spaces()

    End Sub
    Sub Spaces()

        If InStr(Password, " ") Then
            PassCheckSp = False
        Else
            PassCheckSp = True
        End If

        Call TextFile()

    End Sub
    Sub TextFile()

        If PassCheckUp = True And PassCheckNum = True And PassCheckLen = True And PassCheckSp = True Then
            FileOpen(1, "Password.txt", OpenMode.Output)
            PrintLine(1, Password)
            FileClose(1)

            Console.WriteLine("Password Saved!")

            ExPassCheck = True
        End If


    End Sub
    Sub ErrorList()

        If PassCheckUp = False Then Console.WriteLine("Password should contain atleast one uppercase letter")
        If PassCheckNum = False Then Console.WriteLine("Password should contain atleast one number")
        If PassCheckLen = False Then Console.WriteLine("Password should have 10-20 characters")
        If PassCheckSp = False Then Console.WriteLine("Password should not contain spaces")
        Console.WriteLine()

    End Sub

    Sub Main()

        Dim Num As Integer
        Dim ExPass As String

        Do
            Call List()

            Console.Write("Enter option = ")
            Num = Console.ReadLine
            Console.WriteLine()

            If Num = 1 Then

                If ExPassCheck = True Then
                    Console.WriteLine("Password is already set, to change password choose option 3")
                    Console.WriteLine()
                ElseIf ExPassCheck = False Then
                    Console.Write("Enter New Password = ")
                    Password = Console.ReadLine
                    Console.WriteLine()

                    Call Uppercase()
                    Call ErrorList()

                End If

            ElseIf Num = 2 Then

                If ExPassCheck = False Then
                    Console.WriteLine("Please assign a password first using option 1")
                    Console.WriteLine()
                ElseIf PassCheckLen = True Then
                    Console.Write("Enter Existing Password = ")
                    ExPass = Console.ReadLine
                    Console.WriteLine()

                    FileOpen(1, "Password.txt", OpenMode.Input)
                    If ExPass = LineInput(1) Then
                        Console.WriteLine("Your Password is CORRECT")
                        Console.WriteLine()
                    Else
                        Console.WriteLine("Your Password is INCORRECT")
                        Console.WriteLine()
                    End If
                    FileClose(1)
                End If

            ElseIf Num = 3 Then

                If ExPassCheck = False Then
                    Console.WriteLine("Please assign a password first using option 1")
                    Console.WriteLine()
                ElseIf PassCheckLen = True Then
                    Console.Write("Enter Existing Password = ")
                    ExPass = Console.ReadLine
                    Console.WriteLine()

                    FileOpen(1, "Password.txt", OpenMode.Input)
                    If ExPass = LineInput(1) Then
                        Console.WriteLine("Your Password is CORRECT")
                        Console.WriteLine()
                        Console.Write("Enter New Password = ")
                        Password = Console.ReadLine
                        Console.WriteLine()
                        FileClose(1)
                        Call Uppercase()
                        Call ErrorList()
                    Else
                        Console.WriteLine("Your Password is INCORRECT")
                        Console.WriteLine()
                    End If
                    FileClose(1)
                End If

            End If

        Loop Until Num = 4

    End Sub

End Module