﻿Imports System.Collections.ObjectModel

Module Module1
    Enum HashType
        null = 0
        md5 = 1
        sha160 = 2
        sha256 = 3
        sha384 = 4
        sha512 = 5
        ripemd160 = 6
    End Enum

    Private Function ComputeHash(filename As String, Optional mode As HashType = HashType.sha160) As String
        Dim byteOutput As Byte()

        Using fileStreamObject As New IO.FileStream(filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 10 * 1048576, IO.FileOptions.SequentialScan)
            If mode = HashType.md5 Then
                Using hashEngine As New Security.Cryptography.MD5CryptoServiceProvider
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            ElseIf mode = HashType.sha160 Then
                Using hashEngine As New Security.Cryptography.SHA1CryptoServiceProvider
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            ElseIf mode = HashType.sha256 Then
                Using hashEngine As New Security.Cryptography.SHA256CryptoServiceProvider
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            ElseIf mode = HashType.sha384 Then
                Using hashEngine As New Security.Cryptography.SHA384CryptoServiceProvider
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            ElseIf mode = HashType.sha512 Then
                Using hashEngine As New Security.Cryptography.SHA512CryptoServiceProvider
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            ElseIf mode = HashType.ripemd160 Then
                Using hashEngine As New Security.Cryptography.RIPEMD160Managed
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            Else
                Using hashEngine As New Security.Cryptography.SHA1CryptoServiceProvider
                    byteOutput = hashEngine.ComputeHash(fileStreamObject)
                End Using
            End If
        End Using

        Return BitConverter.ToString(byteOutput).ToLower().Replace("-", "")
    End Function

    Private Function ParseArguments(args As ReadOnlyCollection(Of String)) As Dictionary(Of String, Object)
        Dim parsedArguments As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim strValue As String

        For Each strArgument As String In args
            If strArgument.StartsWith("--") Then
                Dim splitArg As String() = strArgument.Substring(2).Split(New Char() {"="c}, 2)
                Dim key As String = splitArg(0)

                If splitArg.Length = 2 Then
                    ' Argument with a value
                    strValue = splitArg(1)
                    parsedArguments(key) = strValue
                Else
                    ' Boolean flag
                    parsedArguments(key) = True
                End If
            Else
                Console.WriteLine($"Unrecognized argument format: {strArgument}")
            End If
        Next

        Return parsedArguments
    End Function

    ''' <summary>
    ''' Writes a line of text to the console with color.
    ''' </summary>
    ''' <param name="strStringToWriteToTheConsole">The text that's written to the console.</param>
    ''' <param name="boolNewLine">A Boolean value. This tells the function if it should include a new line after printing the text to the console.</param>
    ''' <param name="color">The color in which the color of the text is written to the console.</param>
    Private Sub ColoredConsoleLineWriter(strStringToWriteToTheConsole As String, boolNewLine As Boolean, Optional color As ConsoleColor = ConsoleColor.Green)
        Console.ForegroundColor = color

        If boolNewLine Then
            Console.WriteLine(strStringToWriteToTheConsole)
        Else
            Console.Write(strStringToWriteToTheConsole)
        End If

        Console.ResetColor()
    End Sub

    Sub Main()
        Dim parsedArguments As Dictionary(Of String, Object) = ParseArguments(My.Application.CommandLineArgs)
        Dim boolAsterisk As Boolean = parsedArguments.ContainsKey("asterisk")
        Dim boolFullPath As Boolean = parsedArguments.ContainsKey("fullpath")

        If parsedArguments.ContainsKey("hash") And parsedArguments.ContainsKey("file") Then
            Dim strHashMode As String = parsedArguments("hash").ToString.Trim
            Dim strFileName As String = parsedArguments("file")
            Dim htHashType As HashType = HashType.null

            If strHashMode.Equals("md5", StringComparison.CurrentCulture) Then
                htHashType = HashType.md5
            ElseIf strHashMode.Equals("sha1", StringComparison.CurrentCulture) Or strHashMode.Equals("sha160", StringComparison.CurrentCulture) Then
                htHashType = HashType.sha160
            ElseIf strHashMode.Equals("sha256", StringComparison.CurrentCulture) Then
                htHashType = HashType.sha256
            ElseIf strHashMode.Equals("sha384", StringComparison.CurrentCulture) Then
                htHashType = HashType.sha384
            ElseIf strHashMode.Equals("sha512", StringComparison.CurrentCulture) Then
                htHashType = HashType.sha512
            ElseIf strHashMode.Equals("ripemd160", StringComparison.CurrentCulture) Then
                htHashType = HashType.ripemd160
            Else
                ColoredConsoleLineWriter("ERROR:", False, ConsoleColor.Red)
                Console.WriteLine(" Invalid hash mode.")
                Exit Sub
            End If

            If Not IO.Path.IsPathRooted(strFileName) Then strFileName = IO.Path.Combine(Environment.CurrentDirectory, strFileName)

            If IO.File.Exists(strFileName) Then
                Dim fileInfo As New IO.FileInfo(strFileName)
                Dim strHashString As String

                If boolAsterisk Then
                    strHashString = $"{ComputeHash(strFileName, htHashType)} *{If(boolFullPath, fileInfo.FullName, fileInfo.Name)}"
                Else
                    strHashString = $"{ComputeHash(strFileName, htHashType)} {If(boolFullPath, fileInfo.FullName, fileInfo.Name)}"
                End If

                If parsedArguments.ContainsKey("output") AndAlso Not String.IsNullOrWhiteSpace(parsedArguments.ContainsKey("output").ToString) Then
                    Dim strOutputFile As String = parsedArguments("output")

                    If Not IO.Path.IsPathRooted(strOutputFile) Then strOutputFile = IO.Path.Combine(Environment.CurrentDirectory, strOutputFile)

                    IO.File.WriteAllText(strOutputFile, strHashString)

                    ColoredConsoleLineWriter("--== Checksum Command Line ==--", True, ConsoleColor.Green)
                    ColoredConsoleLineWriter("INFO:", False, ConsoleColor.Green)
                    Console.WriteLine($" The hash data has been written to file ""{strOutputFile}"".")
                Else
                    Console.WriteLine(strHashString)
                End If
            Else
                ColoredConsoleLineWriter("ERROR:", False, ConsoleColor.Red)
                Console.WriteLine($" The file ""{strFileName}"" doesn't exist.")
                Exit Sub
            End If
        Else
            ColoredConsoleLineWriter("--== Checksum Command Line ==--", True, ConsoleColor.Green)
            Console.WriteLine("You must provide some input for this program to work. For example...")
            Console.WriteLine()
            Console.WriteLine("checksum.exe --hash=[mode] --file=[filename] --asterisk --output=[outputFilePath] --fullpath")
            Console.WriteLine()
            Console.WriteLine("Your hash modes are as follows... md5 sha1/sha160 sha256 sha384 sha512 ripemd160")
            Console.WriteLine()
            Console.WriteLine("The asterisk flag tells the program if it should include an asterisk before the file name. For example...")
            Console.WriteLine("[checksum] *[file name]")
            Console.WriteLine("If you don't want an asterisk, don't include the flag.")
            Console.WriteLine()
            Console.WriteLine("The --output variable allows you to write the hash data to a specified file.")
            Console.WriteLine("The --fullpath tells the program if it should include a full path to the hashed file or just the name of the file without the full path.")
        End If

        If Debugger.IsAttached Then Threading.Thread.Sleep(4000)
    End Sub
End Module