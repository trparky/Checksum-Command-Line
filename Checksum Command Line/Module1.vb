Imports System.Text.RegularExpressions

Module Module1
    Enum HashType
        null = 0
        md5 = 1
        sha160 = 2
        sha256 = 3
        sha384 = 4
        sha512 = 5
        ripemd160 = 6
        crc32 = 7
    End Enum

    Public Function ComputeHash(ByVal filename As String, Optional mode As HashType = HashType.sha160) As String
        Dim byteOutput As Byte()

        Using fileStreamObject As New IO.FileStream(filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 10 * 1048576, IO.FileOptions.SequentialScan)
            If mode = HashType.md5 Then
                Dim hashEngine As New Security.Cryptography.MD5CryptoServiceProvider
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            ElseIf mode = HashType.sha160 Then
                Dim hashEngine As New Security.Cryptography.SHA1CryptoServiceProvider
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            ElseIf mode = HashType.sha256 Then
                Dim hashEngine As New Security.Cryptography.SHA256CryptoServiceProvider
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            ElseIf mode = HashType.sha384 Then
                Dim hashEngine As New Security.Cryptography.SHA384CryptoServiceProvider
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            ElseIf mode = HashType.sha512 Then
                Dim hashEngine As New Security.Cryptography.SHA512CryptoServiceProvider
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            ElseIf mode = HashType.ripemd160 Then
                Dim hashEngine As New Security.Cryptography.RIPEMD160Managed
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            Else
                Dim hashEngine As New Security.Cryptography.SHA1CryptoServiceProvider
                byteOutput = hashEngine.ComputeHash(fileStreamObject)
                hashEngine.Dispose()
            End If
        End Using

        Return BitConverter.ToString(byteOutput).ToLower().Replace("-", "")
    End Function

    Sub Main()
        Dim commandLineEntry As String
        Dim htHashType As HashType = HashType.null
        Dim strFileToBeHashed As String = Nothing
        Dim boolAsterisk As Boolean = False

        For i = 0 To Environment.GetCommandLineArgs().Count - 1
            commandLineEntry = Environment.GetCommandLineArgs(i)

            If Regex.IsMatch(commandLineEntry, "\A-hash", RegexOptions.IgnoreCase) Or Regex.IsMatch(commandLineEntry, "\A-h", RegexOptions.IgnoreCase) Then
                commandLineEntry = Regex.Replace(commandLineEntry, "-(?:hash|h)=", "", RegexOptions.IgnoreCase)

                Select Case commandLineEntry.ToLower.Trim
                    Case "md5"
                        htHashType = HashType.md5
                    Case "sha1"
                        htHashType = HashType.sha160
                    Case "sha160"
                        htHashType = HashType.sha160
                    Case "sha256"
                        htHashType = HashType.sha256
                    Case "sha384"
                        htHashType = HashType.sha384
                    Case "sha512"
                        htHashType = HashType.sha512
                    Case "ripemd160"
                        htHashType = HashType.ripemd160
                    Case Else
                        Console.WriteLine("ERROR: Invalid hash mode")
                        Exit Sub
                End Select
            ElseIf Regex.IsMatch(commandLineEntry, "\A-file", RegexOptions.IgnoreCase) Then
                commandLineEntry = Regex.Replace(commandLineEntry, "-file=", "", RegexOptions.IgnoreCase)
                strFileToBeHashed = commandLineEntry
            ElseIf commandLineEntry.Equals("-asterisk", StringComparison.OrdinalIgnoreCase) Then
                boolAsterisk = True
            End If
        Next

        If htHashType = HashType.null Or strFileToBeHashed = Nothing Then
            Console.WriteLine("You must provide some input for this program to work. For example...")
            Console.WriteLine()
            Console.WriteLine("checksum.exe -hash=[mode] -file=[filename] -asterisk")
            Console.WriteLine()
            Console.WriteLine("Your hash modes are as follows... crc/crc32 md5 sha1/sha160 sha256 sha384 sha512 ripemd160")
            Console.WriteLine()
            Console.WriteLine("The asterisk flag tells the program if it should include an asterisk before the file name. For example...")
            Console.WriteLine("[checksum] *[file name]")
            Console.WriteLine("If you don't want an asterisk, don't include the flag.")
        Else
            Dim fileInfo As New IO.FileInfo(strFileToBeHashed)

            If fileInfo.Exists() Then
                If boolAsterisk Then
                    Console.WriteLine(ComputeHash(strFileToBeHashed, htHashType) & " *" & fileInfo.Name)
                Else
                    Console.WriteLine(ComputeHash(strFileToBeHashed, htHashType) & " " & fileInfo.Name)
                End If
            Else
                Console.WriteLine("ERROR: The file """ & strFileToBeHashed & """ doesn't exist.")
            End If
        End If

        If Debugger.IsAttached Then Threading.Thread.Sleep(4000)
    End Sub
End Module