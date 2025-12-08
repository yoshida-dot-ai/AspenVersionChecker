Option Strict On
Option Explicit On

Imports System.Windows.Forms

Public Class AspenVersionCheckerForm
    Private ReadOnly listVersionInfo As New Dictionary(Of Tuple(Of String, String), String)

    Private Sub AspenVersionCheckerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = False

        Try
            LoadVersionInfo("AspenVersionInfo.txt")
            GetAspenVersion()
        Catch ex As Exception
            MessageBox.Show(
                "An error has occurred:" & Environment.NewLine & ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        Finally
            Me.Close()
        End Try
    End Sub

    Private Sub LoadVersionInfo(fileVersionInfo As String)
        listVersionInfo.Clear()

        Dim pathVersionInfo As String = IO.Path.Combine(Application.StartupPath, fileVersionInfo)

        If Not IO.File.Exists(pathVersionInfo) Then
            Throw New IO.FileNotFoundException(
                "Required version info file 'AspenVersionInfo.txt' was not found." &
                Environment.NewLine &
                "Path: " & pathVersionInfo,
                pathVersionInfo)
        End If

        Dim fileExtensions As String() = {}

        Using sr As New IO.StreamReader(pathVersionInfo, System.Text.Encoding.ASCII)
            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine()
                If line Is Nothing Then Exit While

                Dim readLine As String = line.Trim()
                If readLine = "" Then Continue While

                ' Block of file extensions such as: [apw,apwz]
                If readLine.StartsWith("[") AndAlso readLine.EndsWith("]") Then
                    readLine = readLine.Trim("["c, "]"c).Replace(" ", "")
                    fileExtensions = readLine.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)

                    ' Pattern,Version mapping
                Else
                    Dim versionInfo As String() = readLine.Split(","c)
                    If versionInfo.Length > 1 Then
                        Dim pattern As String = versionInfo(0).Trim()
                        Dim version As String = versionInfo(1).Trim()

                        For Each ext As String In fileExtensions
                            Dim trimmedExt = ext.Trim()
                            If trimmedExt = "" Then Continue For

                            Dim key = Tuple.Create(trimmedExt, pattern)
                            If Not listVersionInfo.ContainsKey(key) Then
                                listVersionInfo.Add(key, version)
                            End If
                        Next
                    End If
                End If
            End While
        End Using
    End Sub

    Private Sub GetAspenVersion()
        Dim commandLineArgs As String() = Environment.GetCommandLineArgs()

        If commandLineArgs IsNot Nothing AndAlso commandLineArgs.Length > 2 Then
            MessageBox.Show(
            "Multiple files were provided. Please drag and drop only one file.",
            "Multiple Files Detected",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)
            Return
        End If

        If commandLineArgs Is Nothing OrElse commandLineArgs.Length <= 1 Then
            MessageBox.Show(
                "Drag and drop an Aspen file onto this application.",
                "No File Provided",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Return
        End If

        Dim fileName As String = commandLineArgs(1)

        If Not IO.File.Exists(fileName) Then
            MessageBox.Show(
                "The specified file does not exist:" & Environment.NewLine & fileName,
                "File Not Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
            Return
        End If

        Dim extension As String = IO.Path.GetExtension(fileName).Replace(".", "").ToLowerInvariant()
        Dim readLine As String = String.Empty
        Dim version As String = String.Empty

        Select Case extension

            Case "bkp", "acmf", "dynf", "ada", "cra", "auf", "edr"
                Using sr As New IO.StreamReader(fileName, System.Text.Encoding.ASCII)
                    Dim firstLine As String = sr.ReadLine()
                    Dim secondLine As String = sr.ReadLine()
                    readLine = firstLine & secondLine.Trim()
                    version = GetVersionFromList(extension, readLine, True)
                End Using

            Case "apwz"
                Using za As IO.Compression.ZipArchive = IO.Compression.ZipFile.OpenRead(fileName)
                    For Each entry As IO.Compression.ZipArchiveEntry In za.Entries
                        Select Case IO.Path.GetExtension(entry.Name).Replace(".", "").ToLowerInvariant()
                            Case "bkp"
                                Using sr As New IO.StreamReader(entry.Open(), System.Text.Encoding.ASCII)
                                    readLine = sr.ReadLine()
                                    If readLine Is Nothing Then readLine = String.Empty
                                    version = GetVersionFromList("bkp", readLine, True)
                                End Using
                                Exit For
                            Case "apw"
                                Using sr As New IO.StreamReader(entry.Open(), System.Text.Encoding.ASCII)
                                    While Not sr.EndOfStream
                                        Dim line As String = sr.ReadLine()
                                        If line Is Nothing Then Exit While
                                        readLine = line.Trim()
                                        If readLine.Contains("APV") Then
                                            version = GetVersionFromList("apw", readLine, False)
                                            If version <> "" Then Exit While
                                        End If
                                    End While
                                End Using
                                Exit For
                        End Select
                    Next
                End Using

            Case "apw"
                Using sr As New IO.StreamReader(fileName, System.Text.Encoding.ASCII)
                    While Not sr.EndOfStream
                        Dim line As String = sr.ReadLine()
                        If line Is Nothing Then Exit While

                        readLine = line.Trim()
                        If readLine.Contains("APV") Then
                            version = GetVersionFromList(extension, readLine, False)
                            If version <> "" Then Exit While
                        End If
                    End While
                End Using

            Case "hscz"
                Using za As IO.Compression.ZipArchive = IO.Compression.ZipFile.OpenRead(fileName)
                    For Each entry As IO.Compression.ZipArchiveEntry In za.Entries
                        If IO.Path.GetExtension(entry.Name).Replace(".", "").ToLowerInvariant() = "hsc" Then
                            Using br As New IO.BinaryReader(entry.Open())
                                Dim bytes(31) As Byte
                                Dim readCount As Integer = br.Read(bytes, 0, bytes.Length)
                                If readCount >= 26 Then
                                    readLine = String.Empty
                                    For i As Integer = 22 To 25
                                        readLine &= bytes(i).ToString("X2")
                                    Next
                                    version = GetVersionFromList("hsc", readLine, True)
                                End If
                            End Using
                            Exit For
                        End If
                    Next
                End Using

            Case "hsc"
                Using fs As New IO.FileStream(fileName, IO.FileMode.Open, IO.FileAccess.Read)
                    Dim bytes(31) As Byte
                    Dim readCount As Integer = fs.Read(bytes, 0, bytes.Length)
                    If readCount >= 26 Then
                        readLine = ""
                        For i As Integer = 22 To 25
                            readLine &= bytes(i).ToString("X2")
                        Next
                        version = GetVersionFromList(extension, readLine, True)
                    End If
                End Using

            Case "fnwx"
                Using za As IO.Compression.ZipArchive = IO.Compression.ZipFile.OpenRead(fileName)
                    For Each entry As IO.Compression.ZipArchiveEntry In za.Entries
                        If IO.Path.GetExtension(entry.Name).Replace(".", "").ToLowerInvariant() = "fnwxz" Then
                            Using sr As New IO.StreamReader(entry.Open, System.Text.Encoding.UTF8)
                                Dim firstLine As String = sr.ReadLine()
                                readLine = firstLine
                                version = GetVersionFromList("fnwx", readLine, True)
                            End Using
                            Exit For
                        End If
                    Next
                End Using

            Case Else
                MessageBox.Show(
                    "This file extension is not supported by this tool.",
                    "Unsupported File Extension",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

        End Select

        If version <> "" Then
            MessageBox.Show(
                version,
                "Version",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
        End If
    End Sub

    Private Function GetVersionFromList(extension As String,
                                        versionLine As String,
                                        returnVersionLine As Boolean) As String

        If String.IsNullOrEmpty(versionLine) Then
            Return ""
        End If

        For Each kvp In listVersionInfo
            Dim key = kvp.Key

            If key.Item1 = extension AndAlso
               Not String.IsNullOrEmpty(key.Item2) AndAlso
               versionLine.Contains(key.Item2) Then

                Return kvp.Value
            End If
        Next

        If returnVersionLine Then
            Return "Update AspenVersionInfo.txt." & Environment.NewLine & versionLine
        End If

        Return ""
    End Function

End Class
