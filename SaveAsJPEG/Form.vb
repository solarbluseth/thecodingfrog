Imports Microsoft.Win32
Imports System.Text.RegularExpressions

Public Class Form
    Public Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Integer) As Integer
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer

    Const SW_HIDE = 0
    Const SW_SHOWNORMAL = 1
    Const SW_NORMAL = 1
    Const SW_SHOWMINIMIZED = 2
    Private args As String()
    Private conf_loaded As Boolean = False

    Const FOUND = "Found"
    Const NOT_FOUND = "Not found"
    Private strRootCS3 As String = "\\Photoshop.Image.10\\shell\\Save as JPEG 100%\\command"
    Private strRootCS4 As String = "\\Photoshop.Image.11\\shell\\Save as JPEG 100%\\command"
    Private strRootCS5 As String = "\\Photoshop.Image.12\\shell\\Save as JPEG 100%\\command"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Dim strTitle As String
        Dim rtnLen As Long
        Dim hwnd As Int32

        strTitle = Space(256)
        rtnLen = GetConsoleTitle(strTitle, 256)
        If rtnLen > 0 Then
            strTitle = Strings.Left$(strTitle, rtnLen)
        End If

        hwnd = FindWindow(vbNullString, strTitle)

        Me.Text = System.Reflection.Assembly.GetExecutingAssembly.GetName().Name.ToString() & " v" & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Major.ToString() & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Minor.ToString() & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Build.ToString()

        loadConf()
        If isSetup() Then
            args = Environment.GetCommandLineArgs()
            Dim num As Integer = UBound(args)
            If num < 2 Then
                ShowWindow(hwnd, SW_SHOWNORMAL)
                Setup()
            Else
                ' hide the app
                ShowWindow(hwnd, SW_HIDE)
                Me.Visible = False
                Me.ShowInTaskbar = False
                Me.WindowState = FormWindowState.Minimized
                ProcessFile()
            End If
        Else
            Setup()
        End If
    End Sub

    Private Sub loadConf()
        Try
            Dim key As RegistryKey
            key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            If key.GetValue("AutoArchive") = 1 Then
                Me.AutoArchive.Checked = True
            Else
                Me.AutoArchive.Checked = False
            End If

            key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            If key.GetValue("ArchiveDirectory") <> "" Then
                Me.ArchiveDirectory.Text = key.GetValue("ArchiveDirectory")
            Else
                Me.ArchiveDirectory.Text = "Archives"
            End If

            key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            'MessageBox.Show(">>> " & key.GetValue("NamedExportQuality"))
            If key.GetValue("NamedExportQuality") <> "" Then
                Me.NamedExportQuality.Value = CInt(key.GetValue("NamedExportQuality"))
            Else
                Me.NamedExportQuality.Value = 6
            End If

            conf_loaded = True
        Catch e As Exception
        End Try
    End Sub

    Private Function isSetup() As Boolean
        If (isVersionInstalled("CS3") Or isVersionInstalled("CS4") Or isVersionInstalled("CS5")) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function hasVersion(ByVal version As String) As Boolean
        Dim key As RegistryKey

        Select Case version
            Case "CS3" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10")
            Case "CS4" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11")
            Case "CS5" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12")
        End Select

        If (key Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function isVersionInstalled(ByVal version As String) As Boolean
        Dim key As RegistryKey

        Select Case Version
            Case "CS3" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10\\shell\\Save as JPEG 100% (by index)\\command\")
            Case "CS4" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Save as JPEG 100% (by index)\\command\")
            Case "CS5" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12\\shell\\Save as JPEG 100% (by index)\\command\")
        End Select

        If (key Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Setup()
        If isSetup() Then
            Me.Install.Text = "Uninstall"
        Else
            Me.Install.Text = "Install"
        End If
        If (hasVersion("CS3")) Then
            Me.LabelCS3.Text = FOUND
        Else
            Me.LabelCS3.Text = NOT_FOUND
        End If
        If (hasVersion("CS4")) Then
            Me.LabelCS4.Text = FOUND
        Else
            Me.LabelCS4.Text = NOT_FOUND
        End If
        If (hasVersion("CS5")) Then
            Me.LabelCS5.Text = FOUND
        Else
            Me.LabelCS5.Text = NOT_FOUND
        End If
    End Sub

    Private Sub RegInstall(version as String)

        Dim newKey As RegistryKey
        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by index)\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 60% (by index)\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG (by name)\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-1"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG config\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by index)\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60% (by index)\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG (by name)\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-1"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG config\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1""", RegistryValueKind.String)
        newKey.Close()

        newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG config\\command")
        newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
        newKey.Close()
        Setup()

    End Sub

    Private Sub RegUninstall(ByVal version As String)
        Try
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 100%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 60%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60%\")

            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 60% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG config\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG config\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG config\")
        Catch e As Exception
        End Try
        Setup()
    End Sub

    Private Sub ProcessFile()

        Dim appRef As Photoshop.Application = New Photoshop.Application()
        Dim docRef As Photoshop.Document
        Dim openDoc As Boolean = True
        Dim stayOpen As Boolean = False
        Dim i As Integer
        Dim isNamedLayerComp As Boolean = False

        'MessageBox.Show("ok")
        If args(1) = -1 Then isNamedLayerComp = True


        On Error Resume Next
        If appRef.Documents.Count > 0 Then
            For i = 1 To appRef.Documents.Count
                If args(2) = appRef.Documents.Item(i).FullName Then
                    appRef.ActiveDocument = appRef.Documents.Item(i)
                    docRef = appRef.ActiveDocument
                    openDoc = False
                    stayOpen = True
                End If
            Next
        End If
        If Err.Number <> 0 Then
            Call showErr("opened")
        End If
        On Error GoTo 0

        If openDoc Then
            On Error Resume Next
            docRef = appRef.Open(args(2))
            If Err.Number <> 0 Then
                Call showErr("opened")
            End If
            On Error GoTo 0
        End If

        Dim compsCount As Integer
        Dim compsIndex As Integer
        Dim compRef As Photoshop.LayerComp
        Dim duppedDocument As Photoshop.Document
        Dim fileNameBody As String

        compsCount = docRef.LayerComps.Count

        Dim jpgSaveOptions As Photoshop.JPEGSaveOptions = New Photoshop.JPEGSaveOptions
        jpgSaveOptions.EmbedColorProfile = False
        jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline 
        jpgSaveOptions.Matte = 1 ' psNoMatte 
        If Not isNamedLayerComp Then
            jpgSaveOptions.Quality = CInt(args(1))
        Else
            jpgSaveOptions.Quality = CInt(Me.NamedExportQuality.Value)
        End If

        If compsCount <= 1 Then
            'Set textItemRef = appRef.ActiveDocument.Layers(1) 

            'textItemRef.TextItem.Contents = Args.Item(1) 

            'outFileName = Args.Item(1)
            docRef.SaveAs(args(2), jpgSaveOptions, True)
        Else
            'msgbox("comps!")
            For compsIndex = 1 To compsCount
                'MsgBox(docRef.LayerComps.Count)
                'End
                compRef = docRef.LayerComps.Item(compsIndex)
                'if (exportInfo.selectionOnly && !compRef.selected) continue; // selected only
                compRef.Apply()
                duppedDocument = docRef.Duplicate()
                'msgbox(compRef.Name)
                If Not isNamedLayerComp Then
                    fileNameBody = Split(docRef.Name, ".")(0) & "." & compsIndex & ".jpg"
                Else
                    fileNameBody = compRef.Name & ".jpg"
                End If
                'msgbox(fileNameBody)
                duppedDocument.SaveAs(docRef.Path & fileNameBody, jpgSaveOptions, True)
                duppedDocument.Close(2)
                'fileNameBody += "_" + zeroSuppress(compsIndex, 4);
                'fileNameBody += "_" + compRef.name;
                'if (null != compRef.comment)    fileNameBody += "_" + compRef.comment;
                'fileNameBody = fileNameBody.replace(/[:\/\\*\?\"\<\>\|\\\r\\\n]/g, "_");  // '/\:*?"<>|\r\n' -> '_'
                'if (fileNameBody.length > 120) fileNameBody = fileNameBody.substring(0,120);
                'saveFile(duppedDocument, fileNameBody, exportInfo);
                'duppedDocument.close(SaveOptions.DONOTSAVECHANGES);
            Next
        compRef = docRef.LayerComps.Item(1)
        compRef.Apply()
        End If

        If Me.AutoArchive.Checked Then

            Dim di As DirectoryInfo
            Dim afi() As FileInfo
            Dim fi As FileInfo


            di = New DirectoryInfo(docRef.Path)
            Dim currentFileName As String = docRef.Name.Substring(0, docRef.Name.LastIndexOf("."))

            Dim RegexObj As Regex = New Regex("\d*$")
            Dim myMatches As Match
            Dim currentVersion
            Dim cleanFileName As String
            If RegexObj.IsMatch(currentFileName) Then
                myMatches = RegexObj.Match(currentFileName)
                currentVersion = myMatches.Value
                cleanFileName = RegexObj.Replace(currentFileName, "")
                'MsgBox(currentVersion)
                Dim RegexObj2 As Regex = New Regex("^" & cleanFileName & "(\d+|\.)")

                afi = di.GetFiles("*.*")
                For Each fi In afi
                    If RegexObj2.IsMatch(fi.Name) Then
                        'MsgBox(fi.Name)
                        If isOldFileVersion(fi.Name, currentVersion) Then
                            On Error Resume Next
                            'MsgBox(fi.Name)
                            File.Move(docRef.Path & fi.Name, docRef.Path & "\" & Me.ArchiveDirectory.Text & "\" & fi.Name)
                        End If
                    End If
                Next
            End If
        End If

        If Not stayOpen Then docRef.Close(2)
        End
    End Sub

    Private Function isOldFileVersion(ByVal filename As String, ByVal version As String) As String
        Dim newFileName As String = filename.Substring(0, filename.LastIndexOf("."))
        'MsgBox(newFileName)
        Dim RegexObj As Regex = New Regex("(\d+)\.*\d*$")
        If RegexObj.IsMatch(newFileName) Then
            'MsgBox("> " & RegexObj.Match(newFileName).Groups(1).Value)
            If CInt(RegexObj.Match(newFileName).Groups(1).Value) < CInt(version) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Private Sub showErr(ByVal errType)
        Select Case errType
            Case "opened"
                MsgBox("Photoshop est en mode saisie sur un layer !")
                End
        End Select
    End Sub

    Private Sub Install_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Install.Click
        If isSetup() Then
            If hasVersion("CS3") Then RegUninstall("10")
            If hasVersion("CS4") Then RegUninstall("11")
            If hasVersion("CS5") Then RegUninstall("12")
        Else
            If hasVersion("CS3") Then RegInstall("10")
            If hasVersion("CS4") Then RegInstall("11")
            If hasVersion("CS5") Then RegInstall("12")
        End If
    End Sub

    Private Sub AutoArchive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoArchive.CheckedChanged
        Dim newKey As RegistryKey
        newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Me.AutoArchive.Checked Then
            newKey.SetValue("AutoArchive", "1", RegistryValueKind.String)
        Else
            newKey.SetValue("AutoArchive", "0", RegistryValueKind.String)
        End If
        newKey.Close()
    End Sub

    Private Sub ArchiveDirectory_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchiveDirectory.TextChanged
        Dim newKey As RegistryKey
        newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Trim(Me.ArchiveDirectory.Text) <> "" Then
            newKey.SetValue("ArchiveDirectory", Me.ArchiveDirectory.Text, RegistryValueKind.String)
        Else
            Me.ArchiveDirectory.Text = "Archives"
            newKey.SetValue("ArchiveDirectory", "Archives", RegistryValueKind.String)
        End If
        newKey.Close()
    End Sub

    Private Sub NamedExportQuality_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NamedExportQuality.ValueChanged
        Dim newKey As RegistryKey
        newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If conf_loaded Then
            'MessageBox.Show(Me.NamedExportQuality.Value.ToString())
            newKey.SetValue("NamedExportQuality", Me.NamedExportQuality.Value, RegistryValueKind.String)
        End If
        newKey.Close()
    End Sub
End Class