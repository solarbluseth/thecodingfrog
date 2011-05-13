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
    Private doExportLayerComps As Boolean = False
    Private isJPEG As Boolean = True

    Const FOUND = "Found"
    Const NOT_FOUND = "Not found"
    Private strRootCS3 As String = "\\Photoshop.Image.10\\shell\\Save as JPEG 100%\\command"
    Private strRootCS4 As String = "\\Photoshop.Image.11\\shell\\Save as JPEG 100%\\command"
    Private strRootCS5 As String = "\\Photoshop.Image.12\\shell\\Save as JPEG 100%\\command"
    Private strRootCS55 As String = "\\Photoshop.Image.13\\shell\\Save as JPEG 100%\\command"

    Private appRef As Photoshop.Application
    Private docRef As Photoshop.Document
    Private openDoc As Boolean = True
    Private stayOpen As Boolean = False
    Private i As Integer

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

        Me.Text = System.Reflection.Assembly.GetExecutingAssembly.GetName().Name.ToString()
        Me.LabelCompiled.Text = "V " & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Major.ToString() & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Minor.ToString() & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Build.ToString() & ", Compiled " & CompileDate.BuildDate.ToString()

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
            If key.GetValue("ExcludeDirectories") <> "" Then
                Me.ExcludeDirectories.Text = key.GetValue("ExcludeDirectories")
            Else
                Me.ExcludeDirectories.Text = "+ Elements"
            End If

            key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            'MessageBox.Show(">>> " & key.GetValue("NamedExportQuality"))
            If key.GetValue("NamedExportQuality") <> "" Then
                Me.NamedExportQuality.Value = CInt(key.GetValue("NamedExportQuality"))
            Else
                Me.NamedExportQuality.Value = 6
            End If

            key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            'MessageBox.Show(">>> " & key.GetValue("ExportLayerComps"))
            If key.GetValue("ExportLayerComps") = 1 Then
                Me.ExportLayerComps.Checked = True
                doExportLayerComps = True
            Else
                Me.ExportLayerComps.Checked = False
                doExportLayerComps = False
            End If

            conf_loaded = True
        Catch e As Exception
        End Try
    End Sub

    Private Function isSetup() As Boolean
        If (isVersionInstalled("CS3") Or isVersionInstalled("CS4") Or isVersionInstalled("CS5") Or isVersionInstalled("CS55")) Then
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
            Case "CS55" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.13")
        End Select

        If (key Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function isVersionInstalled(ByVal version As String) As Boolean
        Dim key As RegistryKey
        Dim os = Environment.OSVersion
        Dim res As Object

        If os.Version.Major >= 6 And os.Version.Minor >= 1 Then
            'MessageBox.Show(version)
            Select Case version
                Case "CS3"
                    Try
                        key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

                Case "CS4"
                    Try
                        key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

                Case "CS5"
                    Try
                        key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

                Case "CS55"
                    Try
                        key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.13\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

            End Select

            Try
                res = key.GetValue("SubCommands")
                If (res.ToString <> String.Empty) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        Else
            Select Case version
                Case "CS3" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10\\shell\\Save as JPEG 100% (by index)\\command\")
                Case "CS4" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Save as JPEG 100% (by index)\\command\")
                Case "CS5" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12\\shell\\Save as JPEG 100% (by index)\\command\")
                Case "CS55" : key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.13\\shell\\Save as JPEG 100% (by index)\\command\")
            End Select
            If (key Is Nothing) Then
                Return False
            Else
                Return True
            End If
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
        If (hasVersion("CS55")) Then
            Me.LabelCS55.Text = FOUND
        Else
            Me.LabelCS55.Text = NOT_FOUND
        End If
    End Sub

    Private Sub RegInstall(version as String)

        Dim os = Environment.OSVersion
        Dim newKey As RegistryKey

        If os.Version.Major >= 6 And os.Version.Minor >= 1 Then

            ' Save as JPEG
            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG")
            newKey.SetValue("MUIVerb", "Save as...", RegistryValueKind.String)
            newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.Gif;SaveAsJPEG.Config", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG")
            newKey.SetValue("MUIVerb", "Save as...", RegistryValueKind.String)
            newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.Gif;SaveAsJPEG.Config", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG")
            newKey.SetValue("MUIVerb", "Save as...", RegistryValueKind.String)
            newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.Gif;SaveAsJPEG.Config", RegistryValueKind.String)
            newKey.Close()


            ' SaveAsJPEG.100
            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.100")
            newKey.SetValue("MUIVerb", "JPEG 100% (by index)", RegistryValueKind.String)
            newKey.SetValue("Icon", "shell32.dll,43", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.100\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            ' SaveAsJPEG.60
            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.60")
            newKey.SetValue("MUIVerb", "JPEG 60% (by index)", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.60\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            ' SaveAsJPEG.ByName
            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName")
            newKey.SetValue("MUIVerb", "JPEG 100% (by name)", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""name""", RegistryValueKind.String)
            newKey.Close()

            ' SaveAsJPEG.Gif
            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Gif")
            newKey.SetValue("MUIVerb", "GIF", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Gif\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""gif""", RegistryValueKind.String)
            newKey.Close()

            ' SaveAsJPEG.Config
            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Config")
            newKey.SetValue("MUIVerb", "Configuration", RegistryValueKind.String)
            newKey.SetValue("Icon", "shell32.dll,21", RegistryValueKind.String)
            'newKey.SetValue("CommandFlags ", "20", RegistryValueKind.DWord)
            newKey.Close()

            newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Config\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
            newKey.Close()
        Else
            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by index)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 60% (by index)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG (by name)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""name""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG (Gif)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""gif""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG config\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by index)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60% (by index)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG (by name)\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""name""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG config\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG config\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """", RegistryValueKind.String)
            newKey.Close()
        End If
        Setup()

    End Sub

    Private Sub RegUninstall(ByVal version As String)
        'MessageBox.Show(version)
        Dim os = Environment.OSVersion

        If os.Version.Major >= 6 And os.Version.Minor >= 1 Then
            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.100")
            Catch e As Exception
                'MessageBox.Show (e.Message)
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.60")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Gif")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Config")
            Catch e As Exception
            End Try
        Else
            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 100%\")
            Catch e As Exception
            End Try
            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 60%\")
            Catch e As Exception
            End Try
            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100%\")
            Catch e As Exception
            End Try
            Try
                Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60%\")
            Catch e As Exception
            End Try


            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 60% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG (Gif)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG config\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG config\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG config\")
        End If
        
        Setup()
    End Sub

    Private Sub ProcessFile()
        
        Dim isNamedLayerComp As Boolean = False

        Dim jpgSaveOptions As Photoshop.JPEGSaveOptions = New Photoshop.JPEGSaveOptions
        jpgSaveOptions.EmbedColorProfile = False
        jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline 
        jpgSaveOptions.Matte = 1 ' psNoMatte 

        Dim gifExportOptionsSaveForWeb As Photoshop.ExportOptionsSaveForWeb = New Photoshop.ExportOptionsSaveForWeb
        'gifExportOptionsSaveForWeb.MatteColor = 255
        gifExportOptionsSaveForWeb.Format = 3
        gifExportOptionsSaveForWeb.ColorReduction = 1
        gifExportOptionsSaveForWeb.Colors = 256
        gifExportOptionsSaveForWeb.Dither = 3
        gifExportOptionsSaveForWeb.DitherAmount = 100
        gifExportOptionsSaveForWeb.Quality = 100
        gifExportOptionsSaveForWeb.Transparency = True
        gifExportOptionsSaveForWeb.TransparencyAmount = 100
        gifExportOptionsSaveForWeb.TransparencyDither = 2
        gifExportOptionsSaveForWeb.IncludeProfile = False
        gifExportOptionsSaveForWeb.Lossy = 0
        gifExportOptionsSaveForWeb.WebSnap = 0

        Select Case args(3)
            Case "name"
                isNamedLayerComp = True
                jpgSaveOptions.Quality = CInt(Me.NamedExportQuality.Value)
            Case "index"
                isNamedLayerComp = False
                jpgSaveOptions.Quality = CInt(args(1))
            Case "gif"
                isNamedLayerComp = False
                isJPEG = False
            Case Else
                isNamedLayerComp = False
                jpgSaveOptions.Quality = CInt(args(1))
        End Select


        appRef = New Photoshop.Application()
        Try
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
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        If openDoc Then
            Try
                docRef = appRef.Open(args(2))
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

        Dim compsCount As Integer
        Dim compsIndex As Integer
        Dim compRef As Photoshop.LayerComp
        Dim duppedDocument As Photoshop.Document
        Dim fileNameBody As String

        compsCount = docRef.LayerComps.Count


        ' Exporting layercomps by index or name
        If doExportLayerComps Then
            If compsCount <= 1 Then
                'Set textItemRef = appRef.ActiveDocument.Layers(1) 

                'textItemRef.TextItem.Contents = Args.Item(1) 

                'outFileName = Args.Item(1)
                If isJPEG Then
                    docRef.SaveAs(args(2), jpgSaveOptions, True)
                Else
                    fileNameBody = docRef.Name.Substring(0, docRef.Name.LastIndexOf(".")) & ".gif"
                    docRef.Export(docRef.Path & fileNameBody, 2, gifExportOptionsSaveForWeb)
                End If

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
                        fileNameBody = docRef.Name.Substring(0, docRef.Name.LastIndexOf(".")) & "." & compsIndex
                    Else
                        fileNameBody = compRef.Name
                    End If
                    'msgbox(fileNameBody)
                    If isJPEG Then
                        fileNameBody = fileNameBody & ".jpg"
                        duppedDocument.SaveAs(docRef.Path & fileNameBody, jpgSaveOptions, True)
                    Else
                        fileNameBody = fileNameBody & ".gif"
                        duppedDocument.Export(docRef.Path & fileNameBody, 2, gifExportOptionsSaveForWeb)
                    End If
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

            'MsgBox(Me.AutoArchive.Checked)
            Dim di As DirectoryInfo
            di = New DirectoryInfo(docRef.Path)

            'MsgBox(di.Name)
            If Me.AutoArchive.Checked And Not isExcludeDirectory(di.Name) Then

                'Dim di As DirectoryInfo
                Dim afi() As FileInfo
                Dim fi As FileInfo

                'MsgBox(Directory.Exists(docRef.Path & "\" & Me.ArchiveDirectory.Text & "\"))
                If Not Directory.Exists(docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
                    Directory.CreateDirectory(docRef.Path & "\" & Me.ArchiveDirectory.Text & "\")
                End If

                'di = New DirectoryInfo(docRef.Path)
                Dim currentFileName As String = docRef.Name.Substring(0, docRef.Name.LastIndexOf("."))

                Dim RegexObj As Regex = New Regex("\d*$")
                Dim myMatches As Match
                Dim currentVersion
                Dim cleanFileName As String
                If RegexObj.IsMatch(currentFileName) Then
                    myMatches = RegexObj.Match(currentFileName)
                    currentVersion = myMatches.Value
                    cleanFileName = RegexObj.Replace(currentFileName, "")
                    cleanFileName = cleanFileName.Replace("+", "\+")
                    cleanFileName = cleanFileName.Replace(" ", "\s")
                    'MsgBox(cleanFileName)
                    Dim RegexObj2 As Regex = New Regex("^" & cleanFileName & "(\d+|\.)")

                    afi = di.GetFiles("*.*")
                    For Each fi In afi
                        If RegexObj2.IsMatch(fi.Name) Then
                            'MsgBox(fi.Name)
                            If isOldFileVersion(fi.Name, currentVersion) And Directory.Exists(docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
                                Try
                                    File.Copy(docRef.Path & fi.Name, docRef.Path & "\" & Me.ArchiveDirectory.Text & "\" & fi.Name, True)
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                Finally
                                    File.Delete(docRef.Path & fi.Name)
                                End Try

                            End If
                        End If
                    Next
                End If
            End If
        Else 'Exporting each layers by name
            Dim oLayer
            For compsIndex = 1 To docRef.Layers.Count()
                oLayer = docRef.Layers.Item(compsIndex)
                'isVisible = oLayer.visible
                appRef.ActiveDocument.ActiveLayer = oLayer
                'oLayer.Apply()
                'duppedDocument = docRef.Duplicate()
                'msgbox(compRef.Name)
                fileNameBody = oLayer.Name & ".jpg"
                'msgbox(fileNameBody)
                docRef.SaveAs(docRef.Path & fileNameBody, jpgSaveOptions, True)
                appRef.ActiveDocument.ActiveLayer.visible = False
                'duppedDocument.Close(2)
            Next
        End If

        If Not stayOpen Then docRef.Close(2)

        ' End program
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

    Private Function isExcludeDirectory(ByVal dirName As String)
        Dim ExcludeDirectories As Array
        Dim ExcludeDirectory As String

        If Me.ExcludeDirectories.Text <> "" Then
            ExcludeDirectories = Split(Me.ExcludeDirectories.Text & ";" & me.ArchiveDirectory.Text, ";")
            For Each ExcludeDirectory In ExcludeDirectories
                If ExcludeDirectory = dirName Then
                    Return True
                End If
            Next
            Return False
        Else
            Return False
        End If
    End Function

    Private Sub Install_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Install.Click
        If isSetup() Then
            If hasVersion("CS3") Then RegUninstall("10")
            If hasVersion("CS4") Then RegUninstall("11")
            If hasVersion("CS5") Then RegUninstall("12")
            If hasVersion("CS55") Then RegUninstall("13")
        Else
            If hasVersion("CS3") Then RegInstall("10")
            If hasVersion("CS4") Then RegInstall("11")
            If hasVersion("CS5") Then RegInstall("12")
            If hasVersion("CS55") Then RegInstall("13")
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

    Private Sub ExportLayerComps_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportLayerComps.CheckedChanged
        Dim newKey As RegistryKey
        newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Me.ExportLayerComps.Checked Then
            newKey.SetValue("ExportLayerComps", "1", RegistryValueKind.String)
            doExportLayerComps = True
        Else
            newKey.SetValue("ExportLayerComps", "0", RegistryValueKind.String)
            doExportLayerComps = False
        End If
        newKey.Close()
    End Sub

    Private Sub ToolTip1_Popup(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub ExportLayerComps_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExportLayerComps.MouseHover
        ToolTip1.Show("Export layercomps if checked, else it will save each layers", ExportLayerComps)
    End Sub

    Private Sub ExcludeDirectories_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcludeDirectories.TextChanged
        Dim newKey As RegistryKey
        newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Trim(Me.ExcludeDirectories.Text) <> "" Then
            newKey.SetValue("ExcludeDirectories", Me.ExcludeDirectories.Text, RegistryValueKind.String)
        Else
            Me.ExcludeDirectories.Text = "+ Elements"
            newKey.SetValue("ExcludeDirectories", "+ Elements", RegistryValueKind.String)
        End If
        newKey.Close()
    End Sub
End Class