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
    Private __args As String()
    Private __confLoaded As Boolean = False
    Private __doExportLayerComps As Boolean = True
    Private __imageType As String

    Const NS = "http://www.smartobjectlinks.com/1.0/"
    Const FOUND = "Found"
    Const NOT_FOUND = "Not found"
    'Private __strRootCS3 As String = "\\Photoshop.Image.10\\shell\\Save as JPEG 100%\\command"
    'Private __strRootCS4 As String = "\\Photoshop.Image.11\\shell\\Save as JPEG 100%\\command"
    'Private __strRootCS5 As String = "\\Photoshop.Image.12\\shell\\Save as JPEG 100%\\command"
    'Private __strRootCS55 As String = "\\Photoshop.Image.55\\shell\\Save as JPEG 100%\\command"

    Private __appRef As Photoshop.Application
    Private __docRef As Photoshop.Document
    Private __openDoc As Boolean = True
    Private __stayOpen As Boolean = False
    Private i As Integer

    'Public Enum Versions
    '    CS3 = 10
    '    CS4 = 11
    '    CS5 = 12
    '    CS55 = 55
    'End Enum

    Public Enum Colors
        NONE
        RED
        ORANGE
        YELLOW
        GREEN
        BLUE
        VIOLET
        GRAY
    End Enum

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Dim __strTitle As String
        Dim __rtnLen As Long
        Dim __hwnd As Int32

        __strTitle = Space(256)
        __rtnLen = GetConsoleTitle(__strTitle, 256)
        If __rtnLen > 0 Then
            __strTitle = Strings.Left$(__strTitle, __rtnLen)
        End If

        __hwnd = FindWindow(vbNullString, __strTitle)

        Me.Text = System.Reflection.Assembly.GetExecutingAssembly.GetName().Name.ToString()
        Me.LabelCompiled.Text = "V " & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Major.ToString() & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Minor.ToString() & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.Build.ToString() & ", Compiled " & CompileDate.BuildDate.ToString()

        loadConf()
        If isSetup() Then
            __args = Environment.GetCommandLineArgs()
            Dim __num As Integer = UBound(__args)
            'MessageBox.Show(args(1))
            If __num = 1 Then
                Select Case __args(1).ToLower
                    Case "-c" : ShowWindow(__hwnd, SW_SHOWNORMAL)
                        Setup()
                End Select
            ElseIf __num = 2 Then
                ShowWindow(__hwnd, SW_HIDE)
                Me.Visible = False
                Me.ShowInTaskbar = False
                Me.WindowState = FormWindowState.Minimized

                Select Case __args(1).ToLower
                    Case "-so" : ExportSmartObjects()
                    Case "-r" : ExportImagesRights()
                    Case "-w" : CleanLayersName()
                    Case "-sc" : SaveScreenSelection()
                End Select
            ElseIf __num = 4 Then
                ' hide the app
                ShowWindow(__hwnd, SW_HIDE)
                Me.Visible = False
                Me.ShowInTaskbar = False
                Me.WindowState = FormWindowState.Minimized
                ProcessFile()
            Else
                Setup()
            End If
        Else
            Setup()
        End If
    End Sub

    Private Sub loadConf()
        Try
            Dim __key As RegistryKey
            __key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            If __key.GetValue("AutoArchive") = 1 Then
                Me.AutoArchive.Checked = True
            Else
                Me.AutoArchive.Checked = False
            End If

            __key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            If __key.GetValue("ArchiveDirectory") <> "" Then
                Me.ArchiveDirectory.Text = __key.GetValue("ArchiveDirectory")
            Else
                Me.ArchiveDirectory.Text = "Archives"
            End If

            __key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            If __key.GetValue("ExcludeDirectories") <> "" Then
                Me.ExcludeDirectories.Text = __key.GetValue("ExcludeDirectories")
            Else
                Me.ExcludeDirectories.Text = ""
            End If

            __key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            'MessageBox.Show(">>> " & key.GetValue("NamedExportQuality"))
            If __key.GetValue("NamedExportQuality") <> "" Then
                Me.NamedExportQuality.Value = CInt(__key.GetValue("NamedExportQuality"))
            Else
                Me.NamedExportQuality.Value = 6
            End If

            __key = Registry.CurrentUser.OpenSubKey("Software\\SaveAsJPEG\\")
            'MessageBox.Show(">>> " & key.GetValue("ExportLayerComps"))
            If __key.GetValue("ExportLayerComps") = 1 Then
                Me.ExportLayerComps.Checked = True
                __doExportLayerComps = True
            Else
                Me.ExportLayerComps.Checked = False
                __doExportLayerComps = False
            End If

            __confLoaded = True
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

    Private Function hasVersion(ByVal __version As String) As Boolean
        'Dim Version As Versions
        Dim __key As RegistryKey

        'For Each Version In [Enum].GetValues(GetType(Versions))
        'key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image." & Version.ToString)
        'Next


        Select Case __version
            Case "CS3" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10")
            Case "CS4" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11")
            Case "CS5" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12")
            Case "CS55" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.55")
        End Select

        If (__key Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function isVersionInstalled(ByVal __version As String) As Boolean
        Dim __key As RegistryKey
        Dim __os = Environment.OSVersion
        Dim __res As Object

        If __os.Version.Major >= 6 And __os.Version.Minor >= 1 Then
            'MessageBox.Show(version)
            Select Case __version
                Case "CS3"
                    Try
                        __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

                Case "CS4"
                    Try
                        __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

                Case "CS5"
                    Try
                        __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

                Case "CS55"
                    Try
                        __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.55\\shell\\Save as JPEG")
                    Catch ex As Exception

                    End Try

            End Select

            Try
                __res = __key.GetValue("SubCommands")
                If (__res.ToString <> String.Empty) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        Else
            Select Case __version
                Case "CS3" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.10\\shell\\Save as JPEG 100% (by index)\\command\")
                Case "CS4" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Save as JPEG 100% (by index)\\command\")
                Case "CS5" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.12\\shell\\Save as JPEG 100% (by index)\\command\")
                Case "CS55" : __key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.55\\shell\\Save as JPEG 100% (by index)\\command\")
            End Select
            If (__key Is Nothing) Then
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

    Private Sub RegInstall(ByVal version As String, ByVal illustratorversion As String)

        Dim __os = Environment.OSVersion
        Dim __newKey As RegistryKey

        If __os.Version.Major >= 6 And __os.Version.Minor >= 1 Then

            ' JPEG BASE64
            __newKey = Registry.ClassesRoot.CreateSubKey("ACDSee Pro 4.jpg\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.Base64;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("jpegfile\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.Base64;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            ' GIF BASE64
            __newKey = Registry.ClassesRoot.CreateSubKey("ACDSee Pro 4.gif\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.Base64;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("giffile\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.Base64;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            ' PNG BASE64
            __newKey = Registry.ClassesRoot.CreateSubKey("ACDSee Pro 4.png\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.Base64;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("pngfile\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.Base64;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()




            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Base64")
            __newKey.SetValue("MUIVerb", "Export Base64 URI", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,43", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Base64\\Command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""base64"" ""index""", RegistryValueKind.String)
            __newKey.Close()




            ' Save as JPEG
            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.ByName60;SaveAsJPEG.PngIndex;SaveAsJPEG.PngName;SaveAsJPEG.Gif;SaveAsJPEG.Screen;SaveAsJPEG.ImagesRights;SaveAsJPEG.SO;SaveAsJPEG.Clean;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.ByName60;SaveAsJPEG.PngIndex;SaveAsJPEG.PngName;SaveAsJPEG.Gif;SaveAsJPEG.Screen;SaveAsJPEG.ImagesRights;SaveAsJPEG.SO;SaveAsJPEG.Clean;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.ByName60;SaveAsJPEG.PngIndex;SaveAsJPEG.PngName;SaveAsJPEG.Gif;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator." & illustratorversion & "\\shell\\Save as JPEG")
            __newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String)
            __newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.SetValue("SubCommands", "SaveAsJPEG.100;SaveAsJPEG.60;SaveAsJPEG.ByName;SaveAsJPEG.ByName60;SaveAsJPEG.PngIndex;SaveAsJPEG.PngName;SaveAsJPEG.Gif;SaveAsJPEG.Config", RegistryValueKind.String)
            __newKey.Close()


            ' SaveAsJPEG.100
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.100")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 100% (by index)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,43", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.100\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""jpg"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.60
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.60")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 60% (by index)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.60\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""jpg"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.ByName
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 100% (by name)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""jpg"" ""name""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.ByName60
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName60")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 60% (by name)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName60\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""6"" ""%1"" ""jpg"" ""name""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.PngIndex
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.PngIndex")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As PNG (by index)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.PngIndex\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""png"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.PngName
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.PngName")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As PNG (by name)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.PngName\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""png"" ""name""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.Gif
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Gif")
            __newKey.SetValue("MUIVerb", "Save Layer Comps As GIF (by index)", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String)
            'newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Gif\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""gif"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.ImagesRights
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ImagesRights")
            __newKey.SetValue("MUIVerb", "List Images Rights", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,54", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ImagesRights\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-r"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.SO
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.SO")
            __newKey.SetValue("MUIVerb", "Export Smart Objects", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,132", RegistryValueKind.String)
            'newKey.SetValue("CommandFlags ", "20", RegistryValueKind.DWord)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.SO\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-so"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.Clean
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Clean")
            __newKey.SetValue("MUIVerb", "Clean Layers Name", RegistryValueKind.String)
            '__newKey.SetValue("Icon", "shell32.dll,238", RegistryValueKind.String)
            'newKey.SetValue("CommandFlags ", "20", RegistryValueKind.DWord)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Clean\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-w"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.Screen
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Screen")
            __newKey.SetValue("MUIVerb", "Save Screen Selection As JPEG", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,43", RegistryValueKind.String)
            'newKey.SetValue("CommandFlags ", "20", RegistryValueKind.DWord)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Screen\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-sc"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            ' SaveAsJPEG.Config
            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Config")
            __newKey.SetValue("MUIVerb", "Configuration", RegistryValueKind.String)
            __newKey.SetValue("Icon", "shell32.dll,21", RegistryValueKind.String)
            'newKey.SetValue("CommandFlags ", "20", RegistryValueKind.DWord)
            __newKey.Close()

            __newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Config\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-c""", RegistryValueKind.String)
            __newKey.Close()
        Else
            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by index)\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            '__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 60% (by index)\\command")
            '__newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            '__newKey.Close()

            '__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by name)\\command")
            '__newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""name""", RegistryValueKind.String)
            '__newKey.Close()

            '__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG (Png)\\command")
            '__newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""png""", RegistryValueKind.String)
            '__newKey.Close()

            '__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG (Gif)\\command")
            '__newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""gif""", RegistryValueKind.String)
            '__newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG List Images Rights\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-r"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            '__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG Export Smart Objects\\command")
            '__newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-so"" ""%1""", RegistryValueKind.String)
            '__newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG Clean Layers Name\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-w"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." & version & "\\shell\\Save as JPEG Config\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-c""", RegistryValueKind.String)
            __newKey.Close()




            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by index)\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60% (by index)\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by name)\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""name""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG List Images Rights\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-r"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG Export Smart Objects\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-so"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG Clean Layers Name\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-w"" ""%1""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG Config\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-c""", RegistryValueKind.String)
            __newKey.Close()



            __newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1"" ""index""", RegistryValueKind.String)
            __newKey.Close()

            __newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG Config\\command")
            __newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""-c""", RegistryValueKind.String)
            __newKey.Close()
        End If
        Setup()

    End Sub

    Private Sub RegUninstall(ByVal version As String, ByVal illustratorversion As String)
        'MessageBox.Show(version)
        Dim __os = Environment.OSVersion

        If __os.Version.Major >= 6 And __os.Version.Minor >= 1 Then
            Try
                Registry.ClassesRoot.DeleteSubKeyTree("ACDSee Pro 4.jpg\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("jpegfile\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("ACDSee Pro 4.gif\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("giffile\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("ACDSee Pro 4.png\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.ClassesRoot.DeleteSubKeyTree("pngfile\\shell\\Save as JPEG")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Base64")
            Catch e As Exception
                'MessageBox.Show (e.Message)
            End Try






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
                Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator." & illustratorversion & "\\shell\\Save as JPEG")
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
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ByName60")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.PngIndex")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.PngName")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Gif")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.ImagesRights")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.SO")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Clean")
            Catch e As Exception
            End Try

            Try
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\SaveAsJPEG.Screen")
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
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 100% (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG 60% (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG (Png)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG (Gif)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG List Images Rights\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG Export Smart Objects\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG Clean Layers Name\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." & version & "\\shell\\Save as JPEG Config\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 60% (by index)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG 100% (by name)\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG List Images Rights\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG Export Smart Objects\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG Clean Layers Name\")
            Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." & version & "\\shell\\Save as JPEG Config\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\")
            Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\shell\\Save as JPEG Config\")
        End If

        Setup()
    End Sub

    Private Sub ProcessFile()
        'MessageBox.Show("ProcessFile")
        Dim __isNamedLayerComp As Boolean = False

        Dim __jpgSaveOptions As Photoshop.JPEGSaveOptions

        Try
            __jpgSaveOptions = New Photoshop.JPEGSaveOptions()
            __jpgSaveOptions.EmbedColorProfile = False
            __jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline 
            __jpgSaveOptions.Matte = 1 ' psNoMatte 
        Catch ex As Exception
            Dim __dr As DialogResult = MessageBox.Show("Photoshop is busy with open dialog or something." & vbCrLf & vbCrLf & "Please switch to Photoshop then close open dialogs or leave editing state", "Photoshop not ready", MessageBoxButtons.OK)
            If __dr.OK Then
                End
            End If
        End Try

        Dim __gifExportOptionsSaveForWeb As Photoshop.ExportOptionsSaveForWeb = New Photoshop.ExportOptionsSaveForWeb()
        'gifExportOptionsSaveForWeb.MatteColor = 255
        __gifExportOptionsSaveForWeb.Format = 3
        __gifExportOptionsSaveForWeb.ColorReduction = 1
        __gifExportOptionsSaveForWeb.Colors = 256
        __gifExportOptionsSaveForWeb.Dither = 3
        __gifExportOptionsSaveForWeb.DitherAmount = 100
        __gifExportOptionsSaveForWeb.Quality = 100
        __gifExportOptionsSaveForWeb.Transparency = True
        __gifExportOptionsSaveForWeb.TransparencyAmount = 100
        __gifExportOptionsSaveForWeb.TransparencyDither = 2
        __gifExportOptionsSaveForWeb.IncludeProfile = False
        __gifExportOptionsSaveForWeb.Lossy = 0
        __gifExportOptionsSaveForWeb.WebSnap = 0

        Dim __pngExportOptionsSaveForWeb As Photoshop.ExportOptionsSaveForWeb = New Photoshop.ExportOptionsSaveForWeb()
        __pngExportOptionsSaveForWeb.Format = 13
        __pngExportOptionsSaveForWeb.PNG8 = False
        __pngExportOptionsSaveForWeb.Transparency = True

        Select Case __args(3)
            Case "jpg"
                __jpgSaveOptions.Quality = CInt(__args(1))
                __imageType = "JPG"
                __doExportLayerComps = True 'force using this mode
            Case "png"
                __imageType = "PNG"
                __doExportLayerComps = True
            Case "gif"
                __imageType = "GIF"
            Case "base64"
                ExportBase64()
                End
            Case Else
                __jpgSaveOptions.Quality = CInt(__args(1))
                __imageType = "JPG"
        End Select
        'MessageBox.Show(__imageType)

        Select Case __args(4)
            Case "name"
                __isNamedLayerComp = True
            Case "index"
                __isNamedLayerComp = False
            Case Else
                __isNamedLayerComp = False
        End Select

        Call OpenDocument()

        Dim __compsCount As Integer
        Dim __compsIndex As Integer
        Dim __compRef As Photoshop.LayerComp
        Dim __duppedDocument As Photoshop.Document
        Dim __fileNameBody As String

        __compsCount = __docRef.LayerComps.Count

        ' Exporting layercomps by index or name
        If __doExportLayerComps Then
            If __compsCount <= 1 Then
                'Set textItemRef = appRef.ActiveDocument.Layers(1) 

                'textItemRef.TextItem.Contents = Args.Item(1) 

                'outFileName = Args.Item(1)
                If __imageType = "JPG" Then
                    __docRef.SaveAs(__args(2), __jpgSaveOptions, True)
                ElseIf __imageType = "PNG" Then
                    __fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) & ".png"
                    __docRef.Export(__docRef.Path & __fileNameBody, 2, __pngExportOptionsSaveForWeb)
                Else
                    __fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) & ".gif"
                    __docRef.Export(__docRef.Path & __fileNameBody, 2, __gifExportOptionsSaveForWeb)
                End If

            Else
                'msgbox("comps!")
                For __compsIndex = 1 To __compsCount
                    'MsgBox(docRef.LayerComps.Count)
                    'End
                    __compRef = __docRef.LayerComps.Item(__compsIndex)
                    'if (exportInfo.selectionOnly && !compRef.selected) continue; // selected only
                    __compRef.Apply()
                    __duppedDocument = __docRef.Duplicate()
                    'msgbox(compRef.Name)
                    If Not __isNamedLayerComp Then
                        __fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) & "." & __compsIndex
                    Else
                        __fileNameBody = __compRef.Name
                    End If
                    'msgbox(fileNameBody)
                    If __imageType = "JPG" Then
                        __fileNameBody = __fileNameBody & ".jpg"
                        __duppedDocument.SaveAs(__docRef.Path & __fileNameBody, __jpgSaveOptions, True)
                    ElseIf __imageType = "PNG" Then
                        __fileNameBody = __fileNameBody & ".png"
                        __duppedDocument.Export(__docRef.Path & __fileNameBody, 2, __pngExportOptionsSaveForWeb)
                    Else
                        __fileNameBody = __fileNameBody & ".gif"
                        __duppedDocument.Export(__docRef.Path & __fileNameBody, 2, __gifExportOptionsSaveForWeb)
                    End If
                    __duppedDocument.Close(2)
                    'fileNameBody += "_" + zeroSuppress(compsIndex, 4);
                    'fileNameBody += "_" + compRef.name;
                    'if (null != compRef.comment)    fileNameBody += "_" + compRef.comment;
                    'fileNameBody = fileNameBody.replace(/[:\/\\*\?\"\<\>\|\\\r\\\n]/g, "_");  // '/\:*?"<>|\r\n' -> '_'
                    'if (fileNameBody.length > 120) fileNameBody = fileNameBody.substring(0,120);
                    'saveFile(duppedDocument, fileNameBody, exportInfo);
                    'duppedDocument.close(SaveOptions.DONOTSAVECHANGES);
                Next
                __compRef = __docRef.LayerComps.Item(1)
                __compRef.Apply()
            End If

            'MsgBox(Me.AutoArchive.Checked)
            Dim __di As DirectoryInfo
            Try
                __di = New DirectoryInfo(__docRef.Path)
            Catch ex As Exception
                GoTo finish
            End Try


            'MsgBox(di.Name)
            If Me.AutoArchive.Checked And Not isExcludeDirectory(__di.Name) Then

                'Dim di As DirectoryInfo
                Dim __afi() As FileInfo
                Dim __fi As FileInfo

                'MsgBox(Directory.Exists(docRef.Path & "\" & Me.ArchiveDirectory.Text & "\"))
                'If Not Directory.Exists(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
                'Directory.CreateDirectory(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\")
                'End If

                'di = New DirectoryInfo(docRef.Path)
                Dim __currentFileName As String = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf("."))

                'Dim __RegexObj As Regex = New Regex("\d*$")
                Dim __RegexObj As Regex = New Regex("(\d*)\.*\d*$")
                Dim __myMatches As Match
                Dim __currentVersion
                Dim __cleanFileName As String

                If __RegexObj.IsMatch(__currentFileName) Then
                    __myMatches = __RegexObj.Match(__currentFileName)
                    __currentVersion = __myMatches.Value

                    If __currentVersion.length < 1 Then
                        GoTo finish
                    End If

                    __cleanFileName = __RegexObj.Replace(__currentFileName, "")
                    __cleanFileName = __cleanFileName.Replace("+", "\+")
                    __cleanFileName = __cleanFileName.Replace(" ", "\s")
                    'MsgBox(__cleanFileName)
                    Dim __RegexObj2 As Regex = New Regex("^" & __cleanFileName & "(\d+|\.)")

                    __afi = __di.GetFiles("*.*")
                    'MsgBox(__currentVersion)
                    For Each __fi In __afi
                        If __RegexObj2.IsMatch(__fi.Name) Then
                            'MsgBox(__fi.Name)
                            If isOldFileVersion(__fi.Name, __currentVersion) Then 'And Directory.Exists(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
                                'MsgBox(__fi.Name)
                                If Not Directory.Exists(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
                                    Directory.CreateDirectory(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\")
                                End If
                                Try
                                    File.Copy(__docRef.Path & __fi.Name, __docRef.Path & "\" & Me.ArchiveDirectory.Text & "\" & __fi.Name, True)
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                Finally
                                    File.Delete(__docRef.Path & __fi.Name)
                                End Try

                            End If
                        End If
                    Next
                End If


            End If
        Else 'Exporting each layers by name
            Dim __Layer
            For __compsIndex = 1 To __docRef.Layers.Count()
                __Layer = __docRef.Layers.Item(__compsIndex)
                'isVisible = oLayer.visible
                __appRef.ActiveDocument.ActiveLayer = __Layer
                'oLayer.Apply()
                'duppedDocument = docRef.Duplicate()
                'msgbox(compRef.Name)
                __fileNameBody = __Layer.Name & ".jpg"
                'msgbox(fileNameBody)
                __docRef.SaveAs(__docRef.Path & __fileNameBody, __jpgSaveOptions, True)
                __appRef.ActiveDocument.ActiveLayer.visible = False
                'duppedDocument.Close(2)
            Next
        End If

finish:
        'ExportImagesRights()
        If __compsCount > 0 Then
            __compRef = __docRef.LayerComps.Item(1)
            __compRef.Apply()
        End If
        If Not __stayOpen Then __docRef.Close(2)

        ' End program
        End
    End Sub

    Private Function isOldFileVersion(ByVal filename As String, ByVal version As String) As String
        Dim __newFileName As String = filename.Substring(0, filename.LastIndexOf("."))
        Dim __version = version.Split(".")
        'MsgBox(__newFileName)
        Dim __RegexObj As Regex = New Regex("(\d*)(\.*\d*)*$")
        If __RegexObj.IsMatch(__newFileName) Then
            'MsgBox("> " & __RegexObj.Match(__newFileName).Groups(1).Value & ":" & version)
            If CInt(__RegexObj.Match(__newFileName).Groups(1).Value) < CInt(__version(0)) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Private Function isExcludeDirectory(ByVal dirName As String)
        Dim __ExcludeDirectories As Array
        Dim __ExcludeDirectory As String

        If Me.ExcludeDirectories.Text <> "" Then
            __ExcludeDirectories = Split(Me.ExcludeDirectories.Text & ";" & Me.ArchiveDirectory.Text, ";")
            For Each __ExcludeDirectory In __ExcludeDirectories
                If __ExcludeDirectory = dirName Then
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
            If hasVersion("CS3") Then RegUninstall("10", "12")
            If hasVersion("CS4") Then RegUninstall("11", "13")
            If hasVersion("CS5") Then RegUninstall("12", "14")
            If hasVersion("CS55") Then RegUninstall("55", "15.1")
        Else
            If hasVersion("CS3") Then RegInstall("10", "12")
            If hasVersion("CS4") Then RegInstall("11", "13")
            If hasVersion("CS5") Then RegInstall("12", "14")
            If hasVersion("CS55") Then RegInstall("55", "15.1")
        End If
    End Sub

    Private Sub AutoArchive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoArchive.CheckedChanged
        Dim __newKey As RegistryKey
        __newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Me.AutoArchive.Checked Then
            __newKey.SetValue("AutoArchive", "1", RegistryValueKind.String)
        Else
            __newKey.SetValue("AutoArchive", "0", RegistryValueKind.String)
        End If
        __newKey.Close()
    End Sub

    Private Sub ArchiveDirectory_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchiveDirectory.TextChanged
        Dim __newKey As RegistryKey
        __newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Trim(Me.ArchiveDirectory.Text) <> "" Then
            __newKey.SetValue("ArchiveDirectory", Me.ArchiveDirectory.Text, RegistryValueKind.String)
        Else
            Me.ArchiveDirectory.Text = "Archives"
            __newKey.SetValue("ArchiveDirectory", "Archives", RegistryValueKind.String)
        End If
        __newKey.Close()
    End Sub

    Private Sub NamedExportQuality_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NamedExportQuality.ValueChanged
        Dim __newKey As RegistryKey
        __newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If __confLoaded Then
            'MessageBox.Show(Me.NamedExportQuality.Value.ToString())
            __newKey.SetValue("NamedExportQuality", Me.NamedExportQuality.Value, RegistryValueKind.String)
        End If
        __newKey.Close()
    End Sub

    Private Sub ExportLayerComps_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportLayerComps.CheckedChanged
        Dim __newKey As RegistryKey
        __newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Me.ExportLayerComps.Checked Then
            __newKey.SetValue("ExportLayerComps", "1", RegistryValueKind.String)
            __doExportLayerComps = True
        Else
            __newKey.SetValue("ExportLayerComps", "0", RegistryValueKind.String)
            __doExportLayerComps = False
        End If
        __newKey.Close()
    End Sub

    Private Sub ToolTip1_Popup(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub ExportLayerComps_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExportLayerComps.MouseHover
        ToolTip1.Show("Export Layer Comps if checked, else it will save each layers", ExportLayerComps)
    End Sub

    Private Sub ExcludeDirectories_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcludeDirectories.TextChanged
        Dim __newKey As RegistryKey
        __newKey = Registry.CurrentUser.CreateSubKey("Software\\SaveAsJPEG")
        If Trim(Me.ExcludeDirectories.Text) <> "" Then
            __newKey.SetValue("ExcludeDirectories", Me.ExcludeDirectories.Text, RegistryValueKind.String)
        Else
            Me.ExcludeDirectories.Text = ""
            __newKey.SetValue("ExcludeDirectories", "", RegistryValueKind.String)
        End If
        __newKey.Close()
    End Sub

    Private Sub OpenDocument()
        'MessageBox.Show("OpenDocument")
        __appRef = New Photoshop.Application()
        Try
            If __appRef.Documents.Count > 0 Then
                For i = 1 To __appRef.Documents.Count
                    If __args(2) = __appRef.Documents.Item(i).FullName Then
                        __appRef.ActiveDocument = __appRef.Documents.Item(i)
                        __docRef = __appRef.ActiveDocument
                        __openDoc = False
                        __stayOpen = True
                    End If
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        If __openDoc Then
            Try
                __docRef = __appRef.Open(__args(2))
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub ExportSmartObjects()
        Call OpenDocument()
        Call ProcessExportSmartObjects(__docRef.Layers)
        If Not __stayOpen Then __docRef.Close(2)
        End
    End Sub

    Private Sub ProcessExportSmartObjects(ByVal obj)
        'Dim __LayerRef
        Dim __Layer As Object
        Dim __isVisible As Boolean
        Dim __j As Integer

        For __j = 1 To obj.count
            __Layer = obj.Item(__j)
            __isVisible = __Layer.visible
            __appRef.ActiveDocument.ActiveLayer = __Layer
            'set oLayer = oLayerRef.ActiveLayer
            If __Layer.typename = "LayerSet" Then
                ProcessExportSmartObjects(__Layer.Layers)
            ElseIf __Layer.typename = "ArtLayer" Then
                If __Layer.Kind = Photoshop.PsLayerKind.psSmartObjectLayer Then

                    Dim __idplacedLayerExportContents
                    __idplacedLayerExportContents = __appRef.StringIDToTypeID("placedLayerExportContents")


                    Dim __desc4 As Photoshop.ActionDescriptor
                    __desc4 = New Photoshop.ActionDescriptor()

                    Dim __idnull
                    __idnull = __appRef.CharIDToTypeID("null")

                    If Not Directory.Exists(__docRef.Path & "+ Elements\") Then
                        Directory.CreateDirectory(__docRef.Path & "+ Elements\")
                    End If

                    Dim __soType = getSmartObjectType(__appRef)
                    If __soType <> "" Then
                        Call __desc4.putPath(__idnull, __docRef.Path & "+ Elements\" & wipeName(__Layer.Name) & __soType)
                        Call __appRef.ExecuteAction(__idplacedLayerExportContents, __desc4, Photoshop.PsDialogModes.psDisplayNoDialogs)
                    End If
                End If
            End If
            __appRef.ActiveDocument.ActiveLayer = __Layer
            __appRef.ActiveDocument.ActiveLayer.visible = __isVisible
        Next
    End Sub

    Function getSmartObjectType(ByVal appRef)
        Dim __aRef As Photoshop.ActionReference = New Photoshop.ActionReference()
        Call __aRef.PutEnumerated(appRef.charIDToTypeID("Lyr "), appRef.charIDToTypeID("Ordn"), appRef.charIDToTypeID("Trgt"))
        Dim __desc = appRef.ExecuteActionGet(__aRef)
        If __desc.hasKey(appRef.stringIDToTypeID("smartObject")) Then
            __desc = appRef.ExecuteActionGet(__aRef).GetObjectValue(appRef.stringIDToTypeID("smartObject")).GetEnumerationValue(appRef.stringIDToTypeID("placed"))
            Select Case appRef.typeIDToStringID(__desc)
                Case "vectorData"
                    getSmartObjectType = ".ai"
                Case "rasterizeContent"
                    getSmartObjectType = ".psd"
                Case Else
                    getSmartObjectType = ""
            End Select
        Else
            getSmartObjectType = ""
        End If
    End Function

    Function wipeName(ByVal value)
        value = Regex.Replace(value, "(\+)+\s", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = Regex.Replace(value, "\s(copy)\s(\d)*", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = Regex.Replace(value, "(vector)\s*|^v:", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = Regex.Replace(value, ":", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = Regex.Replace(value, "\s", "_", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        wipeName = value
    End Function

    Private Sub ExportImagesRights()
        Call OpenDocument()
        Call ProcessExportImagesRights(__docRef)
        If Not __stayOpen Then __docRef.Close(2)
        End
    End Sub

    Private Function ProcessExportImagesRights(ByVal __ActiveDocument)
        Dim __Layers
        Dim __Layer As Object
        Dim __isVisible As Boolean
        Dim __j As Integer
        Dim __ir As ImageRight
        Dim __soType As String

        __ir = New ImageRight()

        __Layers = __ActiveDocument.Layers

        For __j = 1 To __Layers.count
            __Layer = __Layers.Item(__j)
            __isVisible = __Layer.visible
            __appRef.ActiveDocument.ActiveLayer = __Layer
            'set oLayer = oLayerRef.ActiveLayer
            If __Layer.typename = "LayerSet" Then
                ProcessExportImagesRights(__Layer)
            ElseIf __Layer.typename = "ArtLayer" Then
                'MessageBox.Show(__Layer.Name)
                If __Layer.Kind = 1 Then
                    __ir.Parse(__Layer.Name)
                    If __ir.isValidURL Then
                        __ir.CreateLink(__docRef.Path)
                    End If
                ElseIf __Layer.Kind = Photoshop.PsLayerKind.psSmartObjectLayer Then
                    __ir.Parse(__Layer.Name)
                    If __ir.Code <> vbNullString Then
                        If __ir.isValidURL Then __ir.CreateLink(__docRef.Path)
                    Else
                        __soType = getSmartObjectType(__appRef)
                        If __soType = ".psd" Then

                            Dim __opn
                            __opn = __appRef.StringIDToTypeID("placedLayerEditContents")

                            Dim __desc4
                            __desc4 = New Photoshop.ActionDescriptor()

                            Try
                                __appRef.ExecuteAction(__opn, __desc4, Photoshop.PsDialogModes.psDisplayNoDialogs)
                            Catch ex As InvalidOperationException
                                MessageBox.Show(ex.Message)
                            End Try
                            ProcessExportImagesRights(__appRef.ActiveDocument)
                            __appRef.ActiveDocument.Close(2)
                        End If
                    End If
                End If
            End If
            __appRef.ActiveDocument.ActiveLayer = __Layer
            __appRef.ActiveDocument.ActiveLayer.visible = __isVisible
        Next
        ProcessExportImagesRights = True
    End Function

    Private Sub CleanLayersName()
        Dim __compsCount As Integer
        Dim __compRef As Photoshop.LayerComp

        Call OpenDocument()
        Call ProcessCleanLayersName(__docRef, 1)

        __compsCount = __docRef.LayerComps.Count
        If __compsCount > 0 Then
            __compRef = __docRef.LayerComps.Item(1)
            __compRef.Apply()
        End If
        __docRef.Save()
        If Not __stayOpen Then __docRef.Close(2)
        End
    End Sub

    Private Function ProcessCleanLayersName(ByVal __ActiveDocument, ByVal __idx) As String
        Dim __Layers As Photoshop.Layers
        Dim __Layer As Object
        Dim __isVisible As Boolean
        Dim __j As Integer
        Dim __ir As ImageRight
        Dim __xmlDoc As String
        Dim __FistLayerText As String = ""
        Dim __reg As New Regex("Group\s*\d*", RegexOptions.IgnoreCase)

        'Dim __soType As String

        __ir = New ImageRight()

        __Layers = __ActiveDocument.Layers

        For __j = 1 To __Layers.Count
            __Layer = __Layers.Item(__j)
            __isVisible = __Layer.visible
            __appRef.ActiveDocument.ActiveLayer = __Layer
            'set oLayer = oLayerRef.ActiveLayer
            If __Layer.typename = "LayerSet" Then
                __Layer.Name = New String("+", __idx) & " " & Regex.Replace(__Layer.Name, "(\+)+\s*", "")
                Dim __NewLayerName = ProcessCleanLayersName(__Layer, __idx + 1)
                If __NewLayerName <> "" And __reg.IsMatch(__Layer.name) Then
                    __Layer.name = New String("+", __idx) + " " + __NewLayerName
                End If
            ElseIf __Layer.typename = "ArtLayer" Then
                'MessageBox.Show(__Layer.Kind)
                __ir.Parse(__Layer.Name)
                If __ir.isValidCode Then
                    'MessageBox.Show(__Layer.Name)
                    __Layer.Name = "#" & Regex.Replace(__Layer.Name, "#", "")
                End If
                If __Layer.Kind = Photoshop.PsLayerKind.psSmartObjectLayer Then 'SMARTOBJECT
                    Try
                        __xmlDoc = __Layer.XMPMetadata.RawData
                        If __xmlDoc <> "" Then
                            __Layer.Name = New String("+", __idx) & " " & Regex.Replace(__Layer.Name, "(\+)+\s*", "")
                            ChangeLayerColour(Colors.VIOLET)
                        End If
                    Catch ex As Exception
                    End Try
                    '    __soType = getSmartObjectType(__appRef)
                    '    If __soType = ".psd" Then
                    '        'MessageBox.Show(__Layer.Name)
                    '        Dim __opn
                    '        __opn = __appRef.StringIDToTypeID("placedLayerEditContents")

                    '        Dim __desc4
                    '        __desc4 = New Photoshop.ActionDescriptor()

                    '        Try
                    '            __appRef.ExecuteAction(__opn, __desc4, 3)
                    '        Catch ex As InvalidOperationException
                    '            MessageBox.Show(ex.Message)
                    '        End Try
                    '        ProcessCleanLayersName(__appRef.ActiveDocument, 1)
                    '        __appRef.ActiveDocument.Close(1)
                    '    End If
                ElseIf __Layer.kind = Photoshop.PsLayerKind.psTextLayer Then
                    If __FistLayerText = "" Then
                        __FistLayerText = __Layer.name
                    End If
                End If
            End If
            __appRef.ActiveDocument.ActiveLayer = __Layer
            __appRef.ActiveDocument.ActiveLayer.visible = __isVisible
        Next
        ProcessCleanLayersName = __FistLayerText
    End Function

    Private Sub SaveScreenSelection()
        Call OpenDocument()
        Call ProcessSaveScreenSelection(__docRef)
        If Not __stayOpen Then __docRef.Close(2)
        End
    End Sub

    Private Sub ProcessSaveScreenSelection(ByVal __ActiveDocument)
        Dim __duppedDocument As Photoshop.Document
        Dim __selChannel As Photoshop.Channel
        Dim __SelBounds As Array
        Dim __jpgSaveOptions As Photoshop.JPEGSaveOptions
        Dim __fileNameBody

        __duppedDocument = __ActiveDocument.Duplicate()
        Try
            __selChannel = __duppedDocument.Channels.Item("screen")
            'MessageBox.Show(__selChannel.Name)
            __duppedDocument.Selection.Load(__selChannel)
            __SelBounds = __duppedDocument.Selection.Bounds
            __duppedDocument.Crop(__SelBounds)


            __jpgSaveOptions = New Photoshop.JPEGSaveOptions()
            __jpgSaveOptions.EmbedColorProfile = False
            __jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline 
            __jpgSaveOptions.Matte = 1 ' psNoMatte 
            __jpgSaveOptions.Quality = 12

            __fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) & "_screen.jpg"
            __duppedDocument.SaveAs(__ActiveDocument.Path & __fileNameBody, __jpgSaveOptions, True)
            __duppedDocument.Close(2)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ChangeLayerColour(ByVal colour As Colors)
        Dim __colour As String
        Dim __desc As Photoshop.ActionDescriptor
        Dim __ref As Photoshop.ActionReference
        Dim __desc2 As Photoshop.ActionDescriptor

        Select Case colour
            Case Colors.RED
                __colour = "Rd  "
            Case Colors.ORANGE
                __colour = "Orng"
            Case Colors.YELLOW
                __colour = "Ylw "
            Case Colors.GREEN
                __colour = "Grn "
            Case Colors.BLUE
                __colour = "Bl  "
            Case Colors.VIOLET
                __colour = "Vlt "
            Case Colors.GRAY
                __colour = "Gry "
            Case Colors.NONE
                __colour = "None"
            Case Else
                __colour = "None"
        End Select

        __desc = New Photoshop.ActionDescriptor()
        __ref = New Photoshop.ActionReference()
        __ref.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"))
        __desc.PutReference(__appRef.CharIDToTypeID("null"), __ref)

        __desc2 = New Photoshop.ActionDescriptor()
        __desc2.PutEnumerated(__appRef.CharIDToTypeID("Clr "), __appRef.CharIDToTypeID("Clr "), __appRef.CharIDToTypeID(__colour))
        __desc.PutObject(__appRef.CharIDToTypeID("T   "), __appRef.CharIDToTypeID("Lyr "), __desc2)
        __appRef.ExecuteAction(__appRef.CharIDToTypeID("setd"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs)
    End Sub

    Private Sub ExportBase64()
        'MessageBox.Show(__args(2))
        Dim __path As String
        Dim __fi As FileInfo
        Dim __fn As String
        Dim __f As File
        Dim __ext As String

        If __args(2).LastIndexOf("\") > -1 Then
            __path = __args(2).Substring(0, __args(2).LastIndexOf("\") + 1)
        End If


        __fi = New FileInfo(__args(2))
        __fn = __fi.Name.Substring(0, __fi.Name.Length - __fi.Extension.Length)
        __ext = __fi.Extension.Substring(1)
        'MessageBox.Show(__ext)

        Dim __bytes As Byte() = File.ReadAllBytes(__args(2))

        Dim __b64String = Convert.ToBase64String(__bytes)
        Dim __dataUrl As String = "<html><body><img src=""data:image/" & __ext & ";base64," & __b64String & """/></body></html>"

        '__f = New File(__fn & ".txt")
        __f.WriteAllText(__path & __fn & ".html", __dataUrl)
    End Sub
End Class