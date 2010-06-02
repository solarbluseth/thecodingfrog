Imports Microsoft.Win32

Module Launcher

    Public Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Integer) As Integer
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer

    Const SW_HIDE = 0
    Const SW_SHOWNORMAL = 1
    Const SW_NORMAL = 1
    Const SW_SHOWMINIMIZED = 2

    Sub Main()

        Dim strTitle As String
        Dim rtnLen As Long
        Dim hwnd As Int32

        strTitle = Space(256)
        rtnLen = GetConsoleTitle(strTitle, 256)
        If rtnLen > 0 Then
            strTitle = Left$(strTitle, rtnLen)
        End If
        'MsgBox(strTitle)
        hwnd = FindWindow(vbNullString, strTitle)
        ' hide the app
        ShowWindow(hwnd, SW_HIDE)



        'On Error Resume Next
        Dim strRoot As String = "\\Photoshop.Image.11\\shell\\Save as JPEG 100%\\command"

        Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Save as JPEG 100%\\command\")

        'strRead = Registry.ClassesRoot.OpenSubKey(strRoot, True)

        'MsgBox(Err.Number)
        'MsgBox(Err.Description)
        If key Is Nothing Then
            Dim newKey As RegistryKey
            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image.11\\shell\\Save as JPEG 100%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image.11\\shell\\Save as JPEG 60%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile.11\\shell\\Save as JPEG 100%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile.11\\shell\\Save as JPEG 60%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 100%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""12"" ""%1""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\shell\\Save as JPEG 60%\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""7"" ""%1""", RegistryValueKind.String)
            newKey.Close()

            MsgBox("Script ajouté au registre !")
            End
        End If
        key.Close()
        On Error GoTo 0

        Dim args As String() = Environment.GetCommandLineArgs()
        'MsgBox(args.Length)
        'End
        Dim num As Integer = UBound(args)
        If num < 2 Then
            MsgBox("Utilisez le menu contextuel sur les fichiers .PSD, .PSB, .EPS")
            End
        End If

        Dim appRef As Photoshop.Application = New Photoshop.Application()
        Dim docRef As Photoshop.Document
        Dim openDoc As Boolean = True
        Dim stayOpen As Boolean = False
        Dim i As Integer

        On Error Resume Next
        If appRef.Documents.Count > 0 Then
            For i = 1 To appRef.Documents.Count
                'msgbox(appRef.Documents.Item(i).FullName)
                If args(2) = appRef.Documents.Item(i).FullName Then
                    appRef.ActiveDocument = appRef.Documents.Item(i)
                    docRef = appRef.ActiveDocument
                    openDoc = False
                    stayOpen = True
                End If
            Next
        End If
        If Err.Number <> 0 Then
            'MsgBox("koin")
            Call showErr("opened")
        End If
        On Error GoTo 0

        If openDoc Then
            On Error Resume Next
            'MsgBox(args(2))
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
        'MsgBox(docRef.Name)
        compsCount = docRef.LayerComps.Count
        'msgbox(compsCount)
        Dim jpgSaveOptions As Photoshop.JPEGSaveOptions = New Photoshop.JPEGSaveOptions
        jpgSaveOptions.EmbedColorProfile = False
        jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline 
        jpgSaveOptions.Matte = 1 ' psNoMatte 
        jpgSaveOptions.Quality = CInt(args(1))
        If compsCount <= 1 Then
            'Set textItemRef = appRef.ActiveDocument.Layers(1) 

            'textItemRef.TextItem.Contents = Args.Item(1) 

            'outFileName = Args.Item(1)
            docRef.SaveAs(args(2), jpgSaveOptions, True)
        Else
            'msgbox("comps!")
            For compsIndex = 1 To compsCount
                'msgbox(docRef.LayerComps.Count)
                compRef = docRef.LayerComps.Item(compsIndex)
                'if (exportInfo.selectionOnly && !compRef.selected) continue; // selected only
                compRef.Apply()
                duppedDocument = docRef.Duplicate()
                'msgbox(docRef.Path)
                fileNameBody = Split(docRef.Name, ".")(0) & "." & compsIndex & ".jpg"
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
        If Not stayOpen Then docRef.Close(2)
    End Sub

    Sub showErr(ByVal errType)
        Select Case errType
            Case "opened"
                MsgBox("Photoshop est en mode saisie sur un layer !")
                End
        End Select
    End Sub
End Module
