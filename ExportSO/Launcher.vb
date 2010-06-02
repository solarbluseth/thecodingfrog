Imports Microsoft.Win32
Imports System.Text.RegularExpressions
Imports System.IO
Imports Photoshop.PsSaveOptions

Module Launcher

    Public Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Integer) As Integer
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer

    Const SW_HIDE = 0
    Const SW_SHOWNORMAL = 1
    Const SW_NORMAL = 1
    Const SW_SHOWMINIMIZED = 2
    Dim appRef As Photoshop.Application
    Dim docRef As Photoshop.Document
    Dim openDoc As Boolean = True
    Dim stayOpen As Boolean = False

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

        Dim strRoot As String = "\\Photoshop.Image.11\\shell\\Export Smart Objects\\command"

        Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Export Smart Objects\\command\")

        'strRead = Registry.ClassesRoot.OpenSubKey(strRoot, True)

        'MsgBox(Err.Number)
        'MsgBox(Err.Description)
        If key Is Nothing Then
            Dim newKey As RegistryKey
            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image.11\\shell\\Export Smart Objects\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""%1""", RegistryValueKind.String)
            newKey.Close()

            newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile.11\\shell\\Export Smart Objects\\command")
            newKey.SetValue("", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """ ""%1""", RegistryValueKind.String)
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
        If num < 1 Then
            MsgBox("Utilisez le menu contextuel sur les fichiers .PSD, .PSB")
            End
        End If

        appRef = New Photoshop.Application()
        
        Dim i As Integer

        On Error Resume Next
        If appRef.Documents.Count > 0 Then
            For i = 1 To appRef.Documents.Count
                'msgbox(appRef.Documents.Item(i).FullName)
                If args(1) = appRef.Documents.Item(i).FullName Then
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
            docRef = appRef.Open(args(1))
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
        checkLayer(docRef.Layers)

        Dim sOptions As Photoshop.PsSaveOptions
        'sOptions.psDoNotSaveChanges()
        If Not stayOpen Then docRef.Close(sOptions.psDoNotSaveChanges)

    End Sub

    Sub checkLayer(ByVal obj)
        'Dim fol
        Dim oLayerRef
        Dim oLayer
        'Dim fso
        Dim isVisible
        'fso = CreateObject("Scripting.FileSystemObject")
        Dim j
        For j = 1 To obj.count
            oLayer = obj.Item(j)
            isVisible = oLayer.visible
            appRef.ActiveDocument.ActiveLayer = oLayer
            'set oLayer = oLayerRef.ActiveLayer
            If oLayer.typename = "LayerSet" Then
                checkLayer(oLayer.Layers)
            ElseIf oLayer.typename = "ArtLayer" Then
                If oLayer.Kind = 17 Then

                    'msgbox(oLayer.typename)
                    Dim idplacedLayerExportContents
                    idplacedLayerExportContents = appRef.stringIDToTypeID("placedLayerExportContents")


                    Dim desc4
                    desc4 = New Photoshop.ActionDescriptor()

                    Dim idnull
                    idnull = appRef.charIDToTypeID("null")

                    'msgbox(fso.FolderExists(docRef.Path & "exports\"))
                    If Not Directory.Exists(docRef.Path & "Exports\") Then
                        Directory.CreateDirectory(docRef.Path & "Exports\")
                    End If
                    'MsgBox(oLayer.Name)
                    If isExportable(oLayer.Name) Then
                        If isIllustrator(oLayer.Name) Then
                            Call desc4.putPath(idnull, docRef.Path & "Exports\" & wipeName(oLayer.Name) & ".ai")
                        Else
                            'MsgBox(wipeName(oLayer.Name))
                            Call desc4.putPath(idnull, docRef.Path & "Exports\" & wipeName(oLayer.Name) & ".psd")
                        End If
                        Call appRef.ExecuteAction(idplacedLayerExportContents, desc4, 3)
                    End If
                End If
            End If
            appRef.ActiveDocument.ActiveLayer.visible = isVisible
        Next
    End Sub

    Function isIllustrator(ByVal value)
        Dim myRegExp As Regex
        isIllustrator = myRegExp.IsMatch(value, "^(v:|vector)", RegexOptions.Multiline And RegexOptions.IgnoreCase)
    End Function

    Function isExportable(ByVal value)
        Dim myRegExp As Regex
        isExportable = Not myRegExp.IsMatch(value, "^(layer|vector)", RegexOptions.Multiline And RegexOptions.IgnoreCase)
    End Function

    Function wipeName(ByVal value)
        Dim myRegExp As Regex

        value = myRegExp.Replace(value, "(\+)+\s", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = myRegExp.Replace(value, "\s(copy)\s(\d)*", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = myRegExp.Replace(value, "(vector)\s*|^v:", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = myRegExp.Replace(value, ":", "", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        value = myRegExp.Replace(value, "\s", "_", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        wipeName = value
    End Function

    Sub showErr(ByVal errType)
        Select Case errType
            Case "opened"
                MsgBox("Photoshop est en mode saisie sur un layer !")
                End
        End Select
    End Sub

End Module
