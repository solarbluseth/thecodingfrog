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
        hwnd = FindWindow(vbNullString, strTitle)
        ' hide the app
        ShowWindow(hwnd, SW_HIDE)
        'On Error Resume Next

        Dim strRoot As String = "\\Photoshop.Image.11\\shell\\Export Smart Objects\\command"

        Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey("Photoshop.Image.11\\shell\\Export Smart Objects\\command\")

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
                If args(1) = appRef.Documents.Item(i).FullName Then
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

        compsCount = docRef.LayerComps.Count
        checkLayer(docRef.Layers)

        Dim sOptions As Photoshop.PsSaveOptions
        'sOptions.psDoNotSaveChanges()
        If Not stayOpen Then docRef.Close(sOptions.psDoNotSaveChanges)

    End Sub

    Sub checkLayer(ByVal obj)
        Dim oLayerRef
        Dim oLayer
        Dim isVisible As Boolean
        Dim j As Integer
        For j = 1 To obj.count
            oLayer = obj.Item(j)
            isVisible = oLayer.visible
            appRef.ActiveDocument.ActiveLayer = oLayer
            'set oLayer = oLayerRef.ActiveLayer
            If oLayer.typename = "LayerSet" Then
                checkLayer(oLayer.Layers)
            ElseIf oLayer.typename = "ArtLayer" Then
                If oLayer.Kind = 17 Then

                    Dim idplacedLayerExportContents
                    idplacedLayerExportContents = appRef.stringIDToTypeID("placedLayerExportContents")


                    Dim desc4
                    desc4 = New Photoshop.ActionDescriptor()

                    Dim idnull
                    idnull = appRef.charIDToTypeID("null")

                    If Not Directory.Exists(docRef.Path & "Exports\") Then
                        Directory.CreateDirectory(docRef.Path & "Exports\")
                    End If

                    Dim soType = getSmartObjectType(appRef)
                    If soType <> "" Then
                        Call desc4.putPath(idnull, docRef.Path & "Exports\" & wipeName(oLayer.Name) & soType)
                        Call appRef.ExecuteAction(idplacedLayerExportContents, desc4, 3)
                    End If
                End If
            End If
            appRef.ActiveDocument.ActiveLayer.visible = isVisible
        Next
    End Sub

    Function getSmartObjectType(ByVal appRef)
        Dim aRef As Photoshop.ActionReference = New Photoshop.ActionReference()
        Call aRef.PutEnumerated(appRef.charIDToTypeID("Lyr "), appRef.charIDToTypeID("Ordn"), appRef.charIDToTypeID("Trgt"))
        Dim desc = appRef.ExecuteActionGet(aRef)
        If desc.hasKey(appRef.stringIDToTypeID("smartObject")) Then
            desc = appRef.ExecuteActionGet(aRef).GetObjectValue(appRef.stringIDToTypeID("smartObject")).GetEnumerationValue(appRef.stringIDToTypeID("placed"))
            Select Case appRef.typeIDToStringID(desc)
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
