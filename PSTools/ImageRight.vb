Public Class ImageRight
    Private __bank As String
    Private __bankcode As String
    Private __imagecode As String
    Private __url As String

    Public Sub Parse(ByVal __name As String)
        Dim __mc As MatchCollection
        Dim __gc As GroupCollection
        Dim __cc As CaptureCollection

        __mc = Regex.Matches(__name, "(\w{2})\-\w{2}\-(DA|€€)\-(.*)", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        If __mc.Count > 0 Then
            'MessageBox.Show(__name)
            __gc = __mc.Item(0).Groups

            __cc = __gc.Item(1).Captures
            __bankcode = __cc(0).Value

            __cc = __gc.Item(3).Captures
            __imagecode = __cc(0).Value

            Call setURL()
        Else
            __bank = vbNullString
            __imagecode = vbNullString
        End If
    End Sub

    Private Sub setURL()
        Select Case __bankcode
            Case "GI" : __url = "www.gettyimages.fr/detail/{0}"
                __bank = "Getty Images"
            Case "IS" : __url = "www.istockphoto.com/file_thumbview_approve/{0}"
                __bank = "iStockPhoto"
            Case "CO" : __url = "www.corbisimages.com/stock-photo/rights-managed/{0}/e/?tab=details&caller=search"
                __bank = "Corbis Images"
            Case "FK" : __url = vbNullString
                __bank = vbNullString
            Case Else
                __bank = vbNullString
                __imagecode = vbNullString
                __url = vbNullString
        End Select
    End Sub

    Public Sub CreateLink(ByVal __path As String)
        Dim __sw As StreamWriter

        If Not Directory.Exists(__path & "+ Rights\") Then
            Directory.CreateDirectory(__path & "+ Rights\")
        End If
        __sw = File.CreateText(__path & "+ Rights\" + __bank + " " + __imagecode + ".url")
        __sw.WriteLine("[InternetShortcut]")
        __sw.WriteLine("URL=http://" + String.Format(__url, __imagecode))
        __sw.Close()
    End Sub


    Public ReadOnly Property Code()
        Get
            Return __imagecode
        End Get
    End Property

    Public ReadOnly Property ImageBank()
        Get
            Return __bank
        End Get
    End Property

    Public ReadOnly Property URL()
        Get
            Return __url
        End Get
    End Property

    Public ReadOnly Property isValidURL()
        Get
            Return (__url <> vbNullString And __imagecode <> vbNullString)
        End Get
    End Property

End Class
