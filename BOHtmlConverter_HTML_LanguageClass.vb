Imports System.Xml
Imports System.Text
Imports System.IO
Imports System.Text.RegularExpressions

Public Class LanguageClass
    Public Function depart(input As String, lang As TextLang) As String
        Dim xdoc As XmlDocument = New XmlDocument
        ' load the xml 
        xdoc.LoadXml(input)
        ' get cell list
        Dim nodeList As XmlNodeList = xdoc.GetElementsByTagName("Cell")
        For Each c As XmlNode In nodeList
            'Dim paraList As XmlNodeList = c.SelectNodes("Paragraph")
            Dim paraList As ArrayList = getParagraphNode(c)
            Dim multiLangResult As Hashtable = Nothing
            ' multi paragraphs in one cell
            If (paraList.Count > 1) Then
                multiLangResult = markLanguage(paraList)
            End If
            If (multiLangResult IsNot Nothing) Then
                ' get language count
                Dim multiHash As New Hashtable
                For Each v As TextLang In multiLangResult.Values
                    If (multiHash(v) Is Nothing) Then
                        multiHash.Add(v, v)
                    End If
                Next
                ' multi language in one cell, do the depart
                If (multiHash.Keys.Count > 1) Then
                    For Each para As XmlNode In multiLangResult.Keys
                        If (multiLangResult(para) = lang) Then
                            'c.RemoveChild(para)
                            ' self delete
                            para.ParentNode.RemoveChild(para)
                        End If
                    Next
                End If
            End If
        Next
        Dim output As String = String.Empty
        Dim ms As MemoryStream = New MemoryStream()
        Dim xtw As New XmlTextWriter(ms, Encoding.UTF8)
        'Try
        xtw.Formatting = Formatting.Indented
        xdoc.WriteContentTo(xtw)
        xtw.Flush()
        ms.Seek(0, SeekOrigin.Begin)
        Dim sr As New StreamReader(ms)
        output = sr.ReadToEnd
        'Catch ex As Exception
        'Console.Write(ex.Message)
        'End Try
        Return output
    End Function

    Private Function getParagraphNode(ByRef n As XmlNode) As ArrayList
        Dim result As New ArrayList
        For Each node As XmlNode In n.SelectNodes(".//Paragraph")
            result.Add(node)
        Next
        Return result
    End Function

    ' check if the paragraph has multi-language
    ' used in translation, for example, 
    ' ------------------
    ' 你好嗎
    ' how are you
    ' ------------------
    ' it result in
    ' ------------------
    ' 你好嗎               CN
    ' how are you         PT 
    ' ------------------
    ' 你好嗎,你好嗎
    ' 你好嗎,how are you
    ' ------------------
    ' it result in
    ' ------------------
    ' 你好嗎,你好嗎          CN
    ' 你好嗎,how are you     PT
    ' ------------------
    Public Function markLanguage(ByRef paraList As ArrayList) As Hashtable
        Dim result As New Hashtable
        Dim chnRegex As New Regex("[\u4E00-\u9FA5]+")
        For Each para As XmlNode In paraList
            Dim txt As String = para.InnerText
            If (chnRegex.Match(txt).Success) Then
                If chnRegex.Match(txt).Length = txt.Length Then
                    ' pure chinese
                    result.Add(para, TextLang.CN)
                    Console.WriteLine("CHN: " & txt)
                    Continue For
                End If
                ' chinese count
                Dim chnCount As Integer = 0
                ' get the match
                Dim chnMatch As Match = chnRegex.Match(txt)
                ' count the total match length
                While (chnMatch.Success)
                    chnCount += chnMatch.Length
                    chnMatch = chnMatch.NextMatch()
                End While
                ' remove the space to uncount the space
                txt = txt.Replace(" ", "")
                ' if the chinese count > half of total length, than cn else pt
                If (chnCount > txt.Length / 2) Then
                    result.Add(para, TextLang.CN)
                Else
                    result.Add(para, TextLang.PT)
                End If
            Else
                ' don't mark empty string
                If (txt.Trim = String.Empty) Then
                    Continue For
                End If
                ' pure pt
                result.Add(para, TextLang.PT)
            End If
        Next
        Return result
    End Function
End Class
