
Public Class XmlParser


    Public Function LoadXml(ByVal strFormName) As String

        Dim doc As New System.Xml.XmlDocument
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder("")
        Dim labels As System.Xml.XmlNodeList
        Dim strXPath As String = "/labels/webforms/webform"

        doc.Load("C:\webpub\plv2\labels.xml")

        labels = doc.SelectNodes(strXPath)

        For Each webform As System.Xml.XmlNode In labels

            sb.Append("<hr>Web Form Name: " & webform.Attributes.GetNamedItem("name").Value & "<br>")

            If webform.Attributes.GetNamedItem("name").Value = strFormName Then

                For Each fields As System.Xml.XmlNode In webform.ChildNodes

                    For Each field As System.Xml.XmlNode In fields.ChildNodes

                        sb.Append(field.Name & ": " & field.InnerText & "<br>")

                    Next

                Next

            End If

        Next

        Return sb.ToString

    End Function

End Class
