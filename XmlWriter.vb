Imports System.Xml
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web


Namespace ParaLideres

    Public Class XmlWriter

        Private _strFilePath As String = HttpContext.Current.Request.PhysicalApplicationPath & "\Editor\labels.xml"

        Public Sub CreateXmlFile()

            Dim enc As Encoding
            Dim objXmlWriter As XmlTextWriter
            Dim reader As System.Data.SqlClient.SqlDataReader
            Dim rdFld As System.Data.SqlClient.SqlDataReader
            Dim strTableName As String = ""
            Dim strSQL As String = ""
            Dim strFldName As String = ""
            Dim strFldLabel As String = ""
            Dim strFldInst As String = ""

            Try

                'Create file, overwrite if exists
                'enc is encoding object required by constructor
                'It is null, so default encoding is used
                objXmlWriter = New XmlTextWriter(_strFilePath, enc)

                objXmlWriter.Formatting = objXmlWriter.Formatting.Indented


                objXmlWriter.WriteStartDocument()


                'xml file structure
                '<labels>
                '<webforms>
                '<webform name="Registration">
                '<field 
                'name = "FirstName"
                'Label = "Primer Nombre"
                'instructions = "ingresa tu primer nombre"
                '/>
                '<field 
                'name = "LastName"
                'label = "Apellido(s)"
                'instructions = "Ingresa tu apellido"
                '/>
                '</webform>
                '</webforms>
                '</labels>

                'Top level <LABELS>
                objXmlWriter.WriteStartElement("labels")

                '<WEBFORMS>
                objXmlWriter.WriteStartElement("webforms")

                Try

                    reader = ParaLideres.GenericDataHandler.GetRecords("SELECT name FROM sysobjects WHERE (parent_obj = 0)AND (type = 'U') AND name <> 'dtproperties' ORDER BY name")

                    If reader.HasRows() Then

                        Do While (reader.Read())

                            If Not reader.IsDBNull(0) Then strTableName = reader(0)

                            '<WEBFORM>
                            objXmlWriter.WriteStartElement("webform")
                            objXmlWriter.WriteAttributeString("name", "frm" & Functions.camelNotation(strTableName, True))

                            strSQL = "proc_GetTableStructure '" & strTableName & "'"

                            rdFld = ParaLideres.GenericDataHandler.GetRecords(strSQL)

                            If rdFld.HasRows() Then

                                Do While (rdFld.Read())

                                    '<FIELD>
                                    objXmlWriter.WriteStartElement("field")

                                    strFldName = rdFld(0)

                                    objXmlWriter.WriteAttributeString("name", Functions.camelNotation(strFldName, True))

                                    strFldName = Functions.CreateFormLabel(strFldName)

                                    objXmlWriter.WriteAttributeString("label", strFldName)
                                    objXmlWriter.WriteAttributeString("instructions", "Insert a value for " & strFldName)

                                    objXmlWriter.WriteEndElement()
                                    '</FIELD>

                                Loop

                            End If

                            objXmlWriter.WriteEndElement()
                            '</WEBFORM>

                        Loop

                    End If

                Catch ex As Exception

                    Throw ex

                Finally

                    reader.Close()
                    reader = Nothing

                    rdFld.Close()
                    rdFld = Nothing

                End Try

                objXmlWriter.WriteEndElement()
                '</WEBFORMS>

                objXmlWriter.WriteEndDocument()
                '</LABELS>

                objXmlWriter.Flush() 'Write to file

                objXmlWriter.Close()

            Catch Ex As Exception

                Throw Ex

            Finally

                objXmlWriter = Nothing

            End Try

        End Sub

    End Class

End Namespace