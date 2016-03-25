Imports System.Web

Namespace ParaLideres

    Public Class ImageResize

        Public Shared Sub ResizeImage(ByVal strImagePath As String, ByVal intWidth As Integer, ByVal intHeight As Integer, Optional ByVal strSaveAsImagePath As String = "")

            Dim fs As System.IO.FileStream
            Dim img As System.Drawing.Image

            Try

                If System.IO.File.Exists(strImagePath) Then

                    ' select the format of the image to write according to the current extension
                    Dim imgFormat As System.Drawing.Imaging.ImageFormat = GetImageFormat(strImagePath)

                    fs = System.IO.File.OpenRead(strImagePath)

                    HttpContext.Current.Trace.Write("fs opened")

                    'System.Threading.Thread.Sleep(5000)

                    img = System.Drawing.Image.FromStream(fs)

                    HttpContext.Current.Trace.Write("initialized img")

                    'System.Threading.Thread.Sleep(5000)

                    Try

                        If (img.Height > intHeight) And (img.Width > intWidth) Then

                            If intHeight = 0 Then intHeight = img.Height / (img.Width / intWidth)

                            If intWidth = 0 Then intWidth = img.Width / (img.Height / intHeight)

                            ' create a new empty bitmpat with the specified size
                            Dim bmp As New System.Drawing.Bitmap(intWidth, intHeight)

                            HttpContext.Current.Trace.Write("initialized bmp")

                            ' retrieve a canvas object that allows to draw on the empty bitmap
                            Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(DirectCast(bmp, System.Drawing.Image))

                            HttpContext.Current.Trace.Write("initialized g")

                            ' copy the original image on the canvas, and thus on the new bitmap,
                            '  with the new size
                            g.DrawImage(img, 0, 0, intWidth, intHeight)

                            HttpContext.Current.Trace.Write("g.DrawImage")

                            ' close the original image

                            fs.Close()
                            fs = Nothing

                            img.Dispose()
                            img = Nothing

                            HttpContext.Current.Trace.Write("img.Dispose()")

                            ' save the new image with the proper format

                            If strSaveAsImagePath <> "" Then

                                HttpContext.Current.Trace.Write("bmp.Save: " & strSaveAsImagePath)

                                Try

                                    bmp.Save(strSaveAsImagePath, imgFormat)

                                Catch ex As Exception

                                    HttpContext.Current.Trace.Warn(ex.ToString())

                                    Throw ex

                                End Try

                            Else

                                HttpContext.Current.Trace.Write("bmp.Save: " & strImagePath)

                                Try

                                    System.Threading.Thread.Sleep(5000)

                                    bmp.Save(strImagePath, imgFormat)

                                Catch ex As Exception

                                    HttpContext.Current.Trace.Warn(ex.ToString())

                                    Throw ex

                                End Try

                            End If

                            bmp.Dispose()

                            HttpContext.Current.Trace.Write("bmp.Dispose()")

                            If Not IsNothing(bmp) Then bmp = Nothing

                            If Not IsNothing(g) Then g = Nothing

                        End If 'If img.Height > intHeight And img.Width > intWidth Then

                    Catch ex As Exception

                        Throw

                    Finally

                        If Not IsNothing(imgFormat) Then imgFormat = Nothing

                        If Not IsNothing(img) Then img = Nothing

                    End Try

                End If 'If objFile.Exists(strImagePath) Then

            Catch ex As Exception

                Throw ex

            Finally


            End Try

        End Sub


        Public Shared Function GetImageFormat(ByVal imgPath As String) As System.Drawing.imaging.ImageFormat

            Try

                imgPath = imgPath.ToLower()

                If imgPath.EndsWith(".bmp") Then

                    Return System.Drawing.Imaging.ImageFormat.Bmp

                ElseIf imgPath.EndsWith(".emf") Then

                    Return System.Drawing.Imaging.ImageFormat.Emf

                ElseIf imgPath.EndsWith(".gif") Then

                    Return System.Drawing.Imaging.ImageFormat.Gif

                ElseIf imgPath.EndsWith(".ico") Then

                    Return System.Drawing.Imaging.ImageFormat.Icon

                ElseIf imgPath.EndsWith(".jpg") OrElse imgPath.EndsWith(".jpeg") Then

                    Return System.Drawing.Imaging.ImageFormat.Jpeg

                ElseIf imgPath.EndsWith(".png") Then

                    Return System.Drawing.Imaging.ImageFormat.Png

                ElseIf imgPath.EndsWith(".tif") OrElse imgPath.EndsWith(".tiff") Then

                    Return System.Drawing.Imaging.ImageFormat.Tiff

                ElseIf imgPath.EndsWith(".wmf") Then

                    Return System.Drawing.Imaging.ImageFormat.Wmf

                Else

                    Return System.Drawing.Imaging.ImageFormat.Gif

                End If

            Catch ex As Exception

                Throw ex

            End Try

        End Function


    End Class

End Namespace