Imports System.Reflection
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Globalization
Public Class Unpacker
    Dim s_password As String
    Dim s_name As String
    Dim s_folder As String
    Public s_length As New List(Of String)
    Public s_file As New List(Of String)
    Public s_clickon
    Public s_clicktimes
    Function GetDumpedPassword() As String
        Return s_password
    End Function

    Function GetLastOutputFolder() As String
        Return s_folder
    End Function

    Function GetLastOutputSize(ByVal value As Integer) As String
        Return s_length(value)
    End Function

    Function GetLastFileName(ByVal value As Integer) As String
        Return s_file(value)
    End Function

    Function GetDumpedName() As String
        Return s_name
    End Function
    Dim sRtf
    Public Function InitUnpack(ByVal packedEXE As String) As Integer
        GetNameAndPassword(packedEXE)
        UnpackFiles(sRtf)
        Return 1
    End Function
    Public rtfnotused = False
    Public Sub UnpackFiles(ByVal sRtf As String, Optional ByVal onlyGetSize As Boolean = False)
        Dim Path = My.Application.Info.DirectoryPath & "\4dotsEXEDocumentPacker_Extracted\" & IO.Path.GetFileNameWithoutExtension(s_name) & "\"
        If onlyGetSize = True Then
            s_length.Add(GetSizeInMemory(System.Text.Encoding.Unicode.GetBytes(sRtf).LongLength))
            Return
        Else
            If Directory.Exists(My.Application.Info.DirectoryPath & "\4dotsEXEDocumentPacker_Extracted") = False Then
                Directory.CreateDirectory(My.Application.Info.DirectoryPath & "\4dotsEXEDocumentPacker_Extracted")
            End If
            Directory.CreateDirectory(Path)
            rtfnotused = False
            Dim rtf_temp As New RichTextBox
            Try
                rtf_temp.Rtf = sRtf
            Catch ex As Exception
                rtfnotused = True
                rtf_temp.Text = sRtf
            End Try
            File.WriteAllLines(Path & s_name, rtf_temp.Lines)
            s_folder = Path
        End If
    End Sub

    Public Function GetSizeInMemory(ByVal size As Long) As String
        Dim sizes() As String = {"B", "KB", "MB", "GB", "TB"}

        Dim len As Double = Convert.ToDouble(size)
        Dim order As Integer = 0
        While len >= 1024D And order < sizes.Length - 1
            order += 1
            len /= 1024
        End While
        Return String.Format(CultureInfo.CurrentCulture, "{0:0.##} {1}", len, sizes(order))
    End Function

    Public Sub GetNameAndPassword(ByVal packedEXEPath As String)
        Dim executingAssembly As Assembly = Assembly.LoadFile(packedEXEPath)
        Dim manifestResourceNames As String() = executingAssembly.GetManifestResourceNames()
        For i As Integer = 0 To manifestResourceNames.Length - 1
            If manifestResourceNames(i).IndexOf("LockedDocument.rtf") >= 0 Then
                Using binaryReader As BinaryReader = New BinaryReader(executingAssembly.GetManifestResourceStream(manifestResourceNames(i)))
                    Using memoryStream As MemoryStream = New MemoryStream()
                        While True
                            Dim num As Long = 32768L
                            Dim buffer As Byte() = New Byte(num - 1) {}
                            Dim num2 As Integer = binaryReader.Read(buffer, 0, CInt(num))
                            If num2 <= 0 Then
                                Exit While
                            End If
                            memoryStream.Write(buffer, 0, num2)
                        End While
                        sRtf = Encoding.[Default].GetString(memoryStream.ToArray())
                    End Using
                End Using
            ElseIf manifestResourceNames(i).IndexOf("project.xml") >= 0 Then
                Using binaryReader As BinaryReader = New BinaryReader(executingAssembly.GetManifestResourceStream(manifestResourceNames(i)))
                    Using memoryStream As MemoryStream = New MemoryStream()
                        While True
                            Dim num As Long = 32768L
                            Dim buffer As Byte() = New Byte(num - 1) {}
                            Dim num2 As Integer = binaryReader.Read(buffer, 0, CInt(num))
                            If num2 <= 0 Then
                                Exit While
                            End If
                            memoryStream.Write(buffer, 0, num2)
                        End While
                        Dim text As String = Encoding.[Default].GetString(memoryStream.ToArray())
                        text = DecryptString(text, "4dotsSoftware012301230123")
                        Dim xmlDocument As XmlDocument = New XmlDocument()
                        xmlDocument.LoadXml(text)
                        Dim xmlNode As XmlNode = xmlDocument.SelectSingleNode("//Project")
                        s_clicktimes = Integer.Parse(xmlNode.Attributes.GetNamedItem("ClickTimes").Value)
                        s_clickon = Integer.Parse(xmlNode.Attributes.GetNamedItem("ClickOn").Value)
                        s_name = xmlNode.Attributes.GetNamedItem("Name").Value
                        s_password = xmlNode.Attributes.GetNamedItem("Password").Value
                        s_file.Add(s_name)
                    End Using
                End Using
            End If
        Next
        sRtf = DecryptString(sRtf, s_password)
    End Sub
    Public Shared Function DecryptBytes(ByVal Message As String, ByVal Passphrase As String) As Byte()
        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
        Dim md5CryptoServiceProvider As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim key As Byte() = md5CryptoServiceProvider.ComputeHash(utf8Encoding.GetBytes(Passphrase))
        Dim tripleDESCryptoServiceProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDESCryptoServiceProvider.Key = key
        tripleDESCryptoServiceProvider.Mode = CipherMode.ECB
        tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7
        Dim array As Byte() = Convert.FromBase64String(Message)
        Dim bytes As Byte()
        Try
            Dim cryptoTransform As ICryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor()
            bytes = cryptoTransform.TransformFinalBlock(array, 0, array.Length)
        Finally
            tripleDESCryptoServiceProvider.Clear()
            md5CryptoServiceProvider.Clear()
        End Try
        Return bytes
    End Function

    Public Shared Function DecryptString(ByVal Message As String, ByVal Passphrase As String) As String
        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
        Dim md5CryptoServiceProvider As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim key As Byte() = md5CryptoServiceProvider.ComputeHash(utf8Encoding.GetBytes(Passphrase))
        Dim tripleDESCryptoServiceProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDESCryptoServiceProvider.Key = key
        tripleDESCryptoServiceProvider.Mode = CipherMode.ECB
        tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7
        Dim array As Byte() = Convert.FromBase64String(Message)
        Dim bytes As Byte()
        Try
            Dim cryptoTransform As ICryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor()
            bytes = cryptoTransform.TransformFinalBlock(array, 0, array.Length)
        Finally
            tripleDESCryptoServiceProvider.Clear()
            md5CryptoServiceProvider.Clear()
        End Try
        Dim [string] As String = utf8Encoding.GetString(bytes)
        Return Encoding.[Default].GetString(Convert.FromBase64String([string]))
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
