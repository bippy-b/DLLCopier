Imports System.IO

Module IVRS

    Sub Main()
        Dim strDir As String
        Dim intStart As Integer
        Dim intEnd As Integer
        Try
            intEnd = CType(New System.Configuration.AppSettingsReader().GetValue("EndCopyValue", GetType(Integer)), Integer)
            intStart = CType(New System.Configuration.AppSettingsReader().GetValue("StartCopyValue", GetType(Integer)), Integer)
            strDir = CType(New System.Configuration.AppSettingsReader().GetValue("CheckDir", GetType(String)), String)
            Dim myDir As New DirectoryInfo(strDir)
            Dim myFiles As FileInfo
            For Each myFiles In myDir.GetFileSystemInfos("*_1.DLL")
                Debug.WriteLine(myFiles.FullName.ToString())
                Dim i As Integer
                Dim newfile As String
                Dim j As Integer
                Dim updatedfile As String

                'Debug.WriteLine(Replace(myFiles.FullName.ToString(), "1.DLL", i.ToString() + ".DLL"))
                For i = intStart To intEnd
                    newfile = "_" & i.ToString()
                    Debug.WriteLine(newfile)
                    If File.Exists(myFiles.FullName.Replace("_1", newfile)) = False Then
                        Try
                            File.Copy(myFiles.FullName.ToString(), myFiles.FullName.Replace("_1", newfile), True)
                        Catch ex As Exception
                            Debug.WriteLine(ex.Message)
                        End Try
                    Else
                        If File.GetLastWriteTime(myFiles.FullName.Replace("_1", newfile)) <> File.GetLastWriteTime(myFiles.FullName) Then
                            For j = 2 To intEnd
                                updatedfile = myFiles.FullName.Replace("_1", "_" & j)
                                File.Copy(myFiles.FullName.ToString(), updatedfile, True)
                            Next

                        End If
                    End If

                Next ' For i = intStart to intEnd
            Next ' For each myFiles
        Catch ex As Exception
            'write out error to text file for viewing
            Dim sw As New System.IO.StreamWriter("errorlog.txt", True)
            sw.Write("***************")
            sw.Write(Environment.NewLine)
            sw.Write(Now())
            sw.Write(Environment.NewLine)
            sw.Write("***************")
            sw.Write(Environment.NewLine)
            sw.Write(ex.Message.ToString())
            sw.Write(Environment.NewLine)
            sw.Write("***************")
            sw.Write(Environment.NewLine)
            sw.Write("End Error Log")
            sw.Write(Environment.NewLine)
            sw.Write("***************")
            sw.Write(Environment.NewLine)
            sw.Flush()
            sw.Close()
        End Try

    End Sub

End Module
