Imports System.IO
Imports System.Threading

Public Class frmMain
    Dim Path, StatMsg, Browsing_Dir, Working_File As String

    Private Sub PopulateTreeView(ByVal dir As String, ByVal parentNode As TreeNode)
        Dim folder As String = String.Empty
        Dim filename As String = String.Empty
        Try
            Dim folders() As String = IO.Directory.GetDirectories(dir)
            Dim childNode As TreeNode = Nothing
            Dim fileNode As TreeNode = Nothing

            If folders.Length <> 0 Then
                For Each folder In folders
                    childNode = New TreeNode(Mid(folder, folder.LastIndexOf("\") + 2, folder.Length))
                    childNode.ImageIndex = 1
                    parentNode.Nodes.Add(childNode)
                    PopulateTreeView(folder, childNode)
                Next
            End If

            Dim FileLocation As DirectoryInfo = New DirectoryInfo(dir)
            Dim fi As FileInfo() = FileLocation.GetFiles("*.csv")

            For Each f As FileInfo In fi
                filename = f.ToString
                fileNode = parentNode.Nodes.Add(filename)
                fileNode.ImageIndex = 0
                fileNode.SelectedImageIndex = 0
            Next
        Catch ex As UnauthorizedAccessException
            parentNode.Nodes.Add(folder & ": Access Denied")
        End Try
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        If FolderBrowserDialog1.ShowDialog <> DialogResult.Cancel Then
            txtDirectory.Text = FolderBrowserDialog1.SelectedPath
            BrowseDirectories(txtDirectory.Text)
        End If
    End Sub

    Private Sub BrowseDirectories(ByVal RootDir As String)
        Dim CurrDir As String = String.Empty
        Dim i As Integer = 0

        If Directory.Exists(RootDir) Then
            tvFolder.Nodes.Clear()

            For Each Dir As String In Directory.GetDirectories(RootDir)
                CurrDir = Dir.ToString
                'Add this drive as a root node
                tvFolder.Nodes.Add(Mid(CurrDir, CurrDir.LastIndexOf("\") + 2, CurrDir.Length))
                tvFolder.Nodes(i).ImageIndex = 1
                'Populate this root node
                PopulateTreeView(CurrDir, tvFolder.Nodes(i))

                i += 1
            Next

            Dim FileLocation As DirectoryInfo = New DirectoryInfo(RootDir)
            Dim fi As FileInfo() = FileLocation.GetFiles("*.csv")
            Dim filename As String = String.Empty
            Dim fileNode As TreeNode = Nothing

            For Each f As FileInfo In fi
                filename = f.ToString
                fileNode = tvFolder.Nodes.Add(filename)
                fileNode.ImageIndex = 0
                fileNode.SelectedImageIndex = 0
            Next
        End If
    End Sub

    Private Sub tvFolder_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles tvFolder.BeforeExpand
        e.Node.ImageIndex = 2
        e.Node.SelectedImageIndex = 2
    End Sub

    Private Sub tvFolder_BeforeCollapse(sender As Object, e As TreeViewCancelEventArgs) Handles tvFolder.BeforeCollapse
        e.Node.ImageIndex = 1
        e.Node.SelectedImageIndex = 1
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If Directory.Exists(txtDirectory.Text) Then
            Path = txtDirectory.Text
            ToolStrip1.Enabled = False
            btnBrowse.Enabled = False
            ProgressBar1.Style = ProgressBarStyle.Marquee
            Dim t As New Thread(AddressOf Me.UpdateData)
            t.Start()
        Else
            MsgBox("No Folder Selected.", vbInformation)
        End If
    End Sub

    Private Sub UpdateStatus()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf UpdateStatus))
        Else
            lblStatus.Text = StatMsg
        End If
    End Sub

    Private Sub CurrentlyBrowsing()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf CurrentlyBrowsing))
        Else
            lblBrowse.Text = Browsing_Dir
        End If
    End Sub

    Private Sub CurrentlyWorking()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf CurrentlyWorking))
        Else
            lblWork.Text = Working_File
        End If
    End Sub

    Private Sub EndProcessing()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf EndProcessing))
        Else
            lblStatus.Text = "Ready"
            lblBrowse.Text = "-"
            lblWork.Text = "-"

            Me.ProgressBar1.Style = ProgressBarStyle.Blocks
            Me.ToolStrip1.Enabled = True
            Me.btnBrowse.Enabled = False

            Me.Focus()
            Shell("explorer """ & txtDirectory.Text & """", AppWinStyle.NormalFocus)
        End If
    End Sub

    Private Sub UpdateData()
        BrowseFiles(Path)
    End Sub

    Private Sub BrowseFiles(ByVal Path As String)
        Dim FileLocation As DirectoryInfo
        Dim fi As FileInfo()

        Dim oExcel As Object = Nothing
        Dim oBook As Object = Nothing
        Dim oSheet As Object = Nothing

        If Directory.GetDirectories(Path).Count > 0 Then
            For Each Dir As String In Directory.GetDirectories(Path)
                StatMsg = "Currently Accessing " & Dir
                UpdateStatus()

                Browsing_Dir = Dir
                CurrentlyBrowsing()

                FileLocation = New DirectoryInfo(Dir)
                fi = FileLocation.GetFiles("*.csv")

                Dim fc As Integer = fi.Count
                Dim i As Integer = 9

                If fc > 0 Then
                    oExcel = CreateObject("Excel.Application")
                    oBook = oExcel.Workbooks.Open(Application.StartupPath & "\Templates\Summary.csv", False)
                    oSheet = oBook.Worksheets(1)
                End If

                For Each f As FileInfo In fi
                    Working_File = f.ToString
                    CurrentlyWorking()

                    UpdateFile(Dir, f.ToString, oSheet, i)
                Next

                If fc > 0 Then
                    oBook.SaveAs(Dir & "\Consolidated.csv", 6)
                    oBook.Close()
                    oExcel.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)

                    oSheet = Nothing
                    oBook = Nothing
                    oExcel = Nothing
                End If

                If Directory.GetDirectories(Dir).Count > 0 Then BrowseFiles(Dir)
            Next
        Else
            StatMsg = "Currently Accessing " & Path
            UpdateStatus()

            Browsing_Dir = Path
            CurrentlyBrowsing()

            FileLocation = New DirectoryInfo(Path)
            fi = FileLocation.GetFiles("*.csv")
            Dim upd As FileInfo() = FileLocation.GetFiles("*_updated.csv")

            If upd.Count = 0 Then
                Dim fc As Integer = fi.Count
                Dim i As Integer = 9

                If fc > 0 Then
                    oExcel = CreateObject("Excel.Application")
                    oBook = oExcel.Workbooks.Open(Application.StartupPath & "\Templates\Summary.csv", False)
                    oSheet = oBook.Worksheets(1)
                End If

                For Each f As FileInfo In fi
                    Working_File = f.ToString
                    CurrentlyWorking()

                    UpdateFile(Path, f.ToString, oSheet, i)
                Next

                If fc > 0 Then
                    oBook.SaveAs(Path & "\Consolidated.csv", 6)
                    oBook.Close()
                    oExcel.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)

                    oSheet = Nothing
                    oBook = Nothing
                    oExcel = Nothing
                End If
            End If
        End If

        EndProcessing()
    End Sub

    Private Function GetHours(ByVal date_str As String) As Integer
        Dim dt() As String = date_str.Split(" ")
        Return Hour(dt(1))
    End Function

    Private Sub UpdateFile(ByVal Folder_Path As String, ByVal File_Name As String, ByRef SummSheet As Object, ByRef SummIx As Integer)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Open(Folder_Path & "\" & File_Name, False)
        oSheet = oBook.Worksheets(1)

        Dim i As Integer = 9

        Dim f_val, g_val, i_val, sum_f, sum_g, sum_j, count_j As Single
        Dim c_hour As Integer = -1
        Dim hr_str As String = String.Empty

        sum_f = 0
        sum_g = 0
        sum_j = 0
        count_j = 0

        With oSheet
            While .Range("A" & i).Text <> ""
                If c_hour <> GetHours(.Range("A" & i).Value) Then
                    If c_hour >= 0 Then
                        SummSheet.Range("A" & SummIx).Value = hr_str
                        SummSheet.Range("B" & SummIx).Value = .Range("B" & i - 1).Value
                        SummSheet.Range("C" & SummIx).Value = .Range("C" & i - 1).Value
                        SummSheet.Range("D" & SummIx).Value = .Range("D" & i - 1).Value
                        SummSheet.Range("E" & SummIx).Value = .Range("E" & i - 1).Value
                        SummSheet.Range("F" & SummIx).Value = Math.Round(sum_f, 2)
                        SummSheet.Range("G" & SummIx).Value = Math.Round(sum_g, 2)
                        SummSheet.Range("H" & SummIx).Value = .Range("H" & i - 1).Value
                        SummSheet.Range("I" & SummIx).Value = .Range("I" & i - 1).Value
                        SummSheet.Range("J" & SummIx).Value = Math.Round(sum_j, 2)
                        SummSheet.Range("K" & SummIx).Value = Math.Round(sum_j / count_j, 2)

                        SummIx += 1
                    End If

                    c_hour = GetHours(.Range("A" & i).Value)
                    hr_str = .Range("A" & i).Text
                    sum_f = 0
                    sum_g = 0
                    sum_j = 0
                    count_j = 0
                End If

                f_val = Replace(.Range("F" & i).Text, ",", ".")
                g_val = Replace(.Range("G" & i).Text, ",", ".")
                i_val = Replace(.Range("I" & i).Text, ",", ".")

                .Range("F" & i).Value = Math.Round(If(f_val >= 3.1 And f_val <= 3.99, 10, 0) + f_val, 2)
                .Range("G" & i).Value = Math.Round(If(g_val >= 3.1 And g_val <= 3.99, 10, 0) + g_val, 2)
                .Range("I" & i).Value = Math.Round(i_val, 2)

                If f_val >= 3.1 And f_val <= 3.99 Then
                    .Range("J" & i).Value = Math.Round((.Range("F" & i).Value - 4) * (1600 / 16), 0)
                End If

                sum_f += .Range("F" & i).Value
                sum_g += .Range("G" & i).Value
                sum_j += .Range("J" & i).Value
                count_j += 1

                i += 1
            End While

            SummSheet.Range("A" & SummIx).Value = hr_str
            SummSheet.Range("B" & SummIx).Value = .Range("B" & i - 1).Value
            SummSheet.Range("C" & SummIx).Value = .Range("C" & i - 1).Value
            SummSheet.Range("D" & SummIx).Value = .Range("D" & i - 1).Value
            SummSheet.Range("E" & SummIx).Value = .Range("E" & i - 1).Value
            SummSheet.Range("F" & SummIx).Value = Math.Round(sum_f, 2)
            SummSheet.Range("G" & SummIx).Value = Math.Round(sum_g, 2)
            SummSheet.Range("H" & SummIx).Value = .Range("H" & i - 1).Value
            SummSheet.Range("I" & SummIx).Value = .Range("I" & i - 1).Value
            SummSheet.Range("J" & SummIx).Value = Math.Round(sum_j, 2)
            SummSheet.Range("K" & SummIx).Value = Math.Round(sum_j / count_j, 2)

            SummIx += 1
        End With

        oBook.SaveAs(Folder_Path & "\" & Replace(File_Name, ".csv", "_updated.csv"), 6)
        oBook.Close()
        oExcel.Quit()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)

        oSheet = Nothing
        oBook = Nothing
        oExcel = Nothing
    End Sub
End Class
