Imports System
Imports System.Diagnostics


Public Class Form1
    Public MyProcName As String
    Public MyProcPriority As Integer
    Public MyProcPriorityString As String
    Public MyProcID As Integer
    Public PRCID As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Prc() As Process
        Dim PrCl() = {"AboveNormal", "BelowNormal", "High", "Idle", "Normal", "RealTime"}
        Dim objName(1000) As String
        Dim objNameInt(1000) As Integer
        Dim objParentInt(1000) As Integer
        Dim objParentInt1(1000) As Integer
        Dim objParentString(1000) As String
        Dim objParent(1000) As String
        Dim objParent1(1000) As String
        'Dim CheckIt As Boolean
        Dim PC As String
        Dim i, j As Integer
        Dim temp3, temp4, t, tt, ttt As String
        Dim objService, objProc, str
        'Dim nm As ListViewItem
        'Dim Proc As Process
        Dim nodename1
        ComboBox1.Items.Clear()
        For i = 0 To UBound(PrCl) - 1
            ComboBox1.Items.Add(PrCl(i).ToString)
        Next
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        TreeView1.Nodes.Clear()
        TreeView2.Nodes.Clear()
        Label5.Text = "Процесс " & ":"
        ListView1.Clear()
        ' Set to details view
        ListView1.View = View.Details
        ' Add a column with width 20 and left alignment


        ListView1.Columns.Add("ID:", 50, HorizontalAlignment.Left)
        ListView1.Columns.Add("Имя процесса:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Начал работать:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Приоритет:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Текущее окно:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Завис:", 100, HorizontalAlignment.Left)
        ListView1.GridLines = True

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
        i = 0
        For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Process")
            objName(i) = objProc.Caption
            objParentInt(i) = objProc.ParentProcessId
            objNameInt(i) = objProc.ProcessId
            ListBox3.Items.Add(objProc.ParentProcessId)
            'On Error Resume Next
            'objParentString(i) = Process.GetProcessById(objProc.ParentProcessId, PC).ToString()

            '  TreeView1.Nodes.Add(Process.GetProcessById(objProc.ParentProcessId, PC).ToString)
            '    TreeView1.Nodes(0).Nodes.Add(objProc.Caption)
            '   End If

            i = i + 1
        Next
        'Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(objParentInt)
            For j = 0 To UBound(objNameInt)

                'objParentString(i) = Process.GetProcessById(objParentInt(i), PC).ToString()
                If objParentInt(i) = objNameInt(j) Then objParentString(i) = objName(j)
            Next
        Next
        'str = vbNullString
        For i = 0 To UBound(objName)
            If objName(i) <> "" Then ListBox1.Items.Add(objName(i))
            If objParentString(i) <> "" Then ListBox2.Items.Add(objParentString(i))
        Next
        'For i = 0 To UBound(objParentString)
        'objParent(i) = objParentString(i)
        'Next
        'CheckIt = False
        objParent(0) = objParentString(0)

        For i = 1 To UBound(objParent) - 1
            If objParentString(i) = objParentString(i + 1) Then
                Continue For
            Else
                objParent(i) = objParentString(i)
            End If
        Next




        '--------------Remove Repeated Items---------------

        ListBox4.Items.Clear()
        For i = 0 To UBound(objParent)
            If objParent(i) <> "" Then ListBox4.Items.Add(objParent(i))
        Next

        For i = 0 To UBound(objParent)
            objParent(i) = ""
        Next
        For i = 0 To ListBox4.Items.Count - 1
            objParent(i) = ListBox4.Items.Item(i)
        Next

        For i = 0 To UBound(objParent) - 2
            If objParent(i) = objParent(i + 1) Then
                objParent(i + 1) = ""
            End If
        Next

        ListBox4.Items.Clear()
        For i = 0 To UBound(objParent)
            If objParent(i) <> "" Then ListBox4.Items.Add(objParent(i))
        Next

        For i = 0 To UBound(objParent)
            objParent(i) = ""
        Next
        For i = 0 To ListBox4.Items.Count - 1
            objParent(i) = ListBox4.Items.Item(i)
        Next

        For i = 0 To UBound(objParent) - 2
            If objParent(i) = objParent(i + 1) Then
                objParent(i + 1) = ""
            End If
        Next

        ListBox4.Items.Clear()
        For i = 0 To UBound(objParent)
            If objParent(i) <> "" Then ListBox4.Items.Add(objParent(i))
        Next

        For i = 0 To UBound(objParent)
            objParent(i) = ""
        Next
        For i = 0 To ListBox4.Items.Count - 1
            objParent(i) = ListBox4.Items.Item(i)
        Next



        '--------------Repeated Items Has Been Removed---------------






       
        For i = 0 To ListBox4.Items.Count - 1

            TreeView1.Nodes.Add(objParent(i))
            If objParent(i) = "explorer.exe" Then
                TreeView2.Nodes.Add("DESKTOP")
            End If

            For j = 0 To 100
                If objParent(i) = objParentString(j) Then
                    TreeView1.Nodes(i).Nodes.Add(objName(j))
                    If objParent(i) = "explorer.exe" Then
                        TreeView2.Nodes(0).Nodes.Add(objName(j))
              
                    End If
                    End If

            Next
        Next



        PC = System.Environment.MachineName.ToString
        Try
            Prc = Process.GetProcesses(PC)

            For i = 0 To UBound(Prc)

                'ListBox3.Items.Add(Prc(i).Id)
                temp3 = Prc(i).ProcessName
                temp4 = temp3.IndexOf("MyMonHyb")
                If temp4 <> -1 Then
                    MyProcName = Prc(i).ProcessName
                    MyProcPriority = Prc(i).PriorityClass
                    MyProcPriorityString = Prc(i).PriorityClass.ToString
                    MyProcID = Prc(i).Id

                End If
                '  For j = 0 To TreeView2.Nodes(0).Nodes.Count - 1
                't = (TreeView2.SelectedNode).ToString
                ' temp = Replace(t, "TreeNode: ", "")
                'If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")
                'If Prc(i).ProcessName.ToString = ttt Then
                'If Prc(i).MainWindowTitle <> "" Then
                '           TreeView2.Nodes(.Nodes.Add(Prc(i).MainWindowTitle)
                'Else
                'TreeView2.SelectedNode.Nodes.Add("Окно скрыто.")
                'End If
                'End If
                ' Next
            Next



        Catch
        End Try
        ListBox5.Items.Add("Имя компьютера: " & PC.ToString)
        ListBox5.Items.Add("Название этой программы: " & MyProcName)
        ListBox5.Items.Add("Приоритет этой программы: " & MyProcPriority & ", " & MyProcPriorityString)
        ListBox5.Items.Add("ID этой программы: " & MyProcID)




    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Procek As Process
        Dim Prc() As Process
        Dim PC, temp, temp3, temp4, t, tt, ttt As String
        Dim countIndex As Integer = 0
        Dim selectedNode As String = "Selected customer nodes are : "
        Dim myNode, str As String
        Dim i As Integer
        On Error GoTo ErrorResolve


        t = (TreeView1.SelectedNode).ToString
        temp = Replace(t, "TreeNode: ", "")
        If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")

       
        'Catch
        'End Try

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(Prc)

            If Prc(i).ProcessName.ToString = ttt Then
                If Prc(i).PriorityClass < MyProcPriority Then
                    Prc(i).Kill()
                    TreeView1.SelectedNode.Remove()
                Else
                    MsgBox("Нельзя удалять! Приоритет больше или равен нашему! Приоритет процесса " & Prc(i).ProcessName & ": " & Prc(i).PriorityClass, MsgBoxStyle.OkOnly, "MyMonHyb: Process Tree Edition")
                    Exit For
                End If
            End If
        Next
ErrorResolve:
        Exit Sub
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Prc() As Process
        Dim PrCl() = {"AboveNormal", "BelowNormal", "High", "Idle", "Normal", "RealTime"}
        Dim objName(1000) As String
        Dim objNameInt(1000) As Integer
        Dim objParentInt(1000) As Integer
        Dim objParentInt1(1000) As Integer
        Dim objParentString(1000) As String
        Dim objParent(1000) As String
        Dim objParent1(1000) As String
        'Dim CheckIt As Boolean
        Dim PC As String
        Dim i, j As Integer
        Dim temp3, temp4, t, tt, ttt As String
        Dim objService, objProc, str
        'Dim nm As ListViewItem
        'Dim Proc As Process
        Dim nodename1
        ComboBox1.Items.Clear()
        For i = 0 To UBound(PrCl) - 1
            ComboBox1.Items.Add(PrCl(i).ToString)
        Next
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        TreeView1.Nodes.Clear()
        TreeView2.Nodes.Clear()
        ListView1.Clear()
        Label5.Text = "Процесс " & ":"
        ' Set to details view
        ListView1.View = View.Details
        ' Add a column with width 20 and left alignment


        ListView1.Columns.Add("ID:", 50, HorizontalAlignment.Left)
        ListView1.Columns.Add("Имя процесса:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Начал работать:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Приоритет:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Текущее окно:", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Завис:", 100, HorizontalAlignment.Left)
        ListView1.GridLines = True

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
        i = 0
        For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Process")
            objName(i) = objProc.Caption
            objParentInt(i) = objProc.ParentProcessId
            objNameInt(i) = objProc.ProcessId
            ListBox3.Items.Add(objProc.ParentProcessId)
            'On Error Resume Next
            'objParentString(i) = Process.GetProcessById(objProc.ParentProcessId, PC).ToString()

            '  TreeView1.Nodes.Add(Process.GetProcessById(objProc.ParentProcessId, PC).ToString)
            '    TreeView1.Nodes(0).Nodes.Add(objProc.Caption)
            '   End If

            i = i + 1
        Next
        'Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(objParentInt)
            For j = 0 To UBound(objNameInt)

                'objParentString(i) = Process.GetProcessById(objParentInt(i), PC).ToString()
                If objParentInt(i) = objNameInt(j) Then objParentString(i) = objName(j)
            Next
        Next
        'str = vbNullString
        For i = 0 To UBound(objName)
            If objName(i) <> "" Then ListBox1.Items.Add(objName(i))
            If objParentString(i) <> "" Then ListBox2.Items.Add(objParentString(i))
        Next
        'For i = 0 To UBound(objParentString)
        'objParent(i) = objParentString(i)
        'Next
        'CheckIt = False
        objParent(0) = objParentString(0)

        For i = 1 To UBound(objParent) - 1
            If objParentString(i) = objParentString(i + 1) Then
                Continue For
            Else
                objParent(i) = objParentString(i)
            End If
        Next




        '--------------Remove Repeated Items---------------

        ListBox4.Items.Clear()
        For i = 0 To UBound(objParent)
            If objParent(i) <> "" Then ListBox4.Items.Add(objParent(i))
        Next

        For i = 0 To UBound(objParent)
            objParent(i) = ""
        Next
        For i = 0 To ListBox4.Items.Count - 1
            objParent(i) = ListBox4.Items.Item(i)
        Next

        For i = 0 To UBound(objParent) - 2
            If objParent(i) = objParent(i + 1) Then
                objParent(i + 1) = ""
            End If
        Next

        ListBox4.Items.Clear()
        For i = 0 To UBound(objParent)
            If objParent(i) <> "" Then ListBox4.Items.Add(objParent(i))
        Next

        For i = 0 To UBound(objParent)
            objParent(i) = ""
        Next
        For i = 0 To ListBox4.Items.Count - 1
            objParent(i) = ListBox4.Items.Item(i)
        Next

        For i = 0 To UBound(objParent) - 2
            If objParent(i) = objParent(i + 1) Then
                objParent(i + 1) = ""
            End If
        Next

        ListBox4.Items.Clear()
        For i = 0 To UBound(objParent)
            If objParent(i) <> "" Then ListBox4.Items.Add(objParent(i))
        Next

        For i = 0 To UBound(objParent)
            objParent(i) = ""
        Next
        For i = 0 To ListBox4.Items.Count - 1
            objParent(i) = ListBox4.Items.Item(i)
        Next



        '--------------Repeated Items Has Been Removed---------------







        For i = 0 To ListBox4.Items.Count - 1

            TreeView1.Nodes.Add(objParent(i))
            If objParent(i) = "explorer.exe" Then
                TreeView2.Nodes.Add("DESKTOP")
            End If

            For j = 0 To 100
                If objParent(i) = objParentString(j) Then
                    TreeView1.Nodes(i).Nodes.Add(objName(j))
                    If objParent(i) = "explorer.exe" Then
                        TreeView2.Nodes(0).Nodes.Add(objName(j))

                    End If
                End If

            Next
        Next



        PC = System.Environment.MachineName.ToString
        Try
            Prc = Process.GetProcesses(PC)

            For i = 0 To UBound(Prc)

                'ListBox3.Items.Add(Prc(i).Id)
                temp3 = Prc(i).ProcessName
                temp4 = temp3.IndexOf("MyMonHyb")
                If temp4 <> -1 Then
                    MyProcName = Prc(i).ProcessName
                    MyProcPriority = Prc(i).PriorityClass
                    MyProcPriorityString = Prc(i).PriorityClass.ToString
                    MyProcID = Prc(i).Id

                End If
                '  For j = 0 To TreeView2.Nodes(0).Nodes.Count - 1
                't = (TreeView2.SelectedNode).ToString
                ' temp = Replace(t, "TreeNode: ", "")
                'If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")
                'If Prc(i).ProcessName.ToString = ttt Then
                'If Prc(i).MainWindowTitle <> "" Then
                '           TreeView2.Nodes(.Nodes.Add(Prc(i).MainWindowTitle)
                'Else
                'TreeView2.SelectedNode.Nodes.Add("Окно скрыто.")
                'End If
                'End If
                ' Next
            Next



        Catch
        End Try
        ListBox5.Items.Add("Имя компьютера: " & PC.ToString)
        ListBox5.Items.Add("Название этой программы: " & MyProcName)
        ListBox5.Items.Add("Приоритет этой программы: " & MyProcPriority & ", " & MyProcPriorityString)
        ListBox5.Items.Add("ID этой программы: " & MyProcID)





    End Sub







    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim Procek As Process
        Dim Prc() As Process
        Dim PC, temp, temp3, temp4, t, tt, ttt, tttUP As String
        Dim countIndex As Integer = 0
        Dim selectedNode As String = "Selected customer nodes are : "
        Dim myNode, str As String
        Dim nm As ListViewItem
        Dim ThreadID As String

        Dim i, j As Integer

        On Error GoTo ErrorResolve
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        t = (TreeView1.SelectedNode).ToString
        temp = Replace(t, "TreeNode: ", "")
        If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")



        'Catch
        'End Try

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(Prc) - 1

            If Prc(i).ProcessName.ToString = ttt Then


                'PRCID = Prc(i).Id
                nm = ListView1.Items.Add(Prc(i).Id.ToString)
                nm.SubItems.Add(Prc(i).ProcessName.ToString)
                Label5.Text = "Процесс " & Prc(i).ProcessName.ToString & ":"
                nm.SubItems.Add(Prc(i).StartTime.ToLocalTime)
                nm.SubItems.Add(Prc(i).PriorityClass & ", " & Prc(i).PriorityClass.ToString)
                TextBox1.Text = Prc(i).Id.ToString

                nm.SubItems.Add(Prc(i).MainWindowTitle)


                If Prc(i).Responding.ToString = "True" Then nm.SubItems.Add("Нет")
                If Prc(i).Responding.ToString = "False" Then nm.SubItems.Add("Да")

                For j = 0 To Prc(i).Threads.Count - 1
                    ListBox6.Items.Add(Prc(i).Threads(j).Id.ToString)
                    j = j + 1
                Next
                For j = 0 To Prc(i).Modules.Count - 1
                    t = Prc(i).Modules(j).ToString
                    temp = Replace(t, "System.Diagnostics.ProcessModule", "")
                    t = Replace(temp, "(", "")
                    temp = Replace(t, ")", "")
                    ListBox7.Items.Add(temp)
                    nm.SubItems.Add(temp)
                Next



                ComboBox1.Text = Prc(i).PriorityClass.ToString

            End If
        Next




        Exit Sub
ErrorResolve:
        ListBox6.Items.Add("Процессы не определены.")
        Exit Sub
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub ListBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox6.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Procek As Process
        Dim Prc() As Process
        Dim PC, temp, temp3, temp4, t, tt, ttt, tttUP As String
        Dim countIndex As Integer = 0
        Dim selectedNode As String = "Selected customer nodes are : "
        Dim myNode, str As String
        Dim nm As ListViewItem
        Dim ThreadID As String

        Dim i, j As Integer
        On Error GoTo ErrorResolve


     

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(Prc)

            If Prc(i).Id.ToString = TextBox1.Text Then
                If Prc(i).PriorityClass < MyProcPriority Then
                    Prc(i).Kill()
                    TreeView1.SelectedNode.Remove()
                Else
                    MsgBox("Нельзя удалять! Приоритет больше или равен нашему! Приоритет процесса " & Prc(i).ProcessName & ": " & Prc(i).PriorityClass, MsgBoxStyle.OkOnly, "MyMonHyb: Process Tree Edition")
                    Exit For
                End If
            End If
        Next
ErrorResolve:
        Exit Sub
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Procek As Process
        Dim Prc() As Process
        Dim PC, temp, temp3, temp4, t, tt, ttt As String
        Dim countIndex As Integer = 0
        Dim selectedNode As String = "Selected customer nodes are : "
        Dim myNode, str As String
        Dim i As Integer
        'On Error GoTo ErrorResolve


        t = (TreeView1.SelectedNode).ToString
        temp = Replace(t, "TreeNode: ", "")
        If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")


        'Catch
        'End Try

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(Prc)

            If Prc(i).ProcessName.ToString = ttt Then
                Select Case (ComboBox1.Text)
                    '"AboveNormal", "BelowNormal", "High", "Idle", "Normal", "RealTime"
                    Case "AboveNormal"
                        Prc(i).PriorityClass = ProcessPriorityClass.AboveNormal
                    Case "BelowNormal"
                        Prc(i).PriorityClass = ProcessPriorityClass.BelowNormal
                    Case "High"
                        Prc(i).PriorityClass = ProcessPriorityClass.High
                    Case "Idle"
                        Prc(i).PriorityClass = ProcessPriorityClass.Idle
                    Case "Normal"
                        Prc(i).PriorityClass = ProcessPriorityClass.Normal
                    Case "RealTime"
                        Prc(i).PriorityClass = ProcessPriorityClass.RealTime
                    Case Else
                        Prc(i).PriorityClass = ProcessPriorityClass.Normal
                End Select
                '"ProcessPriorityClass."
            End If
        Next
ErrorResolve:
        Exit Sub
    End Sub

    Private Sub ListBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox5.SelectedIndexChanged

    End Sub



 
    Private Sub TreeView2_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView2.AfterSelect
        Dim Procek As Process
        Dim Prc() As Process
        Dim PC, temp, temp3, temp4, t, tt, ttt, tttUP As String
        Dim countIndex As Integer = 0
        Dim selectedNode As String = "Selected customer nodes are : "
        Dim myNode, str As String
        Dim nm As ListViewItem
        Dim ThreadID As String

        Dim node1

        Dim i, j As Integer

        On Error GoTo ErrorResolve

        t = (TreeView2.SelectedNode).ToString
        temp = Replace(t, "TreeNode: ", "")
        If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")


        'Catch
        'End Try

        PC = System.Environment.MachineName.ToString
        Prc = Process.GetProcesses(PC)
        For i = 0 To UBound(Prc) - 1
            t = (TreeView2.SelectedNode).ToString
            temp = Replace(t, "TreeNode: ", "")
            If temp.IndexOf(".EXE") <> -1 Then ttt = Replace(temp, ".EXE", "") Else ttt = Replace(temp, ".exe", "")
            If Prc(i).ProcessName.ToString = ttt Then
                If Prc(i).MainWindowTitle <> "" Then
                    TreeView2.SelectedNode.Nodes.Clear()
                    TreeView2.SelectedNode.Nodes.Add(Prc(i).MainWindowTitle)

                Else
                    TreeView2.SelectedNode.Nodes.Clear()
                    TreeView2.SelectedNode.Nodes.Add("Невидимое окно.")

                End If
            End If
        Next









ErrorResolve:
        Exit Sub

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub


End Class
