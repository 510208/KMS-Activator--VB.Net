Imports System.IO

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles SetWinKMSBttn.Click
        'Responsible for detecting if IP Address text box is empty. If it is, show error.
        If String.IsNullOrWhiteSpace(IPAddressTxtBox.Text) Then
            MessageBox.Show("Please enter an IP address to set as your Windows KMS server", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Shell("cmd.exe /c" & "slmgr.vbs /skms " & IPAddressTxtBox.Text)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ActWinBttn.Click
        'Responsible for activating Windows
        Shell("cmd.exe /c" & "slmgr.vbs /ato")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles ViewWinActStatBttn.Click
        'Responsible for displaying detailed activation status
        Shell("cmd.exe /c" & "slmgr.vbs /dlv")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles SetOffice2013KMSBttn.Click
        'Responsible for setting the Office 2013 KMS Server to the value in the IP Address text field
        If String.IsNullOrWhiteSpace(IPAddressTxtBox.Text) Then
            MessageBox.Show("Please enter an IP address to set as your Office 2013 KMS server", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Shell("cmd.exe /c" & "cd C:\Program Files (x86)\Microsoft Office\Office15 && cscript ospp.vbs /sethst:" & IPAddressTxtBox.Text & " && pause", vbNormalFocus)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles ViewOffice2013ActStatBttn.Click
        'Responsible for displaying Office 2013 detailed activation status
        Shell("cmd.exe /c" & "cd C:\Program Files (x86)\Microsoft Office\Office15 && cscript ospp.vbs /dstatusall && pause", vbNormalFocus)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles ActOffice2013Bttn.Click
        'Responsible for activating Office 2013
        Shell("cmd.exe /c" & "cd C:\Program Files (x86)\Microsoft Office\Office15 && cscript ospp.vbs /act && pause", vbNormalFocus)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles StartKMSServerBttn.Click
        'Responsible for starting VM
        If VMPathTxtBox.Text = "" Then
            MessageBox.Show("Please enter a path to your KMS virtual machine", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Process.Start(VMPathTxtBox.Text)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles VMPathBrowseBttn.Click
        'Responsible for opening browse dialog box
        OpenFileDialog1.ShowDialog()
        Dim file_path As String = OpenFileDialog1.FileName
        VMPathTxtBox.Text = file_path
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Form load event to restore text back to IP Address and VM text boxes
        VMPathTxtBox.Text = My.Settings.File_path_thing
        If VMPathTxtBox.Text = "" Then
            VMPathTxtBox.Clear()
        Else
            VMPathTxtBox.Text = My.Settings.File_path_thing
        End If

        IPAddressTxtBox.Text = My.Settings.IP_Address
        If IPAddressTxtBox.Text = "" Then
            IPAddressTxtBox.Clear()
        Else
            IPAddressTxtBox.Text = My.Settings.IP_Address
        End If
    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As EventArgs) Handles MyBase.FormClosing
        'Form close event to save VM path and IP Address Info
        My.Settings.File_path_thing = VMPathTxtBox.Text
        My.Settings.IP_Address = IPAddressTxtBox.Text
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles SetOffice2016KMSBttn.Click
        'Responsible for setting the Office 2016 KMS Server to the value in the IP Address text field
        If String.IsNullOrWhiteSpace(IPAddressTxtBox.Text) Then
            MessageBox.Show("Please enter an IP address to set as your Office 2016 KMS server", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Shell("cmd.exe /c" & "cd C:\Program Files (x86)\Microsoft Office\Office16 && cscript ospp.vbs /sethst:" & IPAddressTxtBox.Text & " && pause", vbNormalFocus)
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles ActOffice2016Bttn.Click
        'Responsible for activating Office 2013
        Shell("cmd.exe /c" & "cd C:\Program Files (x86)\Microsoft Office\Office16 && cscript ospp.vbs /act && pause", vbNormalFocus)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles ViewOffice2016ActStatBttn.Click
        'Responsible for displaying Office 2013 detailed activation status
        Shell("cmd.exe /c" & "cd C:\Program Files (x86)\Microsoft Office\Office16 && cscript ospp.vbs /dstatusall && pause", vbNormalFocus)
    End Sub
End Class
