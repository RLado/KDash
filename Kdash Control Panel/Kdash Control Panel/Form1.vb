Public Class Form1
    Dim shObj As Object = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"))
    Dim myLocalAppDataFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
    Dim IsKdashON As Boolean

    Sub UnZip(inputZip, outputFolder)
        'Create directory in which you will unzip your items.
        IO.Directory.CreateDirectory(outputFolder)
        'Declare the folder where the items will be extracted.
        Dim output As Object = shObj.NameSpace((outputFolder))
        'Declare the input zip file.
        Dim input As Object = shObj.NameSpace((inputZip))
        'Extract the items from the zip file.
        output.CopyHere((input.Items), 4)

    End Sub
    Sub StartKdash()
        Dim myProcess As New Process
        myProcess.StartInfo.WorkingDirectory = myLocalAppDataFolder + "\Kdash"
        myProcess.StartInfo.FileName = "kdash.exe"
        myProcess.Start()
        IsKdashON = True
        StatusLabel.Text = "Kdash is: ENABLED"
        Button1.Text = "Stop Kdash"
    End Sub
    Sub KillKdash()
        Dim pList() As System.Diagnostics.Process =
    System.Diagnostics.Process.GetProcessesByName("kdash")
        For Each proc As System.Diagnostics.Process In pList
            proc.Kill()
            IsKdashON = False
            StatusLabel.Text = "Kdash is: DISABLED"
            Button1.Text = "Start Kdash"
        Next
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IsKdashON = False
        If (Not System.IO.Directory.Exists(myLocalAppDataFolder + "\Kdash")) Then
            IO.File.WriteAllBytes(myLocalAppDataFolder + "\Kdash.zip", My.Resources.Kdash1_2)
            UnZip(myLocalAppDataFolder + "\Kdash.zip", myLocalAppDataFolder)
        End If
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles MyBase.FormClosing
        If IsKdashON = True Then
            KillKdash()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsKdashON = False Then
            Button1.Text = "Stop Kdash"
            If (Not System.IO.Directory.Exists(myLocalAppDataFolder + "\Kdash")) Then
                IO.File.WriteAllBytes(myLocalAppDataFolder + "\Kdash.zip", My.Resources.Kdash1_2)
                UnZip(myLocalAppDataFolder + "\Kdash.zip", myLocalAppDataFolder)
            End If

            Dim myProcess As New Process
            myProcess.StartInfo.WorkingDirectory = myLocalAppDataFolder + "\Kdash"
            myProcess.StartInfo.FileName = "kdash.exe"
            myProcess.Start()
            IsKdashON = True
            StatusLabel.Text = "Kdash is: ENABLED"
            'Process.Start(myLocalAppDataFolder + "\Kdash\kdash.exe")
        Else
            Button1.Text = "Start Kdash"
            Dim pList() As System.Diagnostics.Process =
            System.Diagnostics.Process.GetProcessesByName("kdash")
            For Each proc As System.Diagnostics.Process In pList
                proc.Kill()
                IsKdashON = False
                StatusLabel.Text = "Kdash is: DISABLED"
            Next
        End If
    End Sub 'Start stop button [It's Done]
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click 'Add profile button
        Dim FileHeader As String
        Dim ProfileName As String
        FileHeader = "Welcome to kdash config file:
To assign a macro to a key write $xxx=[Code], where xxx is the 3 number KeyID
i.e. $113=Hello World [Add € to press enter]
A wait time in sec. can be added by using |X| i.e. |0.5|
Set keypress interval:
%%0.005
==============================================="
        ProfileName = InputBox("Input the name of the new profile:", "New Profile", "New_Profile")
        'MsgBox(ProfileName)
        My.Computer.FileSystem.WriteAllText(myLocalAppDataFolder + "\Kdash\" + ProfileName + ".kprf", FileHeader, False)
    End Sub
    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        'Executes when combo box is clicked. Is used to list the profile files
        Dim kprofiles As String() = System.IO.Directory.GetFiles(myLocalAppDataFolder + "\Kdash\", "*.kprf")
        Dim ComboIndex As String
        ComboBox1.Items.Clear() 'Clear all current entries
        For i As Integer = 0 To kprofiles.Length - 1
            'MsgBox(kprofiles(i))
            'Obtains name of the file
            ComboIndex = Split(kprofiles(i), "\")(Split(kprofiles(i), "\").Length - 1)
            'Remove file extension
            ComboIndex = Split(ComboIndex, ".")(0)
            'Add to Combobox
            ComboBox1.Items.Add(ComboIndex)
        Next
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim selectedProfile As String = ComboBox1.SelectedItem
        If (System.IO.File.Exists(myLocalAppDataFolder + "\Kdash\" + selectedProfile + ".kprf")) Then
            'Read and write all in bytes or text encoding gets messed up. The python core program seems to work right with notepad default ANSI encoding.
            My.Computer.FileSystem.WriteAllBytes(myLocalAppDataFolder + "\Kdash\Kdash.conf", My.Computer.FileSystem.ReadAllBytes(myLocalAppDataFolder + "\Kdash\" + selectedProfile + ".kprf"), False)
            If (IsKdashON = True) Then
                KillKdash()
                StartKdash()
            End If
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim selectedProfile As String = ComboBox1.SelectedItem
        'Opens default text editor in the system to edit the profile file
        If (System.IO.File.Exists(myLocalAppDataFolder + "\Kdash\" + selectedProfile + ".kprf")) Then
            System.Diagnostics.Process.Start(myLocalAppDataFolder + "\Kdash\" + selectedProfile + ".kprf")
        End If
        If (IsKdashON = True) Then
            KillKdash()
        End If
        ComboBox1.SelectedIndex = -1 'Resets ComboBox
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click 'Deleteprofile button
        Dim selectedProfile As String = ComboBox1.SelectedItem
        'Opens default text editor in the system to edit the profile file
        If (System.IO.File.Exists(myLocalAppDataFolder + "\Kdash\" + selectedProfile + ".kprf")) Then
            My.Computer.FileSystem.DeleteFile(myLocalAppDataFolder + "\Kdash\" + selectedProfile + ".kprf")
        End If
        If (IsKdashON = True) Then
            KillKdash()
        End If
        ComboBox1.SelectedIndex = -1 'Resets ComboBox
    End Sub
End Class
