Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Double

'.LG M Eclipse Public License -v 1.0 
'includes code from:
'VB Forums
'Developer: chris128(http://www.vbforums.com/member.php?87550-chris128)
'Code avaliable at: http://www.vbforums.com/showthread.php?534956-RESOLVED-Detecting-Drive-Letter-of-USB-Drive
'
'MSDN Forums
'Xiaoyun Li(http://social.msdn.microsoft.com/profile/xiaoyun%20li%20%E2%80%93%20msft/?ws=usercard-mini)
'Code avaliable at: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/c1a24688-d844-4adc-9d85-416a7158c6ba/
'
'StackOveflow
'Abluescarab http://stackoverflow.com/users/567983/abluescarab
'Code avaliable at: http://stackoverflow.com/questions/5619481/what-keys-are-represented-by-the-h-hex-codes
'
'Lucas Guimarães
'http://lucasguimaraes.com
'



Public Class form1

    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2
    Public Const WM_HOTKEY As Integer = &H312
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const VK_F1 As Integer = &H70 'F1 key
    Public Const VK_F2 As Integer = &H71 'F2 key
    Public Const VK_F3 As Integer = &H72 'F3 key
    Public Const VK_SUBTRACT As Integer = &H6D '- key
    Public Const VK_ADD As Integer = &H6B '+ key
    Public Const VK_CAPITAL As Integer = &H14 'Caps-lock key
    Dim Destination As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\MBackup\" 'Default place to Save

    'Device information structure
    Public Structure DEV_BROADCAST_HDR
        Public dbch_size As Int32
        Public dbch_devicetype As Int32
        Public dbch_reserved As Int32
    End Structure

    'Volume information Structure
    Private Structure DEV_BROADCAST_VOLUME
        Public dbcv_size As Int32
        Public dbcv_devicetype As Int32
        Public dbcv_reserved As Int32
        Public dbcv_unitmask As Int32
        Public dbcv_flags As Int16
    End Structure

    'Function that gets the drive letter from the unit mask
    Private Function GetDriveLetterFromMask(ByRef Unit As Int32) As Char
        For i As Integer = 0 To 25
            If Unit = (2 ^ i) Then
                Return Chr(Asc("A") + i)
            End If
        Next
    End Function

    Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Integer
    Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As IntPtr, ByVal id As Integer) As Integer

    'Override message processing to check for the DEVICECHANGE message
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        If m.Msg = WM_DEVICECHANGE Then
            If CInt(m.WParam) = DBT_DEVICEARRIVAL Then
                Dim DeviceInfo As DEV_BROADCAST_HDR
                DeviceInfo = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_HDR)), DEV_BROADCAST_HDR)
                If DeviceInfo.dbch_devicetype = DBT_DEVTYP_VOLUME Then
                    Dim Volume As DEV_BROADCAST_VOLUME
                    Volume = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                    Dim DriveLetter As String = (GetDriveLetterFromMask(Volume.dbcv_unitmask) & ":\")

                    NotifyIcon1.ShowBalloonTip(10000, "", "New USB device detected. Starting search...", ToolTipIcon.Info)

                    Dim Source As String = DriveLetter

                    For Each File As String In My.Computer.FileSystem.GetFiles(Source, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
                        My.Computer.FileSystem.CopyFile(File, String.Concat(Destination, Microsoft.VisualBasic.Right(File, Microsoft.VisualBasic.Len(File) - 3)), True)
                    Next
                    NotifyIcon1.ShowBalloonTip(10000, "", "M has finished copying files.", ToolTipIcon.Info)

                End If
            End If
        End If

        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "7"
                    Destination = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\MBackup\"
                    NotifyIcon1.ShowBalloonTip(9000, "", "Files will be saved on:" + Destination, ToolTipIcon.Info)
                Case "8"
                    Destination = My.Computer.FileSystem.CurrentDirectory & "\MBackup\"
                    NotifyIcon1.ShowBalloonTip(9000, "", "Files will be saved on:" + Destination, ToolTipIcon.Info)
                Case "9"
                    Close()
            End Select


        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        RegisterHotKey(Me.Handle, 9, MOD_ALT, VK_F1) 'Registers Alt + F1
        RegisterHotKey(Me.Handle, 8, MOD_ALT, VK_F2) 'Registers Alt + F2
        RegisterHotKey(Me.Handle, 7, MOD_ALT, VK_F2) 'Registers Alt + F3
        Me.Visible = False
        NotifyIcon1.Visible = True
    End Sub

    Private Sub form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Unregister HotKeys before ending
        UnregisterHotKey(Me.Handle, 9)
        UnregisterHotKey(Me.Handle, 8)
        UnregisterHotKey(Me.Handle, 7)

    End Sub


    Private Sub ToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Close()
    End Sub


    Private Sub ToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        AboutBox1.ShowDialog()
    End Sub
End Class
