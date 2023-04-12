Imports System.Threading
Imports System.Runtime.InteropServices
Imports MIDI_Soundboard.libZPlay

Public Class Form1
    Public Declare Function midiInGetNumDevs Lib "winmm.dll" () As Integer
    Public Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As MIDIINCAPS, ByVal uSize As Integer) As Integer
    Public Declare Function midiInOpen Lib "winmm.dll" (ByRef hMidiIn As Integer, ByVal uDeviceID As Integer, ByVal dwCallback As MidiInCallback, ByVal dwInstance As Integer, ByVal dwFlags As Integer) As Integer
    Public Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer

    Public Delegate Function MidiInCallback(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer
    Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)
    Public Const CALLBACK_FUNCTION As Integer = &H30000
    Public Const MIDI_IO_STATUS = &H20
    Dim Knopf As String
    Public Delegate Sub DisplayDataDelegate(dwParam1)
    Dim player1 As New ZPlay()
    Dim player2 As New ZPlay()
    Dim player3 As New ZPlay()
    Dim player4 As New ZPlay()
    Dim player5 As New ZPlay()
    Dim mode As Integer = 1
    Dim modelist As New List(Of String)
    Dim modepath As String = ""
    Dim modelines As New List(Of String)
    Dim Soundlist As String 'Path of the Soundlist of Mode
    Dim Soundlines As New List(Of String) 'Read lines from Soundlistfile
    Dim Soundpaths As New List(Of String) 'Paths of Sounds of Soundlines.Item(i).split(";").

    Public Structure MIDIINCAPS
        Dim wMid As Int16 ' Manufacturer ID
        Dim wPid As Int16 ' Product ID
        Dim vDriverVersion As Integer ' Driver version
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Dim szPname As String ' Product Name
        Dim dwSupport As Integer ' Reserved
    End Structure

    Dim hMidiIn As Integer
    Dim StatusByte As Byte
    Dim DataByte1 As Byte
    Dim DataByte2 As Byte
    Dim MonitorActive As Boolean = False
    Dim HideMidiSysMessages As Boolean = False

    Function MidiInProc(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer
        If MonitorActive = True Then
            TextBox1.Invoke(New DisplayDataDelegate(AddressOf DisplayData), New Object() {dwParam1})
        End If
    End Function

    Private Sub DisplayData(dwParam1)
        If ((HideMidiSysMessages = True) And ((dwParam1 And &HF0) = &HF0)) Then
            Exit Sub
        Else
            StatusByte = (dwParam1 And &HFF)
            DataByte1 = (dwParam1 And &HFF00) >> 8
            DataByte2 = (dwParam1 And &HFF0000) >> 16
            Dim Code As String = String.Format("{0:X2} {1:X2} {2:X2}", StatusByte, DataByte1, DataByte2)
            Knopf = Code
            TextBox1.Text = Code
            If Knopf = "90 00 7F" Then
                Button4.Enabled = False
                mode = mode + 1
                If (modepath IsNot "") Then
                    Soundpaths.Clear()
                    Soundlines.Clear()
                    modelines = New List(Of String)(IO.File.ReadAllLines(modepath))
                    If modelines.Count > 1 Then
                        Dim Soundlists() As String
                        If modelines(mode) = "" Or (modelines.Count - 1) = mode Then
                            mode = 1
                        End If
                        Soundlists = modelines.Item(mode).Split(";")
                        If modelines.Item(mode) IsNot "" Then
                            Soundlines = New List(Of String)(IO.File.ReadAllLines(Soundlists(1)))
                            If (Soundlines.ToArray.Length > 0) Then
                                For h As Integer = 0 To 3
                                    Soundpaths.Add("")
                                Next
                                For i As Integer = 4 To 63
                                    If ((Soundlines.ToArray.Length) > i) Then
                                        Dim Soundinfo() As String = Soundlines.Item(i).Split(";")
                                        DirectCast(Me.Controls("Button" & i), Button).Text = Soundinfo(0)
                                        If Soundinfo.Length > 1 Then
                                            Soundpaths.Add(Soundinfo(1))
                                        Else
                                            Soundpaths.Add("")
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Else
                    MsgBox("Please load modes first.")
                End If
            End If
            If modepath IsNot "" Then
                If Knopf = "91 00 7F" Then
                    Button13.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(13), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "92 00 7F" Then
                    Button18.Enabled = False

                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(18), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If

                End If

                If Knopf = "93 00 7F" Then
                    Button23.Enabled = False

                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(23), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If

                End If

                If Knopf = "94 00 7F" Then
                    Button28.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(28), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "95 00 7F" Then
                    Button33.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(33), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "96 00 7F" Then
                    Button38.Enabled = False

                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(38), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "97 00 7F" Then
                    Button43.Enabled = False

                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(43), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "98 00 7F" Then
                    Button48.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(48), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "99 00 7F" Then
                    Button53.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(53), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "9A 00 7F" Then
                    Button58.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(58), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "9B 00 7F" Then
                    Button63.Enabled = False
                    If player1.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player1.OpenFile(Soundpaths(63), TStreamFormat.sfAutodetect) Then

                            player1.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "9C 00 7F" Then
                    Button64.Enabled = False
                    player1.StopPlayback()
                End If


                'r2 -----------------------------------------------------------------------------------------------------------------------------------------------
                If Knopf = "90 01 7F" Then
                    Button5.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(5), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "91 01 7F" Then
                    Button12.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(12), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "92 01 7F" Then
                    Button17.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(17), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "93 01 7F" Then
                    Button22.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(22), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "94 01 7F" Then
                    Button27.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(27), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "95 01 7F" Then
                    Button32.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(32), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "96 01 7F" Then
                    Button37.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(37), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "97 01 7F" Then
                    Button42.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(42), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "98 01 7F" Then
                    Button47.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(47), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "99 01 7F" Then
                    Button52.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(52), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "9A 01 7F" Then
                    Button57.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(57), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "9B 01 7F" Then
                    Button62.Enabled = False
                    If player2.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player2.OpenFile(Soundpaths(62), TStreamFormat.sfAutodetect) Then

                            player2.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "9C 01 7F" Then
                    Button65.Enabled = False
                    player2.StopPlayback()
                End If


                'r3 -----------------------------------------------------------------------------------------------------------------------------------------------
                If Knopf = "90 02 7F" Then
                    Button6.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(6), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "91 02 7F" Then
                    Button11.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(11), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "92 02 7F" Then
                    Button16.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(16), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "93 02 7F" Then
                    Button21.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(21), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "94 02 7F" Then
                    Button26.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(26), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "95 02 7F" Then
                    Button31.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(31), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "96 02 7F" Then
                    Button36.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(36), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "97 02 7F" Then
                    Button41.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(41), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "98 02 7F" Then
                    Button46.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(46), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "99 02 7F" Then
                    Button51.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(51), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9A 02 7F" Then
                    Button56.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(56), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9B 02 7F" Then
                    Button61.Enabled = False
                    If player3.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player3.OpenFile(Soundpaths(61), TStreamFormat.sfAutodetect) Then

                            player3.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9C 02 7F" Then
                    Button66.Enabled = False
                    player3.StopPlayback()
                End If

                'r4 -----------------------------------------------------------------------------------------------------------------------------------------------
                If Knopf = "90 03 7F" Then
                    Button7.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(7), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "91 03 7F" Then
                    Button10.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(10), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "92 03 7F" Then
                    Button15.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(15), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "93 03 7F" Then
                    Button20.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(20), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "94 03 7F" Then
                    Button25.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(25), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "95 03 7F" Then
                    Button30.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(30), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "96 03 7F" Then
                    Button35.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(35), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "97 03 7F" Then
                    Button40.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(40), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "98 03 7F" Then
                    Button45.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(45), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "99 03 7F" Then
                    Button50.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(50), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9A 03 7F" Then
                    Button55.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(55), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9B 03 7F" Then
                    Button60.Enabled = False
                    If player4.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player4.OpenFile(Soundpaths(60), TStreamFormat.sfAutodetect) Then

                            player4.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9C 03 7F" Then
                    Button67.Enabled = False
                    player4.StopPlayback()
                End If

                'r5 -----------------------------------------------------------------------------------------------------------------------------------------------
                If Knopf = "90 67 7F" Then
                    Button8.Enabled = False

                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(8), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If

                If Knopf = "91 67 7F" Then
                    Button9.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(9), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "92 67 7F" Then
                    Button14.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(14), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "93 67 7F" Then
                    Button19.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(19), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "94 67 7F" Then
                    Button24.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(24), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "95 67 7F" Then
                    Button29.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(29), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "96 67 7F" Then
                    Button34.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(34), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "97 67 7F" Then
                    Button39.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(39), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "98 67 7F" Then
                    Button44.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(44), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "99 67 7F" Then
                    Button49.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(49), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9A 67 7F" Then
                    Button54.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(54), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9B 67 7F" Then
                    Button59.Enabled = False
                    If player5.SetWaveOutDevice(CType(ComboBox2.SelectedIndex, UInteger)) Then

                        If player5.OpenFile(Soundpaths(59), TStreamFormat.sfAutodetect) Then

                            player5.StartPlayback()

                        End If
                    End If
                End If
                If Knopf = "9C 67 7F" Then
                    Button68.Enabled = False
                    player5.StopPlayback()
                End If

                'Black Window
                'Off
                If Knopf = "B0 07 00" Then
                    Black.Hide()
                End If
                'On
                If Knopf = "B0 07 7F" Then
                    Black.Show()
                End If
                RadioButton1.Enabled = True
                Timer1.Enabled = True
                Timer1.Start()
                RadioButton1.Checked = True

            Else
                MsgBox("Please load Modes.")
            End If
        End If
    End Sub


    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Show()
        If midiInGetNumDevs() = 0 Then
            MsgBox("No MIDI devices connected")
            Application.Exit()
        End If

        Dim InCaps As New MIDIINCAPS
        Dim DevCnt As Integer

        For DevCnt = 0 To (midiInGetNumDevs - 1)
            midiInGetDevCaps(DevCnt, InCaps, Len(InCaps))
            ComboBox1.Items.Add(InCaps.szPname)
        Next DevCnt
        Dim WaveOut As New ZPlay()
        Dim WaveOutDev As Integer
        Dim WaveOutInf As New TWaveOutInfo

        WaveOutDev = WaveOut.EnumerateWaveOut

        For x = 0 To WaveOutDev - 1
            If WaveOut.GetWaveOutInfo(CType(x, UInteger), WaveOutInf) Then
                ComboBox2.Items.Add(WaveOutInf.ProductName)
            End If
        Next

        ComboBox2.SelectedIndex = 0
        Timer2.Start()
        player1.SetPlayerVolume(45, 45)
        player2.SetPlayerVolume(45, 45)
        player3.SetPlayerVolume(45, 45)
        player4.SetPlayerVolume(45, 45)
        player5.SetPlayerVolume(45, 45)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.Enabled = False
        Dim DeviceID As Integer = ComboBox1.SelectedIndex
        midiInOpen(hMidiIn, DeviceID, ptrCallback, 0, CALLBACK_FUNCTION Or MIDI_IO_STATUS)
        midiInStart(hMidiIn)
        MonitorActive = True
        Button2.Text = "Stop monitor"
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        TextBox1.Clear()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If MonitorActive = False Then
            midiInStart(hMidiIn)
            MonitorActive = True
            Button2.Text = "Stop monitor"
        Else
            midiInStop(hMidiIn)
            MonitorActive = False
            Button2.Text = "Start monitor"
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        opnmd.ShowDialog()
        modepath = opnmd.FileName
        If (modepath IsNot "") Then
            Soundpaths.Clear()
            Soundlines.Clear()
            modelines = New List(Of String)(IO.File.ReadAllLines(modepath))
            If modelines.Count > 1 Then
                Dim Soundlists() As String
                If Not (modelines.Count() - 1) > mode Then
                    mode = 1
                End If
                Soundlists = modelines.Item(mode).Split(";")
                If modelines.Item(mode) IsNot "" Then
                    Soundlines = New List(Of String)(IO.File.ReadAllLines(Soundlists(1)))
                    If (Soundlines.ToArray.Length > 0) Then
                        For i As Integer = 0 To 63
                            If i > 3 Then
                                If ((Soundlines.ToArray.Length) > i) Then
                                    Dim Soundinfo() As String = Soundlines.Item(i).Split(";")
                                    DirectCast(Me.Controls("Button" & i), Button).Text = Soundinfo(0)
                                    If Soundinfo.Length > 1 Then
                                        Soundpaths.Add(Soundinfo(1))
                                    Else
                                        Soundpaths.Add("")
                                    End If
                                End If
                            Else
                                Soundpaths.Add("")
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MonitorActive = False
        midiInStop(hMidiIn)
        midiInReset(hMidiIn)
        'midiInClose(hMidiIn)
        Application.Exit()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        RadioButton1.Checked = False
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Button9.Enabled = True
        Button10.Enabled = True
        Button11.Enabled = True
        Button12.Enabled = True
        Button13.Enabled = True
        Button14.Enabled = True
        Button15.Enabled = True
        Button16.Enabled = True
        Button17.Enabled = True
        Button18.Enabled = True
        Button19.Enabled = True
        Button20.Enabled = True
        Button21.Enabled = True
        Button22.Enabled = True
        Button23.Enabled = True
        Button24.Enabled = True
        Button25.Enabled = True
        Button26.Enabled = True
        Button27.Enabled = True
        Button28.Enabled = True
        Button29.Enabled = True
        Button30.Enabled = True
        Button31.Enabled = True
        Button32.Enabled = True
        Button33.Enabled = True
        Button34.Enabled = True
        Button35.Enabled = True
        Button36.Enabled = True
        Button37.Enabled = True
        Button38.Enabled = True
        Button39.Enabled = True
        Button40.Enabled = True
        Button41.Enabled = True
        Button42.Enabled = True
        Button43.Enabled = True
        Button44.Enabled = True
        Button45.Enabled = True
        Button46.Enabled = True
        Button47.Enabled = True
        Button48.Enabled = True
        Button49.Enabled = True
        Button50.Enabled = True
        Button51.Enabled = True
        Button52.Enabled = True
        Button53.Enabled = True
        Button54.Enabled = True
        Button55.Enabled = True
        Button56.Enabled = True
        Button57.Enabled = True
        Button58.Enabled = True
        Button59.Enabled = True
        Button60.Enabled = True
        Button61.Enabled = True
        Button62.Enabled = True
        Button63.Enabled = True
        Button64.Enabled = True
        Button65.Enabled = True
        Button66.Enabled = True
        Button67.Enabled = True
        Button68.Enabled = True

        Timer1.Stop()
        Timer1.Enabled = False
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        player1.GetPlayerVolume(ProgressBar1.Value, ProgressBar2.Value)
        player2.GetPlayerVolume(ProgressBar3.Value, ProgressBar4.Value)
        player3.GetPlayerVolume(ProgressBar5.Value, ProgressBar6.Value)
        If CheckBox1.Checked Then
            player1.SetPlayerVolume(0, 0)
        End If
        If CheckBox2.Checked Then
            player2.SetPlayerVolume(0, 0)
        End If
        If CheckBox3.Checked Then
            player3.SetPlayerVolume(0, 0)
        End If

        player4.GetPlayerVolume(ProgressBar7.Value, ProgressBar8.Value)
        player5.GetPlayerVolume(ProgressBar9.Value, ProgressBar10.Value)
        If CheckBox4.Checked Then
            player4.SetPlayerVolume(0, 0)
        End If
        If CheckBox5.Checked Then
            player5.SetPlayerVolume(0, 0)
        End If
    End Sub
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        player1.SetPlayerVolume(TrackBar1.Value, TrackBar1.Value)
    End Sub

    Private Sub TrackBar3_Scroll(sender As Object, e As EventArgs) Handles TrackBar3.Scroll
        player2.SetPlayerVolume(TrackBar3.Value, TrackBar3.Value)
    End Sub

    Private Sub TrackBar5_Scroll(sender As Object, e As EventArgs) Handles TrackBar5.Scroll
        player3.SetPlayerVolume(TrackBar5.Value, TrackBar5.Value)
    End Sub
    Private Sub TrackBar7_Scroll(sender As Object, e As EventArgs) Handles TrackBar7.Scroll
        player4.SetPlayerVolume(TrackBar7.Value, TrackBar7.Value)
    End Sub

    Private Sub TrackBar9_Scroll(sender As Object, e As EventArgs) Handles TrackBar9.Scroll
        player5.SetPlayerVolume(TrackBar9.Value, TrackBar9.Value)
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        player1.SetPitch(TrackBar2.Value)
    End Sub

    Private Sub TrackBar4_Scroll(sender As Object, e As EventArgs) Handles TrackBar4.Scroll
        player2.SetPitch(TrackBar4.Value)
    End Sub

    Private Sub TrackBar6_Scroll(sender As Object, e As EventArgs) Handles TrackBar6.Scroll
        player3.SetPitch(TrackBar6.Value)
    End Sub

    Private Sub TrackBar8_Scroll(sender As Object, e As EventArgs) Handles TrackBar8.Scroll
        player4.SetPitch(TrackBar8.Value)
    End Sub

    Private Sub TrackBar10_Scroll(sender As Object, e As EventArgs) Handles TrackBar10.Scroll
        player5.SetPitch(TrackBar10.Value)
    End Sub

    Private Sub Button69_Click(sender As Object, e As EventArgs) Handles Button69.Click
        player1.StopPlayback()
    End Sub

    Private Sub Button71_Click(sender As Object, e As EventArgs) Handles Button71.Click
        player2.StopPlayback()
    End Sub

    Private Sub Button73_Click(sender As Object, e As EventArgs) Handles Button73.Click
        player3.StopPlayback()
    End Sub

    Private Sub Button75_Click(sender As Object, e As EventArgs) Handles Button75.Click
        player4.StopPlayback()
    End Sub

    Private Sub Button77_Click(sender As Object, e As EventArgs) Handles Button77.Click
        player5.StopPlayback()
    End Sub

    Private Sub Button70_Click(sender As Object, e As EventArgs) Handles Button70.Click
        player1.SetPitch(100)
        TrackBar2.Value = 100
    End Sub

    Private Sub Button72_Click(sender As Object, e As EventArgs) Handles Button72.Click
        player2.SetPitch(100)
        TrackBar4.Value = 100

    End Sub

    Private Sub Button74_Click(sender As Object, e As EventArgs) Handles Button74.Click
        player3.SetPitch(100)
        TrackBar6.Value = 100

    End Sub

    Private Sub Button76_Click(sender As Object, e As EventArgs) Handles Button76.Click
        player4.SetPitch(100)
        TrackBar8.Value = 100

    End Sub

    Private Sub Button78_Click(sender As Object, e As EventArgs) Handles Button78.Click
        player5.SetPitch(100)
        TrackBar10.Value = 100

    End Sub
End Class