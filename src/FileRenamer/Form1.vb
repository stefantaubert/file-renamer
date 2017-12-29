Imports System.IO
Imports System.DateTime
Imports System.Security.Cryptography
Imports System.Text

Public Class Form1
    Public Zaehler As Integer = 0
    Public Zahl As Integer = 0
    Public EndungText As String
    Dim Endung, Endung2 As String

    Public Function Entschlüsseln(ByVal Input As String, ByVal Schlüssel As Object) As String
        Dim Manager As New RijndaelManaged
        Dim Leange As Integer = 16
        Dim Provider As New MD5CryptoServiceProvider
        Dim Schluessel() As Byte = Provider.ComputeHash(Encoding.UTF8.GetBytes(Schlüssel))
        Provider.Clear()
        Dim encdata() As Byte = Convert.FromBase64String(Input)
        Dim Stream As New MemoryStream(encdata)
        Dim iv(15) As Byte
        Stream.Read(iv, 0, Leange)
        Manager.IV = iv
        Manager.Key = Schluessel
        Dim cs As New CryptoStream(Stream, Manager.CreateDecryptor, CryptoStreamMode.Read)
        Dim data(Stream.Length - Leange) As Byte
        Dim I As Integer = cs.Read(data, 0, data.Length)
        Entschlüsseln = System.Text.Encoding.UTF8.GetString(data, 0, I)
        cs.Close()
        Manager.Clear()
    End Function
    Public Function Verschlüsseln(ByVal Input As String, ByVal Schlüssel As Object) As String
        Dim Manager As New RijndaelManaged
        Dim Provider As New MD5CryptoServiceProvider
        Dim Schluessel() As Byte = Provider.ComputeHash(Encoding.UTF8.GetBytes(Schlüssel))
        Provider.Clear()
        Manager.Key = Schluessel
        Manager.GenerateIV()
        Dim iv() As Byte = Manager.IV
        Dim Stream As New MemoryStream
        Stream.Write(iv, 0, iv.Length)
        Dim cs As New CryptoStream(Stream, Manager.CreateEncryptor, CryptoStreamMode.Write)
        Dim data() As Byte = System.Text.Encoding.UTF8.GetBytes(Input)
        cs.Write(data, 0, data.Length)
        cs.FlushFinalBlock()
        Dim encdata() As Byte = Stream.ToArray()
        Verschlüsseln = Convert.ToBase64String(encdata)
        cs.Close()
        Manager.Clear()
    End Function
    Function ZusammengesetzterName(ByVal Schema As String, ByVal Originalname As String, ByVal Umbennen As Boolean) As String
        ZusammengesetzterName = ""
        For j As Integer = 0 To Schema.Length - 1
            If Schema.ElementAt(j) = "Z" Then
                ZusammengesetzterName &= Zaehler
            ElseIf Schema.ElementAt(j) = " " Then
                ZusammengesetzterName &= " "
            ElseIf Schema.ElementAt(j) = "T" Then
                ZusammengesetzterName &= TextBox4.Text
            ElseIf Schema.ElementAt(j) = "D" Then
                ZusammengesetzterName &= Date.Today
            ElseIf Schema.ElementAt(j) = "M" Then
                ZusammengesetzterName &= Format(Date.Now.Month, "00")
            ElseIf Schema.ElementAt(j) = "O" Then
                ZusammengesetzterName &= Originalname
            ElseIf Schema.ElementAt(j) = "U" Then
                ZusammengesetzterName &= Now.Hour & "." & Now.Minute & "." & Now.Second
            ElseIf Schema.ElementAt(j) = "J" Then
                ZusammengesetzterName &= Now.Year
            ElseIf Schema.ElementAt(j) = "e" Then
                If Not Umbennen Then
                    ZusammengesetzterName &= "." & TextBox5.Text
                End If
            ElseIf Schema.ElementAt(j) = "E" Then
                If Not Umbennen Then
                    ZusammengesetzterName &= ".ext"
                End If
            Else
                ZusammengesetzterName &= "???"
            End If
        Next
        If Not Umbennen Then
            Label6.Text = ZusammengesetzterName
        End If
    End Function
    'Sub HIdee()
    '    TextBox1.Hide()
    '    TextBox2.Hide()
    '    TextBox3.Hide()
    '    TextBox4.Hide()
    '    TextBox5.Hide()
    '    Label1.Hide()
    '    Label2.Hide()
    '    Label3.Hide()
    '    Label4.Hide()
    '    Label5.Hide()
    '    Label6.Hide()
    '    Label7.Hide()
    '    Label9.Hide()
    '    Button10.Hide()
    '    Button11.Hide()
    '    Button12.Hide()
    '    Button3.Hide()
    '    Button2.Hide()
    '    Button4.Hide()
    '    Button5.Hide()
    '    Button6.Hide()
    '    Button7.Hide()
    '    Button8.Hide()
    '    Button9.Hide()
    '    CheckBox1.Hide()
    '    GroupBox2.Hide()
    'End Sub
    Function Ordnername(ByVal Pfad As String) As String
        Dim ds As New FileInfo(Pfad)
        Dim Name As String
        Name = ds.DirectoryName()
        Name = Name.Substring(Name.LastIndexOf("\"), Name.Length - Name.LastIndexOf("\"))
        Ordnername = Name.Substring(1, Name.Length - 1)
    End Function
    Sub RenameName(ByVal Pfad As String, ByRef NeuerPfad As String, ByVal Schema As String, ByRef AuchUnterordner As Boolean)
        Button1.Text = "Abbrechen"
        Zaehler = 0
        ProgressBar1.Value = 0
        ProgressBar2.Value = 0
        Dim files() As String = Directory.GetFiles(Pfad)
        Dim files2() As String = Directory.GetDirectories(Pfad)
        Dim file_new, Originalname, DateiPfad, SchemaName As String
        Dim DirektUmbennen As Boolean = False
        files = Directory.GetFiles(Pfad)
        ProgressBar1.Maximum = files.Length
        If TextBox1.Text = TextBox3.Text Then
            If MsgBox("Wollen Sie die Datein direkt umbennen?", MsgBoxStyle.YesNo, "Frage") = MsgBoxResult.Yes Then
                DirektUmbennen = True
            Else
                DirektUmbennen = False
            End If
        End If
        For Each filepath As String In files
            ProgressBar1.Value += 1
            Zaehler += 1
            Endung2 = filepath.Substring(filepath.LastIndexOf("."), filepath.Length - filepath.LastIndexOf("."))
            DateiPfad = filepath.Substring(0, filepath.LastIndexOf("."))
            Originalname = DateiPfad.Substring(DateiPfad.LastIndexOf("\"), filepath.LastIndexOf(".") - (DateiPfad.LastIndexOf("\")))
            Originalname = Originalname.Substring(1, Originalname.Length - 1)
            SchemaName = ZusammengesetzterName(Schema, Originalname, True)
            If TextBox5.Text = "" Then
                file_new = SchemaName & Endung2
            Else
                file_new = SchemaName & "." & TextBox5.Text
            End If
            If DirektUmbennen = True Then
                FileSystem.Rename(filepath, NeuerPfad & "\" & file_new)
            Else
                System.IO.File.Copy(filepath, NeuerPfad & "\" & file_new)
            End If
            SchemaName = ""
            Label8.Text = "Umbenennungsfortschritt Datein " & Zaehler & "/" & files.Length
            Refresh()
            Update()
        Next
        If files2.Length > 0 Then
            For i As Integer = 0 To files2.Length - 1
                files = Directory.GetFiles(files2(i))
                ProgressBar2.Value = 0
                Zaehler = 0
                ProgressBar2.Maximum = files.Length
                For Each filepath As String In files
                    ProgressBar2.Value += 1
                    Zaehler += 1
                    Endung2 = filepath.Substring(filepath.LastIndexOf("."), filepath.Length - filepath.LastIndexOf("."))
                    DateiPfad = filepath.Substring(0, filepath.LastIndexOf("."))
                    Originalname = DateiPfad.Substring(DateiPfad.LastIndexOf("\"), filepath.LastIndexOf(".") - (DateiPfad.LastIndexOf("\")))
                    Originalname = Originalname.Substring(1, Originalname.Length - 1)
                    SchemaName = ZusammengesetzterName(Schema, Originalname, True)
                    Directory.CreateDirectory(TextBox3.Text & "\" & Ordnername(filepath))
                    If TextBox5.Text = "" Then
                        file_new = SchemaName & Endung2
                    Else
                        file_new = SchemaName & "." & TextBox5.Text
                    End If
                    Dim neuerPfadd As String = NeuerPfad & "\" & Ordnername(filepath) & "\" & file_new
                    If DirektUmbennen = True Then
                        FileSystem.Rename(filepath, neuerPfadd)
                    Else
                        System.IO.File.Copy(filepath, neuerPfadd)
                    End If
                    SchemaName = ""
                    Label10.Text = "Umbennenungsfortschritt in Ordnern " & Zaehler & "/" & files.Length
                    Refresh()
                    Update()
                Next
            Next
        End If
        MsgBox("Alle Datein wurden erfolgreich umbenannt.", MsgBoxStyle.Information, "Glückwunsch!")
        Dim f As New Process
        f.StartInfo.FileName = TextBox3.Text
        f.Start()
        Button1.Text = "Umbennen"
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Directory.Exists(TextBox1.Text) And Directory.Exists(TextBox3.Text) Then
            Try
                If CheckBox1.Checked Then
                    RenameName(TextBox1.Text, TextBox3.Text, TextBox2.Text, True)
                Else
                    RenameName(TextBox1.Text, TextBox3.Text, TextBox2.Text, False)
                End If
            Catch ex As Exception
                MsgBox("Das Schema verursacht einen konstanten Dateinamen. Bitte fügen Sie eine Zahl ein! Es kann aber auch sein das ein Dateiname der umbenannt werden soll mehr als 260 Zeichen hat!. Genaueres hier:" & vbNewLine & vbNewLine & ex.ToString, MsgBoxStyle.Information, "Achtung")
            End Try
        Else
            MsgBox("Der angegebene Pfad existiert nicht!", MsgBoxStyle.Information, "Achtung")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged, TextBox4.TextChanged
        ZusammengesetzterName(TextBox2.Text, "Originalname", False)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text &= "Z"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox2.Text &= " "
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox2.Text &= "T"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox2.Text &= "D"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox2.Text &= "O"
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            TextBox2.Text = TextBox2.Text.Remove(TextBox2.TextLength - 1, 1)
        Catch ex As Exception
            MsgBox("Es ist nichts vorhanden zum Löschen!", MsgBoxStyle.Information, "Hinweis")
        End Try
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox2.Text &= "U"
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        TextBox2.Text &= "J"
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Close()
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ErstelleDateiInformationen()
    End Sub
    Sub ErstelleDateiInformationen()
        Dim Filenum As Integer = FreeFile()
        FileOpen(Filenum, "C:\Renamer\renamer.info", OpenMode.Output)
        FileClose()
        Dim Path As String = "C:\Renamer\renamer.info"
        Dim sw As StreamWriter

        sw = File.AppendText(Path)
        sw.WriteLine(Verschlüsseln(TextBox1.Text, 123456789))
        sw.WriteLine(Verschlüsseln(TextBox3.Text, 123456789))
        sw.WriteLine(Verschlüsseln(TextBox2.Text, 123456789))
        sw.WriteLine(Verschlüsseln(TextBox4.Text, 123456789))
        sw.Flush()
        sw.Close()
    End Sub
    Sub LadeInformationen()
        Dim Path As String = "C:\Renamer\renamer.info"
        If Not File.Exists(Path) Then
            Dim f As DirectoryInfo
            Dim sf As Stream
            f = Directory.CreateDirectory("C:\Renamer\")
            sf = File.Create(Path)
            sf.Dispose()
            sf.Close()
            Return
        End If
        Dim Quelle As New Microsoft.VisualBasic.FileIO.TextFieldParser("C:\Renamer\renamer.info")
        Dim Zeile1 As String = Entschlüsseln(Quelle.ReadLine(), 123456789)
        Dim Zeile2 As String = Entschlüsseln(Quelle.ReadLine(), 123456789)
        Dim Zeile3 As String = Entschlüsseln(Quelle.ReadLine(), 123456789)
        Dim Zeile4 As String = Entschlüsseln(Quelle.ReadLine(), 123456789)
        TextBox1.Text = Zeile1
        TextBox2.Text = Zeile3
        TextBox3.Text = Zeile2
        TextBox4.Text = Zeile4
        Quelle.Close()
        Quelle.Dispose()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GroupBox2.Hide()
        LadeInformationen()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If MsgBox("Wollen Sie die Extension beibehalten? Bei NEIN klicken Sie dann bitte nach eingeben der neuen Extension auf OK", MsgBoxStyle.YesNo, "Frage") = MsgBoxResult.Yes Then
            TextBox2.AppendText("E")
        Else
            GroupBox2.Show()
        End If
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox3.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        TextBox2.AppendText("e")
        Endung = TextBox5.Text
        GroupBox2.Hide()
    End Sub
    Sub AuftauchenLassen()
        Timer2.Interval = 10
        Timer2.Tag = 0
        Timer2.Start()
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Timer1.Tag = 0 Then
            Timer1.Stop()
        Else
            Me.Opacity = Timer1.Tag - 0.1
            Timer1.Tag = Me.Opacity.ToString
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If Timer2.Tag > 1 Then
            Timer2.Stop()
        Else
            Me.Opacity = Timer2.Tag + 0.1
            Timer2.Tag = Me.Opacity.ToString
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Refresh()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        TextBox1.Text = Clipboard.GetText
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        If MsgBox("Wollen Sie ein (2) an den Zielort anhängen? (z.B. Neuer Ordner (2))", MsgBoxStyle.YesNo, "Frage") = MsgBoxResult.Yes Then
            TextBox3.Text = Clipboard.GetText & " (2)"
        Else
            TextBox3.Text = Clipboard.GetText
        End If
    End Sub
End Class
