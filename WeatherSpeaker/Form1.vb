Imports System.IO
Imports System.Media
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Speech.Synthesis
Imports System.Runtime.InteropServices

Public Class Form1
    Private wc As WebClient
    Private synth As SpeechSynthesizer
    Private enc As Encoding = Encoding.GetEncoding("UTF-8")
    Private t1 As New Thread(New ThreadStart(AddressOf GetTJMA))

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        wc = New WebClient
        Try
            synth = New SpeechSynthesizer
        Catch ex As COMException
            MessageBox.Show("Microsoft Speech Platform がインストールされていません。")
            Me.Close()
        End Try
    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        Dim canSpeakJapanese = False
        Try
            For Each v As InstalledVoice In synth.GetInstalledVoices()
                If v.VoiceInfo.Culture.Name.Equals("ja-JP") Then
                    synth.SelectVoice(v.VoiceInfo.Name)
                    canSpeakJapanese = True
                    Exit For
                End If
            Next
        Catch ex As PlatformNotSupportedException
            MessageBox.Show("Microsoft Speech Platform Runtime Languages がインストールされていません。")
        End Try

        If Not canSpeakJapanese Then
            MessageBox.Show("私は日本語を喋れません")
            Return
        End If

        t1.Start()
    End Sub

    Private Sub GetTJMA()
        Dim data As Byte() = wc.DownloadData("http://www.jma.go.jp/jp/yoho/319.html")
        Dim html As String = enc.GetString(data)
        html = Regex.Replace(html, "\n", "")
        Dim r As New Regex("<pre class=""textframe"">.*</pre>", RegexOptions.Multiline)
        Dim m As Match = r.Match(html)
        Dim ret As String = ""
        If m.Success Then
            ret = m.Value.Replace("<pre class=""textframe"">", "").Replace("</pre>", "")
            ret = ret.Replace("<b>", "").Replace("</b>", "")
            ret = Regex.Replace(ret, ".*【東京地方】", "東京地方、")
        Else
            ret = "データを取得できませんでした。"
        End If
        Dim o As New MemoryStream()
        synth.SetOutputToWaveStream(o)
        synth.Speak(ret)
        Speak(New MemoryStream(o.ToArray))
    End Sub

    Private Sub Speak(ByVal o As Stream)
        Dim player As New Media.SoundPlayer(o)
        player.PlaySync()
        player.Dispose()
    End Sub

    Private Sub Form1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If synth IsNot Nothing Then
            synth.Dispose()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        t1.Abort()
        Me.Close()
    End Sub
End Class
