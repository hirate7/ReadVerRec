Imports System.IO
Imports System.Text
Imports Ionic.Zip
Imports ReadVerRec.Maple01

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim S As String
        S = Dir("C:\Vup\*.001")
        Do
            If S = "" Then Exit Do
            lstVerRec.Items.Add(S)
            S = Dir()
        Loop

    End Sub

    Private Sub cmdExtractToDirectory_Click(sender As Object, e As EventArgs) Handles cmdExtractToDirectory.Click

        Dim I As Integer

        For I = 0 To lstVerRec.Items.Count - 1

            ' ZIPファイルのパス
            Dim ZipPath As String = "C:\Vup\" + lstVerRec.Items(I).ToString

            ' ファイルを書き出すフォルダーを作成する
            Dim folderPath As String = Strings.Left(ZipPath, InStr(ZipPath, ".") - 1)
            Directory.CreateDirectory(folderPath)

            'ZipFileを作成する 
            Dim options As New ReadOptions
            options.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
            Using zip As Ionic.Zip.ZipFile = Ionic.Zip.ZipFile.Read(ZipPath, options)
                'パスワードが設定されているときは、 
                zip.Password = "stardust"
                'または、ZipEntry.ExtractWithPasswordメソッドで展開する 

                '展開先に同名のファイルがあれば上書きする 
                zip.ExtractExistingFile = Ionic.Zip.ExtractExistingFileAction.OverwriteSilently
                'DoNotOverwriteで上書きしない。Throwで例外をスロー。既定値はThrow。 
                'ZipEntry.Extractに指定して展開することもできる 

                'フォルダ構造を無視する 
                'zip.FlattenFoldersOnExtract = True
                Try
                    FileSystem.Kill(folderPath + "\*.*")
                Catch ex As Exception

                End Try
                'ZIP書庫内のエントリを取得 
                For Each entry As Ionic.Zip.ZipEntry In zip
                    'エントリを展開する 
                    entry.Extract(folderPath)
                Next
            End Using

        Next

        Dim FNa As String
        Dim D() As String
        ReDim D(1)
        Dim DMax As Integer
        Dim St As Integer
        Dim S As String
        Dim C As Integer = 0

        'ファイルを上書きし、Shift JISで書き込む 
        Dim sw As New System.IO.StreamWriter("C:\DDD\VUPREC.CSV", False, System.Text.Encoding.GetEncoding("shift_jis"))

        '"C:\test"以下のサブフォルダをすべて取得する
        'ワイルドカード"*"は、すべてのフォルダを意味する
        Dim subFolders As String() = System.IO.Directory.GetDirectories("C:\Vup\", "*", System.IO.SearchOption.AllDirectories)
        For I = 0 To subFolders.Count - 1

            'ディレクトリの属性を取得する
            Dim di As New System.IO.DirectoryInfo(subFolders(I))

            '隠し属性があるか調べる
            If (di.Attributes And System.IO.FileAttributes.Hidden) = 0 Then

                S = Path.GetFileName(subFolders(I))

                'Select Case S.Substring(0, 9)
                '    Case "O20250710"
                '        'Case "H20250714", "P20250714"
                '    Case Else
                '        GoTo Loop2Last
                'End Select

                FNa = subFolders(I) + "\personal.ini"
                ReadIni2(FNa, "ユーザー", "郵便番号", D, DMax, St)
                S = S + "," + LeftBX(D(1), 8)
                ReadIni2(FNa, "ユーザー", "住所１", D, DMax, St)
                S = S + "," + LeftBX(D(1), 40)
                ReadIni2(FNa, "ユーザー", "屋号", D, DMax, St)
                S = S + "," + LeftBX(D(1), 20)
                ReadIni2(FNa, "ユーザー", "システム", D, DMax, St)
                Select Case D(1)
                    Case "大阪"
                        ReadIni2(FNa, "ユーザー", "新柔整師番号", D, DMax, St)
                        D(1) = "協27" + D(1) + "-0-0"
                    Case Else
                        ReadIni2(FNa, "ユーザー", "知事登録番号", D, DMax, St)
                End Select
                S = S + "," + D(1)
                If Dir(subFolders(I) + "\mprogkdata.mpi") <> "" Then
                    ReadIni2(subFolders(I) + "\mprogkdata.mpi", "[設定]", "マイナ資格確認アプリ連携", D, DMax, St)
                    If St = 9 Then D(1) = "0"
                    S = S + "," + CStr(Val(D(1)))
                Else
                    S = S + ",*"
                End If
                ReadIni2(subFolders(I) + "\mprog.ini", "[設定]", "お知らせインターネット受信", D, DMax, St)
                If St = 9 Then D(1) = "0"
                S = S + "," + CStr(Val(D(1)))

                ReadIni2(subFolders(I) + "\personal.ini", "ユーザー", "プリンター", D, DMax, St)
                S = S + "," + D(1)

                S = S + "," + GetExecDT(subFolders(I) + "\verrecnet.mpi")

                lstNa.Items.Add(S)

                sw.WriteLine(S)

                C += 1

Loop2Last:

            End If

        Next

        lstNa.Sorted = True

        lblCount.Text = CStr(C) + "件"

        '閉じる 
        sw.Close()

    End Sub

    Function GetExecDT(FNa As String) As String

        Dim line As String = ""
        Dim L As String

        Using sr As StreamReader = New StreamReader(FNa, Encoding.GetEncoding("Shift_JIS"))

            Do
                L = sr.ReadLine()
                If L Is Nothing Then Exit Do
                line = L
            Loop

        End Using

        Dim DT() As String
        DT = Split(line, ",")

        Return DT(0)

    End Function

    Function LeftBX(S As String, N As Integer) As String

        Return MidbX(S + Space(N), 1, N)

    End Function

    '
    '機能:半角文字を1、全角文字を2として数えて文字列の一部を返す
    '引渡:S             :基になる文字列
    '    :I             :何文字目からか
    '　　:J             :何文字返すか
    '戻り:S             :保存
    '    :I             :保存
    '    :J             :保存
    '    :LenbX         :文字列の一部
    Public Shared Function MidbX(ByVal S As String, ByVal I As Integer, Optional ByVal J As Integer = -1) As String

        Dim p, Q, L As Integer
        Dim SS, C As String

        If J = -1 Then J = LenbX(S)

        SS = ""
        Q = Len(S)
        L = 1
        For p = 1 To Q
            If I + J <= L Then Exit For
            C = Mid$(S, p, 1)
            If L >= I Then SS = SS + C
            L = L + HZCheck(C)
        Next p

        MidbX = SS

    End Function

    '
    '機能:半角文字を1、全角文字を2として数えた文字列の長さを返す
    '引渡:S             :長さを知りたい文字列
    '戻り:S             :保存
    '    :LenbX         :長さ
    Public Shared Function LenbX(ByVal S As String) As Integer

        Dim I As Integer
        Dim J As Integer
        Dim L As Integer
        Dim C As String

        J = Len(S)
        L = 0
        For I = 1 To J
            C = Mid$(S, I, 1)
            L = L + HZCheck(C)
        Next I

        LenbX = L

    End Function

    Public Shared Function HZCheck(ByVal C As String) As Integer
        '
        '
        '機能:半角文字か、全角文字かをチェックする
        '引渡:S             :チェックしたい文字　1文字

        '戻り:S             :保存

        '　　:HZCheck       :半角：1　全角：2
        '

        If Len(C) <> 1 Then Throw New Exception()

        'Shift_JISの＆H00から＆HFFである場合にTrueを返す

        If C <> "?" And Asc(C) = 63 Then
            'Shift_JISに無い文字
            Return 2
        ElseIf Asc(C) >= 0 And Asc(C) <= &HFF Then
            Return 1
        Else
            Return 2
        End If

        'If 0 <= AscW(C) And AscW(C) <= &HFF Then
        '    HZCheck = 1
        'Else
        '    HZCheck = 2
        'End If

    End Function
End Class
