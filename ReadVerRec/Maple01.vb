Public Class Maple01

    Public Shared Sub Read_INI(ByVal F As String, ByVal Section0 As String, ByRef D() As String, ByRef ST As Integer)
        '機能   :INI ファイルの順次読み込み 指定したセクションの内容を順次呼び出すのに使用する
        '       :セクションの最後まで呼ぶとST=-1を返しファイルはクローズされる
        '       :途中で止めるとファイルがオープンされたままになるので強制クローズする必要がある
        '引渡   :F      :ファイル名
        '       :Section:セクション名
        '       :D()    :１次元配列を指定する 添え字は０から
        '       :ST     :=0 順次読み込み
        '               :=1 強制クローズ
        '戻り   :F      :保存
        '       :Section:保存
        '       :D()    :読み込まれたINIファイルの中身      ST=1のときは保存
        '       :ST     :=0 正常 =-1 読み込めるデーターなし

        'Static File_Section(1 To 5) As String
        'Static FNo(1 To 5) As Integer
        Static File_Section(0 To 5) As String
        Static FNo(0 To 5) As Integer
        Static F_Pointer As Integer
        Dim I As Integer
        Dim DD As String

        '/// ST値異常 ///
        If ST <> 0 And ST <> 1 Then ST = -1 : Exit Sub

        '/// ファイルポインタ-検索 ///
        F = Trim$(F)
        Section0 = Trim$(Section0)
        If Left$(Section0, 1) = "[" Or Left$(Section0, 1) = "［" Then Section0 = Mid$(Section0, 2)
        If Right$(Section0, 1) = "]" Or Left$(Section0, 1) = "］" Then Section0 = Left$(Section0, Len(Section0) - 1)

        If F = "" Or Section0 = "" Then Throw New Exception()

        For F_Pointer = 1 To 5
            If F + "[" + Section0 + "]" = File_Section(F_Pointer) Then Exit For
        Next F_Pointer

        '/// 強制クローズ ///
        If ST = 1 Then
            If F_Pointer > 5 Then Throw New Exception()
            FileClose(FNo(F_Pointer)) : File_Section(F_Pointer) = "" : FNo(F_Pointer) = 0
            ST = 0 : Exit Sub
        End If

        '/// １つめ読み込み ///
        If F_Pointer > 5 Then
            For F_Pointer = 1 To 5
                If FNo(F_Pointer) = 0 Then

                    File_Section(F_Pointer) = F + "[" + Section0 + "]"
                    FNo(F_Pointer) = FreeFile()

                    FileOpen(FNo(F_Pointer), F, OpenMode.Input)

                    Do
                        If EOF(FNo(F_Pointer)) Then
                            FileClose(FNo(F_Pointer))
                            File_Section(F_Pointer) = ""
                            FNo(F_Pointer) = 0
                            ST = -1
                            Exit Sub
                        End If

                        DD = LineInput(FNo(F_Pointer))
                        I = InStr(DD, ";")
                        If I Then DD = Left$(DD, I - 1)

                        I = InStr(DD, "[")
                        If I = 0 Then I = InStr(DD, "［")
                        If I > 0 And Len(DD) > I Then
                            DD = Mid$(DD, I + 1)
                            I = InStr(DD, "]")
                            If I = 0 Then I = InStr(DD, "］")
                            If I > 1 Then
                                DD = Mid$(DD, 1, I - 1)
                                If Trim$(DD) = Trim$(Section0) Then Exit Do
                            End If
                        End If
                    Loop
                    Exit For
                End If
            Next F_Pointer
            If F_Pointer > 5 Then Throw New Exception()
        End If

        '/// ２つめ以降読み込み ///
        Do

            If EOF(FNo(F_Pointer)) Then
                FileClose(FNo(F_Pointer))
                File_Section(F_Pointer) = ""
                FNo(F_Pointer) = 0
                ST = -1
                Exit Sub
            End If

            DD = LineInput(FNo(F_Pointer))
            I = InStr(DD, ";")
            If I Then DD = Left$(DD, I - 1)
            DD = TrimX(DD)
            If Len(DD) > 0 Then
                If InStr(DD, "[") Or InStr(DD, "［") Then
                    FileClose(FNo(F_Pointer))
                    File_Section(F_Pointer) = ""
                    FNo(F_Pointer) = 0
                    F_Pointer = 0
                    ST = -1
                    Exit Sub
                End If
                I = InStr(DD, "=")
                If I > 0 Then Exit Do
                If I = 0 Then DD = DD + "=" : Exit Do
            End If

        Loop

        An_INI(DD, D)

        ST = 0

    End Sub

    Private Shared Sub An_INI(ByVal L As String, ByRef LL() As String)
        '機能:INI File の１行を分解 (,で区切られたもの)
        '引渡:L          :分解対象文字列
        '    :LL()       :
        '戻り:L          :不定
        '    :LL()       :分解結果
        Dim p, I As Integer
        For p = 0 To UBound(LL, 1)
            LL(p) = ""
        Next p

        p = InStr(L, "=")
        If p = 0 Then LL(0) = "" : Exit Sub
        LL(0) = TrimX(Left$(L, p - 1))
        L = Mid$(L, p + 1)
        I = 1

        Do
            If I > UBound(LL) Then ReDim Preserve LL(0 To I)
            p = InStr(L, ",")
            If p = 0 Then LL(I) = TrimX(L) : Exit Do
            LL(I) = TrimX(Left$(L, p - 1))
            L = Mid$(L, p + 1)
            I = I + 1
        Loop

    End Sub

    Public Shared Sub ReadIni2(ByVal F As String, ByVal Section0 As String, ByVal Key As String, ByRef D() As String, ByRef DMax As Integer, ByRef ST As Integer)
        '機能   Read_INI(MAPLE01.BAS)の様な連続読込はせず、
        '       セクション、キーを指定して読み込み、クローズする
        '引数   F       ファイル名
        '       Sec     セクション名
        '       Key     項目名
        '       D()
        '       DMax
        '       ST
        '戻り値 F       保存
        '       Sec     保存
        '       Key     保存
        '       D()     INIファイル内容
        '       DMax    D()の要素数
        '       ST      -1..エラー　0..正常終了  9..項目が読み込めない
        Dim ST0 As Integer
        Dim D0() As String
        ReDim D0(0)

        If Trim(Section0) = "" Or Trim(Key) = "" Then
            'エラー終了　セクション名なし
            ST = -1
            Exit Sub
        End If

        ST0 = 0
        Do
            Read_INI(F, Section0, D0, ST0)

            If ST0 <> 0 Then Exit Do

            If D0(0) = Key Then
                D = D0
                DMax = D.Length - 1
                ST0 = 1
                Read_INI(F, Section0, D0, ST0)
                ST = 0
                Return
            End If
        Loop

        ST = 9

    End Sub

    Public Shared Function TrimX(ByVal S As String) As String
        '機能   :文字列の両端からスペースとTABとNullを外す
        '引渡:S        :文字列
        '戻り:S        :保存

        '    :TrimX    :変換結果

        Dim L As Integer
        Dim I As Integer
        Dim D As String
        Dim S0 As String

        S0 = S
        L = Len(S0)
        For I = 1 To L
            D = Mid$(S0, I, 1)
            If D <> " " And D <> "　" And D <> Chr(9) And D <> Chr(0) Then Exit For
        Next I
        If I = L + 1 Then
            Return ""
        End If

        S0 = Mid$(S0, I)

        L = Len(S0)
        For I = L To 1 Step -1
            D = Mid$(S0, I, 1)
            If D <> " " And D <> "　" And D <> Chr(9) And D <> Chr(0) Then Exit For
        Next I
        S0 = Left$(S0, I)
        Return S0

    End Function

End Class
