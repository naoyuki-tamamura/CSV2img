Imports System
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Module Module1

    Sub Main(args As String())
        Dim SourceFilePath As String = Nothing
        If args.Length = 0 Then
            Console.WriteLine("バイナリに変換するCSVファイルを指定してください。")
            Environment.Exit(1)
        Else
            SourceFilePath = args(0)
        End If

        Dim BaseFileName As String = System.IO.Path.GetFileNameWithoutExtension(SourceFilePath)
        Dim DestFilePath As String = System.IO.Path.GetDirectoryName(SourceFilePath) & "\" & BaseFileName & ".img"

        If System.IO.File.Exists(SourceFilePath) = False Then
            Console.WriteLine(SourceFilePath & "が見つかりません。")
            Environment.Exit(2)
        End If

        If System.IO.File.Exists(DestFilePath) Then
            Console.WriteLine("同じフォルダに '" + BaseFileName & ".img" + "' が存在するため、処理を中止します。")
            Environment.Exit(2)
        End If

        Dim arrPixelValueText()() As String = ReadCsv(SourceFilePath, False, False)

        Dim MatrixX As Long = UBound(arrPixelValueText(0))
        Dim MatrixY As Long = UBound(arrPixelValueText)

        Dim CurrentX As Long = 0
        Dim CurrentY As Long = 0

        Using stream As Stream = File.OpenWrite(DestFilePath)
            ' streamに書き込むためのBinaryWriterを作成
            Using writer As New BinaryWriter(stream)
                For CurrentY = 0 To MatrixY
                    For CurrentX = 0 To MatrixX
                        ' 4バイトのバイト配列を書き込む
                        writer.Write(Single.Parse(arrPixelValueText(CurrentY)(CurrentX)))
                    Next
                Next
            End Using
        End Using
        Console.WriteLine("Done!")

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' 
    ''' CSVファイルの読込処理
    ''' 
    ''' ファイル名
    ''' 区切りの指定(True:タブ区切り, False:カンマ区切り)
    ''' 引用符フラグ(True:引用符で囲まれている, False:囲まれていない)
    ''' 読込結果の文字列の2次元配列
    ''' -----------------------------------------------------------------------------
    Private Function ReadCsv(ByVal astrFileName As String, ByVal ablnTab As Boolean, ByVal ablnQuote As Boolean) As String()()
        ReadCsv = Nothing
        'ファイルStreamReader
        Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser = Nothing
        Try
            'Shift-JISエンコードで変換できない場合は「?」文字の設定
            Dim encFallBack As System.Text.DecoderReplacementFallback = New System.Text.DecoderReplacementFallback("?")
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("UTF-8", System.Text.EncoderFallback.ReplacementFallback, encFallBack)

            'TextFieldParserクラス
            parser = New Microsoft.VisualBasic.FileIO.TextFieldParser(astrFileName, enc)

            '区切りの指定
            parser.TextFieldType = FieldType.Delimited
            If ablnTab = False Then
                'カンマ区切り
                parser.SetDelimiters(",")
            Else
                'タブ区切り
                parser.SetDelimiters(vbTab)
            End If

            If ablnQuote = True Then
                'フィールドが引用符で囲まれているか
                parser.HasFieldsEnclosedInQuotes = True
            End If

            'フィールドの空白トリム設定
            parser.TrimWhiteSpace = False

            Dim strArr()() As String = Nothing
            Dim nLine As Integer = 0
            'ファイルの終端までループ
            While Not parser.EndOfData
                'フィールドを読込
                Dim strDataArr As String() = parser.ReadFields()

                '戻り値領域の拡張
                ReDim Preserve strArr(nLine)

                '退避
                strArr(nLine) = strDataArr
                nLine += 1
            End While

            '正常終了
            Return strArr

        Catch ex As Exception
            'エラー
            Console.Write("Error!")
        Finally
            '閉じる
            If parser IsNot Nothing Then
                parser.Close()
            End If
        End Try
    End Function

End Module
