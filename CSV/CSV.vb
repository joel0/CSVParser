Public Module CSV

    Public Function Parse(CSV As String) As String()
        Dim Cursor As Integer = 0
        Dim Fields As New List(Of String)

        Dim FieldLength As Integer
        Dim NextQuote As Integer
        Do Until Cursor >= CSV.Length
            ' Unescaped field
            If Not CSV(Cursor) = """" Then
                ' Find the length of the field ended by a comma
                FieldLength = InStr(Cursor + 1, CSV, ",") - 1 - Cursor
                ' If there's no comma, this is the last field and the length goes to the end
                If FieldLength < 0 Then
                    FieldLength = CSV.Length - Cursor
                End If
                ' Add the field to the output
                Fields.Add(CSV.Substring(Cursor, FieldLength).Replace("""""", """"))
                ' Update the cursor
                Cursor += FieldLength + 1
            Else ' Escaped field by quotation mark
                Cursor += 1
                ' Find the next quotation mark, possibly the end of the field
                NextQuote = InStr(Cursor + 1, CSV, """") - 1 - Cursor
                Do While Cursor + NextQuote + 1 < CSV.Length AndAlso CSV(Cursor + NextQuote + 1) = """"
                    ' Find the next quotation mark, skipping over the current two
                    NextQuote = InStr(Cursor + NextQuote + 3, CSV, """") - 1 - Cursor
                Loop
                ' The field length without the quote character
                FieldLength = NextQuote '- 1
                ' Add the field to the output
                Fields.Add(CSV.Substring(Cursor, FieldLength).Replace("""""", """"))
                ' Update the cursor
                Cursor += NextQuote + 2 ' skip over the ending quote and the next comma. We aren't checking that the comma exists, bad practice...
            End If
        Loop

        Return Fields.ToArray
    End Function

    Public Function SplitLines(CSVText As String) As String()
        Dim CSVLines As New List(Of String)
        Dim lines() As String = CSVText.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        For x As Integer = 0 To lines.Count - 1
            If InStr(lines(x), """") = 0 Then
                CSVLines.Add(lines(x))
            Else
                Dim cursor = InStr(lines(x), """")
                Dim escaped As Boolean = True
                Dim tempLine As String = ""

                Do Until escaped = False
                    Do
                        cursor = InStr(cursor + 1, lines(x), """")
                        If cursor > 0 Then
                            escaped = Not escaped
                        End If
                    Loop While cursor > 0

                    tempLine += lines(x) & vbCrLf

                    If escaped = True Then
                        x += 1
                        If x >= lines.Count Then
                            Throw New Exception("CSV escape sequence not ended.")
                        End If
                    End If
                Loop

                If tempLine.Length > 0 Then
                    ' Trim the vbCrLf that's used between lines.  One was added assuming that there are more lines coming.
                    tempLine = Mid(tempLine, 1, tempLine.Length - 2)
                End If
                CSVLines.Add(tempLine)
            End If
        Next

        Return CSVLines.ToArray
    End Function
End Module
