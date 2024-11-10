Const ИмяФайлаШаблона = "Шаблон.dotm"
Const КоличествоОбрабатываемыхСтолбцов = 30
Const РасширениеСоздаваемыхФайлов = ".docx"

Sub СформироватьДокументы()
    ПутьШаблона = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, ИмяФайлаШаблона)
    НоваяПапка = NewFolderName & Application.PathSeparator
    r = Cells(Rows.Count, "A").End(xlUp).row: rc = r - 2
      
    Dim AppWord As Object, DocWord As Object: Set AppWord = CreateObject("Word.Application")

    For Each row In ActiveSheet.Rows("3:" & r)
        With row
            Документ = "Документ " & Trim$(.Cells(1))
            Filename = НоваяПапка & Документ & РасширениеСоздаваемыхФайлов

            Set DocWord = AppWord.Documents.Add(ПутьШаблона): DoEvents

            For i = 1 To КоличествоОбрабатываемыхСтолбцов
                FindText = Cells(1, i): ReplaceText = Trim$(.Cells(i))

                With DocWord.Range.Find
                    .Text = FindText
                    .Replacement.Text = ReplaceText
                    .Forward = True
                    .Wrap = 1
                    .Format = False: .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=2
                End With
                                          
                DoEvents

            Next i
            DocWord.SaveAs Filename: DocWord.Close False: DoEvents
           
 
                            ' Открытие документов и удаление лишних строк и смволов
                            Dim AppWord1
                            Set AppWord1 = CreateObject("Word.Application")
                            AppWord1.Visible = False
                            Set DocWord1 = AppWord1.Documents.Open(НоваяПапка & Документ & РасширениеСоздаваемыхФайлов)
                            DocWord1.Activate
                            AppWord1.Run "DeletingRows"
                            AppWord1.Run "Найти_и_заменить"
                            DocWord1.Close saveChanges:=True
                            AppWord1.Quit
                            Set AppWord1 = Nothing: Set DocWord1 = Nothing
                            End With
     

    Next row


    
    AppWord.Quit
    msg = "Сформировано " & rc & " документов. Все они находятся в папке" & vbNewLine & НоваяПапка
    MsgBox msg, vbInformation, "Выполнено"
    
    
    
End Sub


'Проверка существования папки с документами
Function NewFolderExists(strPathName As String) As Boolean
    Dim strFolder As String
    strFolder = Dir(strPathName, vbDirectory)
    If (Len(strFolder) = 0 Or Err = 76) Then
    NewFolderExists = False
    Else
    NewFolderExists = True
    End If
End Function


'Вывод сообщения о том, что документы сегодня уже формировались
Function NewFolderName() As String
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Акты " & Get_Date)
    If NewFolderExists(NewFolderName) Then
    MsgBox "Документы сегодня уже формировались!", vbCritical: End
    Else
    MkDir NewFolderName
    End If
End Function

Function Get_Date() As String: Get_Date = Replace(Replace(DateValue(Now), "/", "-"), ".", "-"): End Function

