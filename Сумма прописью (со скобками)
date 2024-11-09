Option Explicit


'Для получения, в ячейке листа, иходной суммы прописью,
'нужно: в ячейку вставить функцию СУММ_ПРОП_СКОБ_РУБ
'т.е. в ячейке будет прописано: = СУММ_ПРОП_СКОБ_РУБ (адрес_
'ячейки_откуда_берется_исходное_значение)


Public Function СУММ_ПРОП_СКОБ_РУБ(Исх_Значение As Currency) As String

Dim lngGroup As Long, lngCategory As Long, intLen As Integer
Dim strInWords As String, lngRest_str As String
Dim lngSum As Long, lngRest As Long

    If Исх_Значение = 0 Then
        СУММ_ПРОП_СКОБ_РУБ = "(ноль) рублей 00 копеек"
        Exit Function
    End If

    If Исх_Значение > Int(Исх_Значение) Then
            lngSum = Int(Исх_Значение)
        Else
            lngSum = Исх_Значение
    End If
    lngRest = lngSum
    If lngRest = 0 Then
        strInWords = "(ноль) рублей "
    Else
        lngGroup = lngRest \ 1000000
        'Прописываем сотни миллионов
        If lngGroup <> 0 Then
            lngCategory = lngGroup \ 100
            strInWords = strInWords & x_hundr(lngCategory)
            lngRest = lngRest - lngCategory * 100 * 1000000
            lngGroup = lngGroup - lngCategory * 100
    
            'Прописываем десятки миллионов
            If lngGroup > 19 Then
                lngCategory = lngGroup \ 10
                strInWords = strInWords & x_dec(lngCategory)
                lngRest = lngRest - lngCategory * 10 * 1000000
                lngGroup = lngGroup - lngCategory * 10
            End If
    
            'Прописываем единицы миллионов
            lngCategory = lngGroup
            strInWords = strInWords & x_edin(lngCategory, "m")
            lngRest = lngRest - lngCategory * 1000000
    
            strInWords = strInWords & x_mill(lngCategory)
        End If
    
        lngGroup = lngRest \ 1000
        If lngGroup <> 0 Then
            'Прописываем сотни тысяч
            lngCategory = lngGroup \ 100
            strInWords = strInWords & x_hundr(lngCategory)
            lngRest = lngRest - lngCategory * 100 * 1000
            lngGroup = lngGroup - lngCategory * 100
    
            'Прописываем десятки тысяч
            If lngGroup > 19 Then
                lngCategory = lngGroup \ 10
                strInWords = strInWords & x_dec(lngCategory)
                lngRest = lngRest - lngCategory * 10 * 1000
                lngGroup = lngGroup - lngCategory * 10
            End If
    
            'Прописываем единицы тысяч
            lngCategory = lngGroup
            strInWords = strInWords & x_edin(lngCategory, "w")
            lngRest = lngRest - lngCategory * 1000
    
            strInWords = strInWords & x_thousend(lngCategory)
        End If
        
        If lngRest = 0 Then
            lngCategory = lngRest
            strInWords = Left(strInWords, Len(strInWords) - 1) & x_rubli(lngCategory)
            GoTo 10
        End If
        lngGroup = lngRest
        
        If lngGroup <> 0 Then
            'Прописываем сотни
            lngCategory = lngGroup \ 100
            strInWords = strInWords & x_hundr(lngCategory)
            lngRest = lngRest - lngCategory * 100
            lngGroup = lngGroup - lngCategory * 100
    
            'Прописываем десятки
            If lngGroup > 19 Then
                lngCategory = lngGroup \ 10
                strInWords = strInWords & x_dec(lngCategory)
                lngRest = lngRest - lngCategory * 10
                lngGroup = lngGroup - lngCategory * 10
            End If
    
            'Прописываем единицы
            lngCategory = lngGroup
            strInWords = strInWords & x_edin(lngCategory, "m")
            lngRest = lngRest - lngCategory
        End If
        strInWords = Left(strInWords, Len(strInWords) - 1) & x_rubli(lngCategory)
    End If
    
    
    
10  lngRest = (Исх_Значение - lngSum) * 100
    lngRest = CInt(lngRest)
    If lngRest = 0 Then
        strInWords = strInWords & " 00 копеек"
    Else
        lngGroup = lngRest
        
        If lngGroup < 10 Then
            strInWords = strInWords & "0"
        End If
        
        If lngGroup > 19 Then
            lngGroup = lngRest \ 10
            lngCategory = lngRest - lngGroup * 10
        Else
            lngCategory = lngRest
        End If
            
        lngRest_str = lngRest
        strInWords = strInWords & lngRest_str & x_kopeiki(lngCategory)
    End If
    
    intLen = Len(strInWords)
    If IsNull(intLen) Then
       Exit Function
    End If
    
     
    



   СУММ_ПРОП_СКОБ_РУБ = "(" & strInWords
   
   
   
   End Function

Private Function x_dec(lngCategory As Long) As String

    Select Case lngCategory
         Case 2
            x_dec = "двадцать "
         Case 3
            x_dec = "тридцать "
         Case 4
            x_dec = "сорок "
         Case 5
            x_dec = "пятьдесят "
         Case 6
            x_dec = "шестьдесят "
         Case 7
            x_dec = "семьдесят "
         Case 8
            x_dec = "восемьдесят "
         Case 9
            x_dec = "девяносто "
    End Select

End Function

Private Function x_edin(lngCategory As Long, sort As String) As String

    Select Case lngCategory
        Case 1
            If sort = "m" Then
                x_edin = "один "
            Else
                x_edin = "одна "
            End If
        Case 2
            If sort = "m" Then
                x_edin = "два "
            Else
                x_edin = "две "
            End If
        Case 3
            x_edin = "три "
        Case 4
            x_edin = "четыре "
        Case 5
            x_edin = "пять "
        Case 6
            x_edin = "шесть "
        Case 7
            x_edin = "семь "
        Case 8
            x_edin = "восемь "
        Case 9
            x_edin = "девять "
        Case 10
            x_edin = "десять "
        Case 11
            x_edin = "одиннадцать "
        Case 12
            x_edin = "двенадцать "
        Case 13
            x_edin = "тринадцать "
        Case 14
            x_edin = "четырнадцать "
        Case 15
            x_edin = "пятнадцать "
        Case 16
            x_edin = "шестнадцать "
        Case 17
            x_edin = "семнадцать "
        Case 18
            x_edin = "восемнадцать "
        Case 19
            x_edin = "девятнадцать "

    End Select

End Function
Private Function x_mill(lngCategory As Long) As String

    If lngCategory = 1 Then
        x_mill = "миллион "
    ElseIf lngCategory > 1 And lngCategory < 5 Then
        x_mill = "миллиона "
    Else
        x_mill = "миллионов "
    End If

End Function

Private Function x_rubli(lngCategory As Long) As String
    
    If lngCategory = 1 Then
        x_rubli = ") рубль "
    ElseIf lngCategory > 1 And lngCategory < 5 Then
        x_rubli = ") рубля "
    Else
        x_rubli = ") рублей "
    End If

End Function
Private Function x_kopeiki(lngCategory As Long) As String
    
    If lngCategory = 1 Then
        x_kopeiki = " копейка"
    ElseIf lngCategory > 1 And lngCategory < 5 Then
        x_kopeiki = " копейки"
    Else
        x_kopeiki = " копеек"
    End If

End Function
    

Private Function x_hundr(lngCategory As Long) As String
    
    Select Case lngCategory
         Case 1
            x_hundr = "сто "
         Case 2
            x_hundr = "двести "
         Case 3
            x_hundr = "триста "
         Case 4
            x_hundr = "четыреста "
         Case 5
            x_hundr = "пятьсот "
         Case 6
            x_hundr = "шестьсот "
         Case 7
            x_hundr = "семьсот "
         Case 8
            x_hundr = "восемьсот "
         Case 9
            x_hundr = "девятьсот "
    End Select

End Function


    '
Private Function x_thousend(lngCategory As Long) As String

    If lngCategory = 1 Then
        x_thousend = "тысяча "
    ElseIf lngCategory > 1 And lngCategory < 5 Then
        x_thousend = "тысячи "
    Else
        x_thousend = "тысяч "
    End If

End Function
