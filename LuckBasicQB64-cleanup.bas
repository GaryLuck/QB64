$Debug
'---script created on 11-08-2022 14:51:59 by
' LuckBasic
' This is my personal project to build a Basic interpreter.
' It's a high level project to prove to myself that i understand how to build an interpreter.
' I'm intentionaly using very simple algorithms and data structures. Readability outweighs performance.
'
' Gary Luckenbaugh - fall of '22

' Parse buffer
Dim Shared buff$
buff$ = ""

' Used to switch between expression parse
' and RPN evaluation.  This was a bad design choice
Dim Shared buff_after_parse_or$

' Program Counter (aka current line being executed)
Dim Shared pc%

' Operators.  I really should have just created a single
' array.  There should also be a l_ops const
Const ops = 10
Dim Shared op$(ops)
Dim Shared logical_op$(9)

Data "dummy","+","-","*","/","(",")","=","^","@",","
Data "dummy","<","<=","=",">",">=","<>","AND","OR"

Dim t

For t = 0 To ops
    Read op$(t)
Next t

For t = 0 To 8
    Read logical_op$(t)
Next t

' Editor buffer
Dim Shared statement$(5000)

' @ array
Dim Shared at_array%(5000)

' helpful pre declaraion of common FOR-NEXT variable
Dim Shared i%
Dim Shared vars$(26)

' Retained information about FOR command
Dim Shared acvar$(100)
Dim Shared aistart%(100)
Dim Shared ailast%(100)
Dim Shared aistep%(100)
Dim Shared ailine%(100)
Dim Shared for_stackp%

' GOSUB stack
Dim Shared gosub_stack%(100)
Dim Shared gosub_stack_ptr%

' RPN evaluation stack
Dim Shared mystack$(50)
Dim Shared msp%


Call Main

Function myucase$ (s$)

    Dim c$(100)
    Dim flag%
    Dim collect$

    flag% = 1
    collect$ = ""

    For i% = 1 To Len(s$)
        c$(i%) = Mid$(s$, i%, 1)
        If Asc(c$(i%)) = 34 Then
            If flag% = 1 Then
                flag% = 0
            Else
                flag% = 1
            End If
        End If
        If flag% = 1 Then
            c$(i%) = UCase$(c$(i%))
        End If
    Next i%

    For i% = 1 To Len(s$)
        collect$ = collect$ + c$(i%)
    Next i%

    myucase$ = collect$

End Function
  
Sub Main

    Dim k%

    For k% = 1 To 5000
        at_array%(k%) = 2 * k% ' test pattern, should change to zeros
    Next k%

    For k% = 1 To 5000
        statement$(k%) = ""
    Next k%

    Dim in$

    Dim cont%

    Dim tok$

    cont% = 1

    Do While cont% <> 0

        Line Input "basic> ", in$

        in$ = myucase$(in$)

        buff$ = in$

        tok$ = gettok$(0)

        If Left$(tok$, 1) >= "0" And Left$(tok$, 1) <= "9" Then

            process_line (in$)

        ElseIf tok$ <> "BYE" Then

            do_command (in$)

        ElseIf tok$ = "BYE" Then

            cont% = 0

        End If

    Loop

End Sub

Sub edit_delete (linenum%)
    statement$(linenum%) = ""
End Sub

Sub edit_replace (linenum%, buf$)
    statement$(linenum%) = buf$
End Sub

Sub process_line (lin$)

    Dim line_tok$
    Dim nexttok$
    Dim mystring$
    Dim linenum%

    buff$ = lin$

    line_tok$ = gettok$(0)
    nexttok$ = gettok$(0)

    If nexttok$ = "~" Then
        Call edit_delete(Val(line_tok$))
    Else
        linenum% = Val(line_tok$)
        mystring$ = nexttok$ + buff$

        Call edit_replace(linenum%, mystring$)

    End If

End Sub

Sub do_command (cmd$)

    If cmd$ = "LIST" Then

        do_list

    ElseIf Left$(cmd$, 5) = "PRINT" Then

        do_print (cmd$)

    ElseIf Left$(cmd$, 3) = "LET" Then

        do_let (cmd$)
        buff$ = buff_after_parse_or$

        If gettok$(0) <> "~" Then
            Print "warning: extra stuff after expression"
        End If

    ElseIf cmd$ = "NEW" Then

        do_new


    ElseIf Left$(cmd$, 4) = "LOAD" Then

        do_load (LTrim$(RTrim$(Mid$(cmd$, 5)))) ' The trims may not be necessary

    ElseIf Left$(cmd$, 4) = "SAVE" Then

        do_save (LTrim$(RTrim$(Mid$(cmd$, 5)))) '

    ElseIf cmd$ = "RUN" Then

        do_run

    Else

        do_let ("LET " + cmd$)
        buff$ = buff_after_parse_or$
        If gettok$(0) <> "~" Then
            Print "warning: extra stuff after expression"
        End If

    End If

End Sub

Sub do_load (fname$)

    Dim myline$

    do_new

    Open fname$ For Input As #1
    Do While Not EOF(1)
        Line Input #1, myline$
        process_line (myline$)
    Loop
    Close #1

End Sub

Sub do_save (fname$)

    Open fname$ For Output As #2

    For i% = 1 To 5000
        If statement$(i%) <> "" Then
            Print #2, Str$(i%) + " " + statement$(i%)
        End If
    Next i%

    Close #2
 
End Sub
    

Sub do_new

    For i% = 1 To 5000
        statement$(i%) = ""
    Next i%

End Sub

Sub do_list

    For i% = 1 To 5000

        If statement$(i%) <> "" Then
            Print Str$(i%) + " " + statement$(i%)
        End If

    Next i%

End Sub


Sub do_run

    Dim startpc%
    Dim stoprun%
    Dim cmd$
    Dim target%

    startpc% = 0
    for_stackp% = 0

    gosub_stack_ptr% = 0

    stoprun% = 0

    For l% = 1 To 5000

        If statement$(l%) <> "" Then
            startpc% = l%
            Exit For
        End If

    Next l%

    If startpc% = 0 Then
        Print "No statements in buffer!"
        Exit Sub
    End If

    pc% = startpc%

    Do While stoprun% = 0 And pc% <> 0

        cmd$ = statement$(pc%)

        If Left$(cmd$, 3) = "LET" Then
            do_let (cmd$)
            buff$ = buff_after_parse_or$
            If gettok$(0) <> "~" Then
                Print "warning: extra stuff after LET expression on line " + Str$(pc%)
            End If

        ElseIf Left$(cmd$, 5) = "PRINT" Then
            do_print (cmd$)

        ElseIf Left$(cmd$, 4) = "GOTO" Then
            pc% = do_goto(cmd$) - 1

        ElseIf Left$(cmd$, 5) = "GOSUB" Then
            pc% = do_gosub(cmd$) - 1

        ElseIf Left$(cmd$, 7) = "RETURN" Then
            pc% = do_return(0)

        ElseIf cmd$ = "END" Or cmd$ = "STOP" Then
            stoprun% = 1

        ElseIf Left$(cmd$, 2) = "IF" Then
            target% = do_if(cmd$)

            If target% <> 0 Then
                pc% = target% - 1
            End If

        ElseIf Left$(cmd$, 5) = "INPUT" Then
            do_input (cmd$)

        ElseIf Left$(cmd$, 3) = "FOR" Then
            do_for (cmd$)

        ElseIf Left$(cmd$, 4) = "NEXT" Then
            do_next (cmd$)

        Else
            cmd$ = "LET " + cmd$
            do_let (cmd$)
        End If

        If Left$(cmd$, 4) <> "NEXT" Then
            pc% = next_pc%(pc%)
        End If

    Loop

End Sub

Sub do_print (cmd$)

    Dim tok$
    Dim comma_last%

    comma_last% = 0

    buff$ = Mid$(cmd$, 6)

    tok$ = gettok$(0)

    If tok$ = "~" Then
        Print ""
        Exit Sub
    End If

    Do While tok$ <> "~"

        comma_last% = 0

        If Left$(tok$, 1) = Chr$(34) Then
            tok$ = Mid$(tok$, 2, Len(tok$) - 2)
            Print tok$ + " ";

            tok$ = gettok$(0)

            If tok$ = "," Then
                tok$ = gettok$(0)
                comma_last% = 1
            End If

        Else
            Print Expr_eval$(tok$ + buff$) + " ";
            buff$ = buff_after_parse_or$

            tok$ = gettok$(0)

            If tok$ = "," Then
                tok$ = gettok$(0)
                comma_last% = 1
            End If
        End If
    Loop

    If comma_last% <> 1 Then
        Print ""
    End If

End Sub

Function do_if% (cmd$)
    Dim b%
    Dim linenum%

    b% = Val(Expr_eval$(Mid$(cmd$, 3)))

    buff$ = buff_after_parse_or$
    If gettok$(0) <> "THEN" Then
        Print "warning: THEN not found right after expression end on line " + Str$(pc%)
    End If

    i% = InStr(1, cmd$, "THEN")

    If i% <> 0 Then
        linenum% = Val(Mid$(cmd$, i% + 4))
    Else
        Print "malformed command is  :" + cmd$ + ":   missing THEN"
        linenum% = 0
    End If

    If b% Then
        do_if% = linenum%
    Else
        do_if% = 0
    End If

End Function



Sub do_for (cmd$)

    cmd$ = Mid$(cmd$, 4)
    buff$ = cmd$

    Dim cvar$

    cvar$ = gettok$(0)

    If gettok$(0) <> "=" Then
        Print "Missing = in FOR statment on line " + Str$(pc%)
        Exit Sub
    End If

    Dim istart%
    Dim ilast%
    Dim istep%

    istart% = Val(Expr_eval$(buff$))

    buff$ = buff_after_parse_or$

    If gettok$(0) <> "TO" Then
        Print "Missing TO in FOR on line " + Str$(pc%)
        Exit Sub
    End If

    ilast% = Val(Expr_eval$(buff$))

    buff$ = buff_after_parse_or$

    istep% = 1

    If gettok$(0) = "STEP" Then

        istep% = Val(Expr_eval$(buff$))

    End If

    buff$ = buff_after_parse_or$

    If gettok$(0) <> "~" Then

        Print "Extraneous stuff after STEP value on line " + Str$(pc%)

    End If

    for_stackp% = for_stackp% + 1
    acvar$(for_stackp%) = cvar$
    aistart%(for_stackp%) = istart%
    ailast%(for_stackp%) = ilast%
    aistep%(for_stackp%) = istep%
    ailine%(for_stackp%) = pc%

    Call Set_variable(cvar$, Str$(istart%)) ' Why is Call required here?

End Sub
  
Sub do_next (cmd$)

    If for_stackp% <= 0 Then
        Print "stack underflow NEXT without FOR on line " + Str$(pc%)
        pc% = next_pc(pc%)
        Exit Sub
    End If


    buff$ = Mid$(cmd$, 5)

 

    Dim controlvariable$

    controlvariable$ = gettok$(0)

    If controlvariable$ <> acvar$(for_stackp%) Then
        Print "mismatched next on line " + Str$(pc%)
    End If


    Dim cvar$
    Dim start%
    Dim last%
    Dim istep%
    Dim iline%



    start% = aistart%(for_stackp%)
    last% = ailast%(for_stackp%)
    istep% = aistep%(for_stackp%)
    iline% = ailine%(for_stackp%)


    cvar$ = acvar$(for_stackp%)
    Dim v%
    v% = Val(Get_variable(cvar$))

    v% = v% + istep%

    If istep% >= 0 Then
        If v% > last% Then
            pc% = next_pc%(pc%)

            for_stackp% = for_stackp% - 1
        Else
            pc% = next_pc%(iline%)
        End If
    Else
        If v% < last% Then
            pc% = next_pc%(pc%)

            for_stack% = for_stackp% - 1
        Else
            pc% = next_pc%(iline%)
        End If
    End If

    Call Set_variable(cvar$, Str$(v%))

End Sub

Function next_pc% (pc%)

    For i% = pc% + 1 To 5000

        If statement$(i%) <> "" Then
            next_pc% = i%
            Exit Function
        End If

    Next i%

    next_pc% = 0

End Function

Function do_goto% (cmd$)
    do_goto% = Val(Mid$(cmd$, 5))
End Function

Function do_gosub% (cmd$)
    do_gosub% = Val(Mid$(cmd$, 6))
    gosub_stack_ptr% = gosub_stack_ptr% + 1
    gosub_stack%(gosub_stack_ptr%) = pc%
End Function

Function do_return% (dummy)
    do_return% = gosub_stack%(gosub_stack_ptr%)
    gosub_stack_ptr% = gosub_stack_ptr% - 1
End Function

Sub do_let (cmd$)

    Dim cvar$
    Dim firstchar$

    buff$ = Mid$(cmd$, 4)
    cvar$ = gettok$(0)

    If cvar$ = "@" Then
        do_let_at_array (buff$)
        Exit Sub
    End If

    firstchar$ = Left$(cvar$, 1)
    If Not (firstchar$ >= "A" And firstchar$ <= "Z") Then

        Print "Variable name must follow LET"
        Exit Sub

    ElseIf gettok$(0) <> "=" Then

        Print "Missing '=' in LET  " + cmd$
        Exit Sub

    End If

    Call Set_variable(cvar$, Expr_eval$(buff$))

End Sub

Sub do_let_at_array (e$)

    Dim leftval%
    Dim rightval%

    Dim rpn$

    buff$ = e$
    If gettok$(0) <> "(" Then
        Print "an open left paren ( must follow @ in LET"
        Exit Sub
    End If

    rpn$ = parse_or$(0)

    If gettok$(0) <> ")" Then
        Print "missing right paren ) in LET @"
        Exit Sub
    End If

    buff_after_parse_or$ = buff$

    leftval% = Val(Eval_rpn$(rpn$))

    buff$ = buff_after_parse_or$

    If gettok$(0) <> "=" Then
        Print "missing = in LET @"
        Exit Sub
    End If

    rightval% = Val(Eval_rpn$(parse_or$(0)))

    If Not (leftval% >= 1 And leftval% <= 5000) Then
        Print "subscript out of bounds on left side of assignment is " + Str$(leftval%)
        Exit Sub
    End If

    at_array%(leftval%) = rightval%

End Sub

Sub do_input (cmd$)

    Dim cvar$
    Dim firstchar$

    buff$ = Mid$(cmd$, 6)
    cvar$ = gettok$(0)

    If cvar$ = "@" Then
        do_input_at_array (buff$)
        Exit Sub
    End If

    firstchar$ = Left$(cvar$, 1)
    If Not (firstchar$ >= "A" And firstchar$ <= "Z") Then
        Print "Variable name must follow INPUT"
        Exit Sub
    End If

    Dim user_input$

    Input user_input$

    user_input$ = LTrim$(RTrim$(user_input$))

    Call Set_variable(cvar$, Expr_eval$(user_input$))

End Sub

Sub do_input_at_array (e$)

    Dim user_input$

    Dim subval%
    Dim inpval%

    Dim rpn$

    buff$ = e$
    If gettok$(0) <> "(" Then
        Print "an open left paren ( must follow @ in INPUT"
        Exit Sub
    End If

    rpn$ = parse_or$(0)

    If gettok$(0) <> ")" Then
        Print "missing right paren ) in INPUT @"
        Exit Sub
    End If

    subval% = Val(Eval_rpn$(rpn$))

    Line Input user_input$

    inpval = Val(Expr_eval$(user_input$))

    If Not (subval% >= 1 And subval% <= 5000) Then
        Print "subscript out of bounds on left side of INPUT @ is "; subval%
        Exit Sub
    End If

    at_array%(subval%) = inpval

End Sub

Function Eval_rpn$ (rpn$)

    buff$ = rpn$
    Dim tok$

    clearstack
    tok$ = gettok$(0)

    If tok$ = "~" Then
        Eval_rpn = "0"
        Exit Function
    End If

    Do While tok$ <> "~"
        process_tok (tok$)
        tok$ = gettok$(0)
    Loop

    Dim r$
    r$ = mypop(0)

    Eval_rpn$ = r$

End Function
  
Function Expr_eval$ (e$)

    clearstack

    buff$ = e$
    buff$ = parse_or$(0)

    Dim tok$

    tok$ = gettok$(0)

    If tok$ = "~" Then
        Expr_eval = "0"
        Exit Function
    End If

    Do While tok$ <> "~"
        process_tok (tok$)
        tok$ = gettok$(0)
    Loop

    Dim r%
    r% = Val(mypop(0))

    Expr_eval$ = Str$(r%)

End Function

Sub process_tok (tok$)

    If tok$ = "UNM" Then
        mypush (Str$(-Val(mypop$(0))))
        Exit Sub
    End If

    If tok$ = "GETAT" Then

        i% = Val(mypop$(0))
        If i% < 1 Or i% > 5000 Then
            Print "Subscript out of range 1..5000 for " + Str$(i%)
            mypush (Str$(0))
            Exit Sub
        End If

        mypush (Str$(at_array%(i%)))

        Exit Sub
    End If

    For i% = 1 To ops
        If op$(i%) = tok$ Then
            process_op (op$(i%))
            Exit Sub
        End If
    Next i%
 
    For i% = 1 To 8
        If logical_op$(i%) = tok$ Then
            process_op (logical_op$(i%))
            Exit Sub
        End If
    Next i%
 
    process_num_or_var (tok$)

End Sub

Sub process_op (op$)
    Dim leftval%
    Dim rightval%
    Dim res%

    rightval% = Val(mypop$(0))
    leftval% = Val(mypop$(0))

    If op$ = "+" Then
        res% = leftval% + rightval%
    ElseIf op$ = "-" Then
        res% = leftval% - rightval
    ElseIf op$ = "*" Then
        res% = leftval% * rightval%
    ElseIf op$ = "/" Then
        res% = leftval% / rightval%
    ElseIf op$ = "^" Then
        res% = leftval% ^ rightval%
    ElseIf op$ = "<" Then
        res% = leftval% < rightval%
    ElseIf op$ = "<=" Then
        res% = leftval% <= rightval%
    ElseIf op$ = "=" Then
        res% = leftval% = rightval%
    ElseIf op$ = ">" Then
        res% = leftval% > rightval%
    ElseIf op$ = ">=" Then
        res% = leftval% >= rightval%
    ElseIf op$ = "AND" Then
        res% = leftval% And rightval%
    ElseIf op$ = "OR" Then
        res% = leftval% Or rightval%
    Else
        Print "Not an operator " + op$
        Exit Sub
    End If

    mypush (Str$(res%))

End Sub

Sub process_num_or_var (num$)
    If Left$(num$, 1) >= "0" And Left$(num$, 1) <= "9" Then

        mypush (num$)
    Else
        mypush (Get_variable(num$))
    End If
End Sub

Sub clearstack ()
    msp% = 1
End Sub

Sub mypush (tok$)
    mystack$(msp%) = tok$
    msp% = msp% + 1
End Sub

Function mypop$ (dummy)
    msp% = msp% - 1
    mypop$ = mystack$(msp%)
End Function

'This may not be needed - NOP now

Function rid_eol$ (inpval$)
    rid_eol$ = inpval$
End Function

Sub Set_variable (key$, item$)

    If Len(key$) <> 1 Then
        Print "variable name " + key$ + " is too long"
        Exit Sub
    End If

    Dim nkey%

    nkey% = Asc(key$) - Asc("A") + 1

    vars$(nkey%) = item$

End Sub

Function Get_variable$ (key$)

    If Len(key$) <> 1 Then
        'print "variable name " + key + " is too long"
        Get_variable$ = "dummy"
        Exit Function
    End If

    Dim nkey%

    nkey% = Asc(key$) - Asc("A") + 1

    Get_variable$ = vars$(nkey%)

End Function

' here we go, the gettok$ function is the so-called lexical scanner.

Function gettok$ (dummy)
    Dim collect$
    Dim b$
    Dim b1$
    Dim i%

    Do While Left$(buff$, 1) = " "
        buff$ = Mid$(buff$, 2)
    Loop

    If buff$ = "" Then
        gettok$ = "~"
        Exit Function
    End If

    collect$ = ""

    b$ = Left$(buff$, 1)
    b1$ = Mid$(buff$, 2, 1)

    If b$ = "<" And b1$ = "=" Then
        buff$ = Mid$(buff$, 3)
        gettok$ = b$ + b1$
        Exit Function
    End If

    If b$ = "<" And b1$ = ">" Then
        buff$ = Mid$(buff$, 3)
        gettok$ = b$ + b1$
        Exit Function
    End If

    If b$ = "<" Then
        buff$ = Mid$(buff$, 2)
        gettok$ = b$
        Exit Function
    End If

    If b$ = ">" And b1$ = "=" Then
        buff$ = Mid$(buff$, 3)
        gettok$ = b$ + b1$
        Exit Function
    End If

    If b$ = ">" Then
        buff$ = Mid$(buff$, 2)
        gettok$ = b$
        Exit Function
    End If

    For i% = 1 To ops
        If b$ = op$(i%) Then
            gettok$ = b$
            buff$ = Mid$(buff$, 2)
            Exit Function
        End If
    Next i%

    If b$ >= "A" And b$ <= "Z" Then
        Do While b$ >= "A" And b$ <= "Z" Or b$ >= "0" And b$ <= "9"
            collect$ = collect$ + b$
            buff$ = Mid$(buff$, 2)
            b$ = Left$(buff$, 1)
        Loop
        gettok$ = collect$
        Exit Function
    End If

    If b$ >= "0" And b$ <= "9" Then
        Do While b$ >= "0" And b$ <= "9"
            collect$ = collect$ + b$
            buff$ = Mid$(buff$, 2)
            b$ = Left$(buff$, 1)
        Loop

        gettok$ = collect$
        Exit Function

    End If

    If Asc(b$) = 34 Then ' quote mark
        collect$ = b$
        buff$ = Mid$(buff$, 2)
        b$ = Left$(buff$, 1)
        Do While Asc(b$) <> 34
            collect$ = collect$ + b$
            buff$ = Mid$(buff$, 2)
            b$ = Left$(buff$, 1)
        Loop

        gettok$ = collect$ + Chr$(34)
        buff$ = Mid$(buff$, 2)

        Exit Function
    End If

    gettok$ = "~": Rem end of file token

End Function

' The remainder of the program constitutes a recursive descent parser.  Each function can be
' viewed as a production in a context free grammar.  A CFG is a grammar where you can predict what
' is coming based on looking at the next symbol.  High precedence operators appear first in the function
' list.  as you read down the code, you'll see lower and lower precedence operators.

' A factor is a high precedence bundle of tokens. An identier, a number, a parenthesized expression,
' or a unary minus followed by an expression. you will see tok, and rpn in all the following functions.
' each function returns the rpn of the converted expression fragment.


Function parse_factor$ (dummy)
    Dim tok$
    Dim rpn$
    '    Console.WriteLine("Enter parse_factor with buff$ = '" + buff$ +"'")
    tok$ = gettok$(0)

    If tok$ = "(" Then

        rpn$ = parse_or$(0)
        tok$ = gettok$(0)

        If tok$ <> ")" Then
            Print "Missing ) before " + buff$
            buff$ = ""
            parse_factor$ = "~"
            Exit Function
        End If
        parse_factor$ = rpn$
        Exit Function
    End If

    If tok$ = "@" Then

        tok$ = gettok$(0)

        If tok$ <> "(" Then
            Print "Error in let atsign"
            buff$ = ""
            parse_factor = "~"
            Exit Function
        End If

        rpn$ = parse_or$(0)
        
        tok$ = gettok$(0)
        
        If tok$ <> ")" Then
            Print "Missing ) before " + buff$
            buff$ = ""
            parse_factor = "~"
            Exit Function
        End If
        ' Console.WriteLine("@ evaluated to "+rpn+"GETAT")
        parse_factor$ = rpn$ + "GETAT "
        Exit Function



    End If

    If tok$ = "-" Then ' I think this is uneccesary after review
        rpn$ = parse_or$(0)
        parse_factor$ = rpn$ + "UNM "
        Exit Function
    End If

    '    Console.WriteLine("token added to rpn is " + tok$ + " ")
    rpn$ = tok$ + " "
    parse_factor$ = rpn$
    '    Console.WriteLine("Exit parse_factor with buff$ = '" + buff$ +"'")
End Function

Function parse_exponent$ (dummy)
    Dim tok$
    Dim rpn$

    rpn$ = parse_factor$(0)
    tok$ = gettok$(0)

    Do While tok$ = "^"
        rpn$ = rpn$ + parse_exponent$(0) + tok$
        tok$ = gettok$(0)
    Loop

    buff$ = tok$ + " " + buff$
    parse_exponent$ = rpn$
End Function

' parse a unary minus which has less precedence than exponentiation
' that's why unary - is checked here before exponential expressions above
'
' in the rpn we use a separate designation for unary minus (UNM)
' this makes it easy for the vm to understand the different
' interpretations


Function parse_unary_minus$ (dummy)
    Dim tok$
    Dim rpn$
    Dim neg%

    neg% = 0

    tok$ = gettok$(0)
    Do While tok$ = "-" Or tok$ = "+"
        If tok$ = "-" Then
            If neg% = 0 Then
                neg% = 1
            ElseIf neg% = 1 Then
                neg% = 0
            End If
        End If
        tok$ = gettok$(0)
    Loop

    buff$ = tok$ + " " + buff$

    rpn$ = parse_exponent$(0)

    If neg% = 1 Then
        rpn$ = rpn$ + "UNM "
    End If

    parse_unary_minus$ = rpn$

End Function

Function parse_term$ (dummy)
    Dim tok$
    Dim rpn$
    
    rpn$ = parse_unary_minus(0)
    tok$ = gettok$(0)
    
    Do While tok$ = "*" Or tok$ = "/"
        rpn$ = rpn$ + parse_unary_minus(0) + tok$
        tok$ = gettok$(0)
    Loop
    
    buff$ = tok$ + " " + buff$
    parse_term$ = rpn$
    
End Function



Function parse_expr$ (dummy)
    Dim tok$
    Dim rpn$
    
    rpn$ = parse_term(0)
    tok$ = gettok$(0)
    
    Do While tok$ = "+" Or tok$ = "-"
        rpn$ = rpn$ + parse_term(0) + tok$
        tok$ = gettok$(0)
    Loop
    
    buff$ = tok$ + " " + buff$
    
    parse_expr$ = rpn$
    
End Function

Function parse_condition$ (dummy)
    Dim tok$
    Dim rpn$
    
    rpn$ = parse_expr$(0)
    tok$ = gettok$(0)

    If Left$(tok$, 1) = "=" Or Left$(tok$, 1) = ">" Or Left$(tok$, 1) = "<" Then
        rpn$ = rpn$ + parse_expr(0) + tok$
        tok$ = gettok$(0)
    End If
    
    buff$ = tok$ + " " + buff$
    parse_condition$ = rpn$

End Function

Function parse_and$ (dummy)
    Dim tok$
    Dim rpn$

    rpn$ = parse_condition(0)
    tok$ = gettok$(0)

    Do While tok$ = "AND"
        rpn$ = rpn$ + parse_condition(0) + tok$ + " "
        tok$ = gettok$(0)
    Loop

    buff$ = tok$ + " " + buff$
    parse_and$ = rpn$

End Function

Function parse_or$ (dummy)
    Dim tok$
    Dim rpn$
    '    Console.WriteLine("Entering parse_or with buff$ = " + buff$)
    rpn$ = parse_and(0)

    tok$ = gettok$(0)

    Do While tok$ = "OR"
        rpn$ = rpn$ + parse_and(0) + tok$ + " "
        tok$ = gettok$(0)
    Loop

    buff$ = tok$ + " " + buff$
    parse_or$ = rpn$

    buff_after_parse_or$ = buff$

End Function
