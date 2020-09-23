Attribute VB_Name = "modMain"
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "USER32" () As Long


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type Special
    Name As String
    Value As String
End Type

Public Type Starter
    Value As String
    bJoining As Boolean
    strJoints As String
End Type

Public Const BrainFile As String = "\AIG2BRN.data"
Public Const ConversionFile As String = "\AIG2CON.data"
Public Const SpecialFile As String = "\AIG2SPC.data"
Public Const LogFile As String = "\AIG2LOG.data"
Public Const Person1File As String = "\AIG2PN1.data"
Public Const Person2File As String = "\AIG2PN2.data"
Public Const PersonfFile As String = "\AIG2PNF.data"
Public Const JointsFile As String = "\AIG2JNT.data"
Public Const StartersFile As String = "\AIG2STR.data"
Public Const CaseFile As String = "\AIG2CAS.data"

Private mRegion As Long
Public User_Name As String
Public Stars() As String
Public Joining As Boolean
Public Continue As Boolean
Public Joints() As String
Public Specials() As Special

Public Sub Initialize()
    Call LoadSpecials
    For i = 0 To UBound(Specials)
        If Specials(i).Name = "VERSION" Then
            Specials(i).Value = Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
            Exit For
        End If
    Next i
    Call SaveSpecials
    Call WriteLog("LOGGED IN")
    Joining = False
    Continue = True
    frmMain.txtOut = "Hi " + AnswerSpecial("SPC:USER_NAME:HF") + ", welcome to AIG2. Whenever you like, you can say something."
    WriteLog ("STARTER |" + frmMain.txtOut + "|")
    frmMain.tmrStart.Interval = (Int((11) * Rnd) + 10) * 1000
End Sub

Public Sub Terminate()
    Call SaveSpecials
    Call WriteLog("LOGGED OUT")
    Call WriteLog("")
End Sub

Public Function RequestAnswer(strIn As String, readJoints As Boolean) As String
    On Error GoTo exitFunction
    
    frmMain.tmrStart.Enabled = False
    
    Dim Data As String
    Dim DFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim Part3 As String
    Dim strInOld As String
    Dim Rating As Integer
    Dim BestAnswer As String
    Dim BestRating As String
    Dim BestJoints As String
    Dim BestStars() As String
    Dim Continue As Boolean
    Dim Found As Boolean
    
    WriteLog ("QUESTION |" + strIn + "|")
    
    strIn = Conversions(strIn)
    strIn = CorrectCase(strIn)
    strIn = RemoveInvalid(strIn)
    
    strInOld = strIn
    
    strIn = UCase(Trim(strIn))
    DFile = FreeFile
    BestAnswer = ""
    BestRating = 0
    BestJoints = ""
    Found = False
    ReDim Preserve Stars(0 To 0) As String
    
    If readJoints = False Then
        Open App.Path + BrainFile For Input As DFile
    Else
        Open App.Path + JointsFile For Input As DFile
    End If
    
        Do Until EOF(DFile)
            Part1 = ""
            Part2 = ""
            
            Line Input #DFile, Data
            
            If Data <> "" Then
                If readJoints = True Then
                    For i = 0 To UBound(Joints)
                        If Split(UCase(Trim(Split(Data, ">")(0))), "|")(0) = Joints(i) Then Continue = True: Exit For
                    Next i
                    If Continue = False Then GoTo JointsLoop
                End If
                
                Part1 = UCase(Trim(Split(Data, ">")(0)))
                Part2 = Trim(Split(Data, ">")(1))
                
                If UBound(Split(Data, ">")) = 2 Then Part3 = Trim(Split(Data, ">")(2))
                
                Rating = Match(strIn, Part1, strInOld)
                
                If Rating = -1 Then
                    BestAnswer = Part2
                    If UBound(Split(Data, ">")) = 2 Then BestJoints = Part3
                    BestRating = -1
                    BestStars = Stars
                    Found = True
                    Exit Do
                ElseIf Rating > BestRating Then
                    BestAnswer = Part2
                    If UBound(Split(Data, ">")) = 2 Then BestJoints = Part3
                    BestRating = Rating
                    BestStars = Stars
                    Found = True
                End If
JointsLoop:
            End If
        Loop
        
        Stars = BestStars
        RequestAnswer = Answer(BestAnswer)
        
        If Found = False Then
            RequestAnswer = "I'm sorry, I don't know what to say - no match."
            If Joining = True And readJoints = True Then
                For i = 0 To UBound(Joints)
                    If Joints(i) = "*" Or Joints(i) = "*:CONT" Then RequestAnswer = RequestAnswer(strIn, False): Exit For
                Next i
            End If
        End If
        
        If BestJoints <> "" Then
            Joining = True
            Continue = False
            UpdateJoints (BestJoints)
        Else
            Joining = False
            Continue = True
        End If
        
        WriteLog ("ANSWSER |" + RequestAnswer + "|")
        
    Close DFile
    
    If Joining = True Then
        frmMain.tmrStart.Enabled = False
        For i = 0 To UBound(Joints)
            If Joints(i) = "*:CONT" Then frmMain.tmrStart.Enabled = True: Continue = True: Exit For
        Next i
    Else
        frmMain.tmrStart.Enabled = True
    End If

exitFunction:
    Exit Function
    
End Function

Public Function Starter() As String
    Dim Starters() As Starter
    Dim Data As String
    Dim SFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim StrCount As Integer
    Dim Chosen As Integer
    
    SFile = FreeFile
    StrCount = 0
    ReDim Preserve Starters(0) As Starter
    Starters(0).Value = "No valid starters."
    Starters(0).bJoining = False
    
        Open App.Path + StartersFile For Input As SFile
    
        Do Until EOF(SFile)
            Part1 = ""
            Part2 = ""
            
            Line Input #SFile, Data
            
            If Data <> "" Then
                
                StrCount = StrCount + 1
                ReDim Preserve Starters(0 To StrCount) As Starter
                
                Part1 = Trim(Split(Data, ">")(0))
                Part2 = UCase(Trim(Split(Data, ">")(1)))
                
                Starters(StrCount).Value = Part1
                
                If Part2 <> "||" Then
                    Starters(StrCount).bJoining = True
                    Starters(StrCount).strJoints = Part2
                Else
                    Starters(StrCount).bJoining = False
                End If
                
            End If
        Loop
        
        If UBound(Starters) = 0 Then
            Starter = Starters(0).Value
            Joining = Starters(0).bJoining
            If Joining Then
                Continue = False
            Else
                Continue = True
            End If
            If Joining Then UpdateJoints (Starters(0).strJoints)
        Else
            Randomize
            Chosen = Int((UBound(Starters) * Rnd) + 1)
            Starter = Answer(Starters(Chosen).Value)
            If Starters(Chosen).bJoining = True Then
                Joining = True
                Continue = False
                UpdateJoints (Starters(Chosen).strJoints)
            Else
                Joining = False
                Continue = True
            End If
        End If
        
        WriteLog ("STARTER |" + Starter + "|")
        
    Close SFile
    
    If Joining = True Then
        frmMain.tmrStart.Enabled = False
        For i = 0 To UBound(Joints)
            If Joints(i) = "*:CONT" Then frmMain.tmrStart.Enabled = True: Continue = True: Exit For
        Next i
    Else
        frmMain.tmrStart.Enabled = True
    End If
    
End Function


Public Function Match(strQ As String, strP As String, strInOld) As Integer
    On Error GoTo exitFunction
   
    Dim OLDsplitP() As String
    Dim LastStr As Integer
    Dim splitP() As String
    
    Match = 0
    
    OLDsplitP = Split(strP, "|")
    
    If UBound(OLDsplitP) < 2 Then Exit Function
    
    ReDim splitP(0 To UBound(OLDsplitP) - 1) As String
    
    splitP(0) = ""
    
    Position = 0
    
    For i = 1 To UBound(OLDsplitP) - 1
        If OLDsplitP(i) = "*" Then
            If LastStr = 1 Then
                splitP(Position) = splitP(Position) + OLDsplitP(i)
            Else
                Position = Position + 1
                splitP(Position) = OLDsplitP(i)
            End If
            LastStr = 1
        Else
            If LastStr = 2 Then
                splitP(Position) = splitP(Position) + OLDsplitP(i)
            Else
                Position = Position + 1
                splitP(Position) = OLDsplitP(i)
            End If
            LastStr = 2
        End If
    Next i
    
    ReDim Preserve splitP(0 To Position + 1)
    
    If UBound(splitP) = 2 Then
        If splitP(1) = "*" Then
            Match = 1
            Stars(0) = strInOld
            Exit Function
        ElseIf splitP(1) = strQ Then
            Match = -1
            Exit Function
        Else
            Exit Function
        End If
    End If
    
    If UBound(splitP) = 3 Then
        If splitP(1) = "*" Then
            If InStr(Len(strQ) - Len(splitP(2)), strQ, splitP(2)) = 1 Then
                Match = Len(splitP(2))
                Stars(0) = Mid(strInOld, 1, Len(strInOld) - Len(splitP(2)))
                Stars(1) = Stars(0)
                Exit Function
            Else
                Exit Function
            End If
        ElseIf splitP(2) = "*" Then
            If InStr(strQ, splitP(1)) = 1 Then
                Match = Len(splitP(1))
                Stars(0) = Right(strInOld, Len(strInOld) - Len(splitP(1)))
                Stars(1) = Stars(0)
                Exit Function
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    If UBound(splitP) = 4 Then
        If splitP(1) = "*" And splitP(3) = "*" Then
            If InStr(strQ, splitP(2)) > 0 Then
                Match = Len(splitP(2))
                Stars(2) = Right(strInOld, Len(strInOld) - InStr(strQ, splitP(2)) - Len(splitP(2)) + 1)
                Stars(1) = Left(strInOld, Len(strInOld) - Len(splitP(2)) - Len(Stars(2)))
                Stars(0) = Stars(1) + Stars(2)
                Exit Function
            Else
                Exit Function
            End If
        ElseIf splitP(2) = "*" Then
            If InStr(strQ, splitP(1)) = 1 And InStr(Len(strQ) - Len(splitP(3)), strQ, splitP(3)) = (Len(strQ) - Len(splitP(3))) + 1 Then
                Match = Len(splitP(1)) + Len(splitP(3))
                Stars(0) = Mid(strInOld, Len(splitP(1)) + 1, Len(strInOld) - (Len(splitP(1)) + Len(splitP(3))))
                Stars(1) = Stars(0)
                Exit Function
            Else
                Exit Function
            End If
        End If
    End If
    
    Match = CompareWildCards(strQ, strP, strInOld)

exitFunction:
    Exit Function
    
End Function

Public Function Answer(strIn) As String
    Dim splitin() As String
    
    splitin = Split(strIn, "|")
    LoadSpecials
    
    For i = 1 To UBound(splitin) - 1
        Select Case Left(splitin(i), 1)
        Case "*"
            If Len(splitin(i)) >= 2 Then
                Select Case Left(splitin(i), 2)
                Case "*:"
                    Answer = Answer + AnswerStar(splitin(i))
                Case "*_"
                    Answer = Answer + AnswerStar(splitin(i))
                Case Else
                    Answer = Answer + AnswerStar(splitin(i))
                End Select
            Else
                Answer = Answer + AnswerStar(splitin(i))
            End If
        Case "S"
            Select Case Left(splitin(i), 4)
            Case "SPC:"
                Answer = Answer + AnswerSpecial(splitin(i))
            Case "SET:"
                SetSpecial (splitin(i))
            Case "SET_"
                SetSpecial (splitin(i))
            Case Else
                Answer = Answer + splitin(i)
            End Select
        Case Else
            Answer = Answer + splitin(i)
        End Select
    Next i
    
    SaveSpecials
    
End Function

Private Function AnswerStar(strIn) As String
    Dim splitin() As String
    Dim setH As Boolean
    Dim setHF As Boolean
    Dim setL As Boolean
    Dim setN As Boolean
    Dim setHC As Boolean
    Dim setHFC As Boolean
    Dim setLC As Boolean
    Dim setC As Boolean
    Dim setN2 As Boolean
    Dim setP1 As Boolean
    Dim setP2 As Boolean
    Dim setPF As Boolean
    
    setH = False
    setHF = False
    setL = False
    setHC = False
    setHFC = False
    setLC = False
    setC = False
    setN = False
    setN2 = False
    setP1 = False
    setP2 = False
    setPF = False
    splitin = Split(strIn, ":")
    
    If Mid(strIn, 2, 1) = ":" Then
        AnswerStar = Stars(0)
    Else
        AnswerStar = Stars(CInt(Mid(strIn, 3, 1)))
    End If
    
    If UBound(splitin) = 0 Then Exit Function
    
    Select Case splitin(UBound(splitin))
        Case "H"
            setH = True
        Case "HF"
            setHF = True
        Case "L"
            setL = True
        Case "HC"
            setHC = True
        Case "HFC"
            setHFC = True
        Case "LC"
            setLC = True
        Case "C"
            setC = True
        Case Else
            setN = True
    End Select
    
    If setN Then
        Select Case splitin(UBound(splitin))
            Case "P1"
                setP1 = True
            Case "P2"
                setP2 = True
            Case "PF"
                setPF = True
            Case Else
                setN2 = True
        End Select
    Else
        Select Case splitin(UBound(splitin) - 1)
            Case "P1"
                setP1 = True
            Case "P2"
                setP2 = True
            Case "PF"
                setPF = True
            Case Else
                setN2 = True
        End Select
    End If
    
    If setP1 Then
        AnswerStar = Person1(AnswerStar)
    ElseIf setP2 Then
        AnswerStar = Person2(AnswerStar)
    ElseIf setPF Then
        AnswerStar = Personf(AnswerStar)
    End If
    
    If setH Then
        AnswerStar = UCase(AnswerStar)
    ElseIf setHF Then
        AnswerStar = UppercaseFirst(AnswerStar)
    ElseIf setL Then
        AnswerStar = LCase(AnswerStar)
    ElseIf setHC Then
        AnswerStar = CorrectCase(UCase(AnswerStar))
    ElseIf setHFC Then
        AnswerStar = CorrectCase(UppercaseFirst(AnswerStar))
    ElseIf setLC Then
        AnswerStar = CorrectCase(LCase(AnswerStar))
    ElseIf setC Then
        AnswerStar = CorrectCase(AnswerStar)
    End If

End Function

Public Function AnswerSpecial(strIn) As String
    On Error GoTo exitFunction

    Dim splitin() As String
    Dim setH As Boolean
    Dim setHF As Boolean
    Dim setL As Boolean
    Dim setHC As Boolean
    Dim setHFC As Boolean
    Dim setLC As Boolean
    Dim setC As Boolean
    Dim setN As Boolean
    Dim setN2 As Boolean
    Dim setP1 As Boolean
    Dim setP2 As Boolean
    Dim setPF As Boolean
    
    setH = False
    setHF = False
    setL = False
    setHC = False
    setHFC = False
    setLC = False
    setC = False
    setN = False
    setN2 = False
    setP1 = False
    setP2 = False
    setPF = False
    splitin = Split(strIn, ":")
    
    For i = 0 To UBound(Specials)
        If UCase(splitin(1)) = Specials(i).Name Then
            AnswerSpecial = Specials(i).Value
            Exit For
        End If
    Next i
    
    If UBound(splitin) = 2 Then
        Select Case splitin(UBound(splitin))
            Case "H"
                setH = True
            Case "HF"
                setHF = True
            Case "L"
                setL = True
            Case "HC"
                setHC = True
            Case "HFC"
                setHFC = True
            Case "LC"
                setLC = True
            Case "C"
                setC = True
            Case Else
                setN = True
        End Select
    End If
    
    If setN Then
        Select Case splitin(UBound(splitin))
            Case "P1"
                setP1 = True
            Case "P2"
                setP2 = True
            Case "PF"
                setPF = True
            Case Else
                setN2 = True
        End Select
    Else
        Select Case splitin(UBound(splitin) - 1)
            Case "P1"
                setP1 = True
            Case "P2"
                setP2 = True
            Case "PF"
                setPF = True
            Case Else
                setN2 = True
        End Select
    End If
    
    If setP1 Then
        AnswerSpecial = Person1(AnswerSpecial)
    ElseIf setP2 Then
        AnswerSpecial = Person2(AnswerSpecial)
    ElseIf setPF Then
        AnswerSpecial = Personf(AnswerSpecial)
    End If
    
    If setH Then
        AnswerSpecial = UCase(AnswerSpecial)
    ElseIf setHF Then
        AnswerSpecial = UppercaseFirst(AnswerSpecial)
    ElseIf setL Then
        AnswerSpecial = LCase(AnswerSpecial)
    ElseIf setHC Then
        AnswerSpecial = CorrectCase(UCase(AnswerSpecial))
    ElseIf setHFC Then
        AnswerSpecial = CorrectCase(UppercaseFirst(AnswerSpecial))
    ElseIf setLC Then
        AnswerSpecial = CorrectCase(LCase(AnswerSpecial))
    ElseIf setC Then
        AnswerSpecial = CorrectCase(AnswerSpecial)
    End If
    
exitFunction:
    Exit Function
    
End Function

Private Sub SetSpecial(strIn)
    Dim splitin() As String
    Dim setH As Boolean
    Dim setHF As Boolean
    Dim setL As Boolean
    Dim setHC As Boolean
    Dim setHFC As Boolean
    Dim setLC As Boolean
    Dim setC As Boolean
    Dim setN As Boolean
    Dim setN2 As Boolean
    Dim setP1 As Boolean
    Dim setP2 As Boolean
    Dim setPF As Boolean
    Dim Data As String
    
    setH = False
    setHF = False
    setL = False
    setHC = False
    setHFC = False
    setLC = False
    setC = False
    setN = False
    setN2 = False
    setP1 = False
    setP2 = False
    setPF = False
    splitin = Split(strIn, ":")
    
    If Mid(strIn, 4, 1) = ":" Then
        Data = Stars(0)
    Else
        Data = Stars(CInt(Mid(strIn, 5, 1)))
    End If
    
    If UBound(splitin) = 2 Then
        Select Case splitin(UBound(splitin))
            Case "H"
                setH = True
            Case "HF"
                setHF = True
            Case "L"
                setL = True
            Case "HC"
                setHC = True
            Case "HFC"
                setHFC = True
            Case "LC"
                setLC = True
            Case "C"
                setC = True
            Case Else
                setN = True
        End Select
    End If
    
    If setN Then
        Select Case splitin(UBound(splitin))
            Case "P1"
                setP1 = True
            Case "P2"
                setP2 = True
            Case "PF"
                setPF = True
            Case Else
                setN2 = True
        End Select
    Else
        Select Case splitin(UBound(splitin) - 1)
            Case "P1"
                setP1 = True
            Case "P2"
                setP2 = True
            Case "PF"
                setPF = True
            Case Else
                setN2 = True
        End Select
    End If
    
    If setP1 Then
        Data = Person1(Data)
    ElseIf setP2 Then
        Data = Person2(Data)
    ElseIf setPF Then
        Data = Personf(Data)
    End If
    
    If setH Then
        Data = UCase(Data)
    ElseIf setHF Then
        Data = UppercaseFirst(Data)
    ElseIf setL Then
        Data = LCase(Data)
    ElseIf setHC Then
        Data = CorrectCase(UCase(Data))
    ElseIf setHFC Then
        Data = CorrectCase(UppercaseFirst(Data))
    ElseIf setLC Then
        Data = CorrectCase(LCase(Data))
    ElseIf setC Then
        Data = CorrectCase(Data)
    End If
    
    For i = 0 To UBound(Specials)
        If UCase(splitin(1)) = Specials(i).Name Then
            Specials(i).Value = Data
            Exit For
        End If
    Next i
    
End Sub

Public Function RemoveInvalid(strIn As String) As String
    
    Dim strInOld
    
    RemoveInvalid = strIn
    
    Do
        strInOld = RemoveInvalid
        If Right(RemoveInvalid, 1) = "?" Or Right(RemoveInvalid, 1) = "." Or Right(RemoveInvalid, 1) = "!" Then
            RemoveInvalid = Left(RemoveInvalid, Len(RemoveInvalid) - 1)
        End If
    Loop While RemoveInvalid <> strInOld
    
    RemoveInvalid = Replace(RemoveInvalid, ",", "")
    RemoveInvalid = Replace(RemoveInvalid, ":", "")
    RemoveInvalid = Replace(RemoveInvalid, ";", "")
    
End Function

Public Function Conversions(strIn As String) As String
    Dim splitData() As String
    Dim Data As String
    Dim CFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim splitPart1() As String
    Dim splitPart2() As String
    Dim Doing As Boolean
    Dim strConverted As String
    
    CFile = FreeFile
    
    Conversions = strIn
    
    Open App.Path + ConversionFile For Input As CFile
    
    Do Until EOF(CFile)
    
        Line Input #CFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2(splitPart1(1), Conversions, splitPart2(1))
            
            If Conversions = strConverted Then Doing = False
            
            Conversions = strConverted
            
        Loop
        
    Loop
    
    Conversions = Trim(Conversions)
    
    Close CFile
    
End Function

Public Function UppercaseFirst(strIn As String) As String
    Dim splitin() As String
    
    splitin() = Split(strIn, " ")
    
    For i = 0 To UBound(splitin)
        Select Case Len(splitin(i))
        Case 0
        Case 1
            UppercaseFirst = UppercaseFirst + " " + UCase(Left(splitin(i), 1))
        Case Else
            UppercaseFirst = UppercaseFirst + " " + UCase(Left(splitin(i), 1)) + LCase(Right(splitin(i), Len(splitin(i)) - 1))
        End Select
    Next i
    
    UppercaseFirst = Trim(UppercaseFirst)
    
End Function

Public Function CorrectCase(ByVal strIn As String) As String
    Dim splitData() As String
    Dim Data As String
    Dim CFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim splitPart1() As String
    Dim splitPart2() As String
    Dim Doing As Boolean
    Dim strConverted As String
    
    CFile = FreeFile
    
    CorrectCase = strIn
    
    Open App.Path + CaseFile For Input As CFile
    
    Do Until EOF(CFile)
    
        Line Input #CFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2(splitPart1(1), CorrectCase, splitPart2(1))
            
            If CorrectCase = strConverted Then Doing = False
            
            CorrectCase = strConverted
            
        Loop
        
    Loop
    
    CorrectCase = Trim(CorrectCase)
    
    Close CFile
    
End Function

Public Sub LoadSpecials()
    Dim Data As String
    Dim SFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim Counter As Integer
    
    SFile = FreeFile
    
    Open App.Path + SpecialFile For Input As SFile
        
        Counter = 0
    
        Do Until EOF(SFile)
        
            Part1 = ""
            Part2 = ""
            
            Line Input #SFile, Data
            
            If Data <> "" Then
                Part1 = UCase(Trim(Split(Data, ">")(0)))
                Part2 = Trim(Split(Data, ">")(1))
                
                ReDim Preserve Specials(Counter) As Special
                
                Specials(Counter).Name = Split(Part1, "|")(1)
                Specials(Counter).Value = Split(Part2, "|")(1)
                
            End If
            
            Counter = Counter + 1
        Loop
        
    Close SFile
End Sub

Public Sub SaveSpecials()
    Dim SFile As Integer

    SFile = FreeFile
    
    Open App.Path + SpecialFile For Output As SFile
        For i = 0 To UBound(Specials)
            Print #SFile, "|" + Specials(i).Name + "|>|" + Specials(i).Value + "|"
        Next i
    Close SFile
End Sub

Public Sub WriteLog(strIn As String)
    Dim LFile As Integer
    
    LFile = FreeFile
    
    Open App.Path + LogFile For Append As LFile
        If strIn <> "" Then strIn = strIn + " " + Str(Time) + " " + Str(Date)
        Print #LFile, strIn
    Close LFile
End Sub

Public Function Replace2(ByVal strSearch As String, ByVal strIn As String, ByVal strReplace As String) As String
    Dim CurrentPlace As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim Doing As Boolean
    Dim strReplaced As String
    
    strSearch = " " + strSearch + " "
    strIn = " " + strIn + " "
    strReplace = " " + strReplace + " "
    
    If InStr(UCase(strIn), UCase(strSearch)) = 0 Then Replace2 = Trim(strIn): Exit Function
    
    Doing = True
    
    Do While Doing And InStr(UCase(strIn), UCase(strSearch)) <> 0
        CurrentPlace = InStr(UCase(strIn), UCase(strSearch))
        Part1 = Left(strIn, CurrentPlace - 1)
        Part2 = Right(strIn, Len(strIn) - (Len(Part1) + Len(strSearch)))
        strReplaced = Part1 + strReplace + Part2
        If strIn = strReplaced Then Doing = False
        strIn = strReplaced
    Loop

    Replace2 = Trim(strIn)
    
End Function

Public Function Person1(strIn As String) As String
    Dim splitData() As String
    Dim Data As String
    Dim PFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim splitPart1() As String
    Dim splitPart2() As String
    Dim Doing As Boolean
    Dim strConverted As String
    
    PFile = FreeFile
    
    Open App.Path + Person1File For Input As PFile
    
    Person1 = strIn
    
    Do Until EOF(PFile)
    
        Line Input #PFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2(splitPart1(1), Person1, "|" + splitPart2(1) + "|")
            
            If Person1 = strConverted Then Doing = False
            
            Person1 = strConverted
            
        Loop
        
    Loop
    
    Person1 = Trim(Person1)
    
    Close PFile
    
    PFile = FreeFile
    
    Open App.Path + Person1File For Input As PFile
    
    Do Until EOF(PFile)
    
        Line Input #PFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2("|" + splitPart2(1) + "|", Person1, splitPart2(1))
            
            If Person1 = strConverted Then Doing = False
            
            Person1 = strConverted
            
        Loop
        
    Loop
    
    Person1 = Trim(Person1)
    
    Close PFile
End Function

Public Function Person2(strIn As String) As String
    Dim splitData() As String
    Dim Data As String
    Dim PFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim splitPart1() As String
    Dim splitPart2() As String
    Dim Doing As Boolean
    Dim strConverted As String
    
    PFile = FreeFile
    
    Open App.Path + Person2File For Input As PFile
    
    Person2 = strIn
    
    Do Until EOF(PFile)
    
        Line Input #PFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2(splitPart1(1), Person2, "|" + splitPart2(1) + "|")
            
            If Person2 = strConverted Then Doing = False
            
            Person2 = strConverted
            
        Loop
        
    Loop
    
    Person2 = Trim(Person2)
    
    Close PFile
    
    PFile = FreeFile
    
    Open App.Path + Person2File For Input As PFile
    
    Do Until EOF(PFile)
    
        Line Input #PFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2("|" + splitPart2(1) + "|", Person2, splitPart2(1))
            
            If Person2 = strConverted Then Doing = False
            
            Person2 = strConverted
            
        Loop
        
    Loop
    
    Person2 = Trim(Person2)
    
    Close PFile
End Function

Public Function Personf(strIn As String) As String
    Dim splitData() As String
    Dim Data As String
    Dim PFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim splitPart1() As String
    Dim splitPart2() As String
    Dim Doing As Boolean
    Dim strConverted As String
    
    PFile = FreeFile
    
    Open App.Path + PersonfFile For Input As PFile
    
    Personf = strIn
    
    Do Until EOF(PFile)
    
        Line Input #PFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2(splitPart1(1), Personf, "|" + splitPart2(1) + "|")
            
            If Personf = strConverted Then Doing = False
            
            Personf = strConverted
            
        Loop
        
    Loop
    
    Personf = Trim(Personf)
    
    Close PFile
    
    PFile = FreeFile
    
    Open App.Path + PersonfFile For Input As PFile
    
    Do Until EOF(PFile)
    
        Line Input #PFile, Data
        
        splitData() = Split(Data, ">")
        
        Part1 = splitData(0)
        Part2 = splitData(1)
        
        splitPart1() = Split(Part1, "|")
        splitPart2() = Split(Part2, "|")
        
        Doing = True
        
        Do While Doing
            
            strConverted = Replace2("|" + splitPart2(1) + "|", Personf, splitPart2(1))
            
            If Personf = strConverted Then Doing = False
            
            Personf = strConverted
            
        Loop
        
    Loop
    
    Personf = Trim(Personf)
    
    Close PFile
End Function

Public Sub UpdateJoints(strIn As String)
    Dim split1st() As String
    
    split1st = Split(UCase(strIn), "|")
    
    Joints = Split(split1st(1), ",")
End Sub

Public Function CompareWildCards(strQ As String, strP As String, strInOld) As Integer
    Dim StartStar As Boolean
    Dim EndStar As Boolean
    Dim splitP() As String
    Dim List() As String
    Dim Count As Integer
    Dim LeftOverQ As String
    
    splitP = Split(strP, "|")
    If splitP(1) = "*" Then StartStar = True
    If splitP(UBound(splitP) - 1) = "*" Then EndStar = True
    Count = 0
    
    For i = 1 To UBound(splitP) - 1
        If splitP(i) <> "*" Then
            ReDim Preserve List(0 To Count) As String
            List(Count) = UCase(splitP(i))
            CompareWildCards = CompareWildCards + Len(List(Count))
            Count = Count + 1
        End If
    Next i
    
    ReDim Stars(0 To 0) As String
    
    If StartStar Then
        ReDim Preserve Stars(0 To UBound(Stars) + 1) As String
        If InStr(strQ, List(0)) = 0 Then
            CompareWildCards = 0
            Exit Function
        Else
            Stars(0) = Left(strInOld, InStr(strQ, List(0)) - 1)
            Stars(1) = Left(strInOld, InStr(strQ, List(0)) - 1)
            LeftOverQ = Right(strQ, Len(strQ) - Len(List(0)) - Len(Stars(0)))
        End If
    Else
        If InStr(strQ, List(0)) <> 1 Then
            CompareWildCards = 0
            Exit Function
        Else
            LeftOverQ = Right(strQ, Len(strQ) - Len(List(0)))
        End If
    End If
    
    For i = 1 To (UBound(List) - 1)
        ReDim Preserve Stars(0 To UBound(Stars) + 1) As String
        If InStr(LeftOverQ, List(i)) = 0 Then
            CompareWildCards = 0
            Exit Function
        Else
            Stars(0) = Stars(0) + Left(Right(strInOld, Len(LeftOverQ)), InStr(LeftOverQ, List(i)) - 1)
            Stars(UBound(Stars)) = Left(Right(strInOld, Len(LeftOverQ)), InStr(LeftOverQ, List(i)) - 1)
            LeftOverQ = Right(LeftOverQ, Len(LeftOverQ) - Len(Left(Right(strInOld, Len(LeftOverQ)), InStr(LeftOverQ, List(i)) - 1)) - Len(List(i)))
        End If
    Next i
    
    ReDim Preserve Stars(0 To UBound(Stars) + 1) As String
    
    If EndStar Then
        ReDim Preserve Stars(0 To UBound(Stars) + 1) As String
        If InStr(LeftOverQ, List(UBound(List))) <> 0 Then
            Stars(UBound(Stars) - 1) = Left(Right(strInOld, Len(LeftOverQ)), InStr(LeftOverQ, List(UBound(List))) - 1)
            Stars(UBound(Stars)) = Right(Right(strInOld, Len(LeftOverQ)), Len(LeftOverQ) - Len(List(UBound(List))) - Len(Stars(UBound(List))))
            Stars(0) = Stars(0) + Stars(UBound(Stars) - 1) + Stars(UBound(Stars))
            Exit Function
        Else
            CompareWildCards = 0
            Exit Function
        End If
    Else
        If Right(LeftOverQ, Len(List(UBound(List)))) = List(UBound(List)) Then
            Stars(0) = Stars(0) + Left(Right(strInOld, Len(LeftOverQ)), Len(LeftOverQ) - Len(List(UBound(List))))
            Stars(UBound(Stars)) = Left(Right(strInOld, Len(LeftOverQ)), Len(LeftOverQ) - Len(List(UBound(List))))
            Exit Function
        Else
            CompareWildCards = 0
            Exit Function
        End If
    End If
    
End Function

Public Sub Remove_Title()

Dim new_shape() As POINTAPI
        
ReDim new_shape(0 To 5) As POINTAPI

Dim i

With frmMain
new_shape(i).X = 3
new_shape(i).Y = 22
i = i + 1

new_shape(i).X = .Width / 15 - 3
new_shape(i).Y = 22
i = i + 1

new_shape(i).X = .Width / 15 - 3
new_shape(i).Y = .Height / 15 - 3
i = i + 1

new_shape(i).X = 3
new_shape(i).Y = .Height / 15 - 3
i = i + 1

new_shape(i).X = 3
new_shape(i).Y = 22
i = i + 1

End With

mRegion = CreatePolygonRgn(new_shape(0), UBound(new_shape), 1)
Dim lret As Long
lret = SetWindowRgn(frmMain.hWnd, mRegion, True)

End Sub
