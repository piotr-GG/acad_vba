Attribute VB_Name = "Module1"
Option Explicit
Option Base 0

Const acAscending As Integer = 44
Const acDescending As Integer = 55
Const acBlockRef As String = "INSERT"
Const acMText As String = "MTEXT"
Const acText As String = "TEXT"
Const acAttributeDef As String = "ATTDEF"

Const acFilterByObjectType As Integer = 0
Const acFilterByObjectName As Integer = 2
Const acFilterByLogicalExp As Integer = -4

Const acErrorBlankSpaceClicked As Long = -2147352567
Const acErrorKeywordSelected As Long = -2145320928


Sub AutoNumberPages()
    '*****  DEKLARACJA ZMIENNYCH *****
    'Do przechowywania nazwy bloku, ktory ma zostac ponumerowany
    Dim blockName As String
    Dim selectionSet As AcadSelectionSet
    Dim blockRefObj As AcadBlockReference
    Dim ent As AcadEntity
    
    'Do przechowywania atrybutów
    Dim varAttributes As Variant
    Dim attRef As AcadAttributeReference
    
    'Do przechowywania referencji bloków
    Dim blockRefArray() As AcadBlockReference
    'Tablica dynamiczna
    ReDim blockRefArray(0)
    
    Dim blockCount As Integer, i As Integer, j As Integer, pageCount As Integer
    
    '*****  WYBÓR BLOKU *****
    
    Dim blockRef As AcadBlockReference
    Set blockRef = PickBlockReference
    
    If blockRef Is Nothing Then
        MsgBox "Nie wybrano bloku!", vbCritical
        Exit Sub
    Else
        blockName = blockRef.EffectiveName
    End If
    
    pageCount = 1
    i = 0: j = 0: blockCount = 0
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(1) As Variant, filterData(1) As Variant
    
    'Filtrowanie po typie obiektu - AcadBlockReference
    filterType(0) = acFilterByObjectType: filterData(0) = acBlockRef
    'Filtrowanie po nazwie obiektu - BlockName
    filterType(1) = acFilterByObjectName: filterData(1) = blockName
    'Wybor selection setem z filtrem
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    '*****  DODANIE DO TABLICY *****
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = GetItemsFromSelectionSet(selectionSet)
    
    On Error Resume Next
    Debug.Print selectionSet.count
    If Err.Number = -2147467259 Then
        MsgBox "B³¹d VBA. Zrestartuj AutoCADa"
        Exit Sub
    End If
    On Error GoTo 0
    
    '*****  SORTOWANIE BLOKÓW *****
    
    
    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)
    
    '*****  ZMIANA ATRYBUTU *****
    
    Dim attNum As Long
    attNum = 1
    
    'Zmiana atrybutu numer 2 w bloku ( = ARKUSZ)
    For j = LBound(blockRefArray) To UBound(blockRefArray)
        Set blockRefObj = blockRefArray(j)
        varAttributes = blockRefObj.GetAttributes
        
        Set attRef = varAttributes(attNum)
        attRef.textString = CStr(pageCount)
        pageCount = pageCount + 1
    Next j
    
    'Wyœwietlenie liczby stron do zmiany w bloku
    MsgBox "Zakoñczono numerowanie stron." & vbCrLf & _
           "Liczba stron: " & CStr(pageCount - 1)
End Sub

Sub CreateTableOfContents()

    '*****  DEKLARACJA ZMIENNYCH *****
    Dim blockName As String
    
    Dim selectionSet As AcadSelectionSet
    Dim blockRefObj As AcadBlockReference
    'Do przechowywania atrybutów
    Dim varAttributes As Variant
    'Do przechowywania referencji bloków
    Dim blockRefArray() As AcadBlockReference
    'Do przechowywania opisów
    Dim textArray() As String
    
    'Tablica dynamiczna
    ReDim blockRefArray(0)
    ReDim textArray(0)
    
    'Do ustawiania tekstow na rysunku
    Dim textObj As aCadText
    Dim textString As Variant
    Dim insertionPoint As Variant
    
    'Do iterowania
    Dim i As Integer: i = 0
    Dim txtCount As Integer: txtCount = 0
    
    '*****  WYBÓR BLOKU *****
    Dim pickedBlockRef As AcadBlockReference
    
    Set pickedBlockRef = PickBlockReference

    If Not TypeOf pickedBlockRef Is AcadBlockReference Then
        MsgBox "Z³y wybór!", vbCritical
        Exit Sub
    Else
        blockName = pickedBlockRef.EffectiveName
    End If
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(1) As Variant, filterData(1) As Variant
    
    'Filtrowanie po typie obiektu - AcadBlockReference
    filterType(0) = acFilterByObjectType: filterData(0) = acBlockRef
    'Filtrowanie po nazwie obiektu - BlockName
    filterType(1) = acFilterByObjectName: filterData(1) = blockName
    'Wybor selection setem z filtrem
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    '*****  DODANIE DO TABLICY *****
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = GetItemsFromSelectionSet(selectionSet)

    '*****  SORTOWANIE BLOKÓW *****

    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)
    
    '*****  DODANIE TYTU£ÓW DO TABLICY *****
    
    Dim attNum As Long: attNum = 0
    
    For i = LBound(blockRefArray) To UBound(blockRefArray)
        Set blockRefObj = blockRefArray(i)
        varAttributes = blockRefArray(i).GetAttributes
        
        ReDim Preserve textArray(txtCount)
        textArray(txtCount) = varAttributes(attNum).textString
        txtCount = txtCount + 1
    Next i

    '*****  UMIESZCZENIE TEKSTÓW NA RYSUNKU *****
    insertionPoint = ThisDrawing.Utility.GetPoint(, "Podaj punkt wstawienia: ")

    Dim textHeight As Double
    textHeight = 2.06
    
    Dim increment As Double
    increment = ThisDrawing.Utility.GetDistance(, "Podaj odstêp: ")

    Dim tableOfContentsLayer As AcadLayer
    Set tableOfContentsLayer = GetLayer("Table of contents", True)
    
    For Each textString In textArray
        'Kazdy kolejny tekst jest umieszczony na dole o liczbe rowna wartosci increment
        Set textObj = ThisDrawing.ModelSpace.AddText(VBA.UCase(textString), insertionPoint, textHeight)
        textObj.Layer = tableOfContentsLayer.Name
        insertionPoint(1) = insertionPoint(1) - increment
    Next textString
    
End Sub

Sub ArrangeLayouts()
    '*****  DEKLARACJA ZMIENNYCH *****
    Dim i As Integer, nOfLayouts As Integer
    Dim distanceBetween As Double
    Dim Layout As AcadLayout
    Dim mSpacePoint_1(0 To 2) As Double
    Dim mSpacePoint_2(0 To 2) As Double
    Dim pSpacePoint_1(0 To 2) As Double
    Dim pSpacePoint_2(0 To 2) As Double
    
    On Error GoTo 0
    
    Dim layoutRegenCTL As Integer
    layoutRegenCTL = ThisDrawing.GetVariable("LAYOUTREGENCTL")
    
    MsgBox "LayoutRegenCTL: " & CStr(layoutRegenCTL)
    ThisDrawing.SetVariable "LAYOUTREGENCTL", 0
    
    '*****  PUNKTY NA MODELSPACE *****
    'Lewy dolny punkt
    mSpacePoint_1(0) = 0: mSpacePoint_1(1) = 0
    
    'Prawy gorny punkt
    mSpacePoint_2(0) = 420: mSpacePoint_2(1) = 297
    
    '*****  PUNKTY NA PAPERSPACE *****
    'Lewy dolny punkt
    pSpacePoint_1(0) = 6090.6: pSpacePoint_1(1) = -4981.22
    'Prawy gorny punkt
    pSpacePoint_2(0) = 6510.6: pSpacePoint_2(1) = -4684.22
    
    '*****  DYSTANS MIÊDZY LAYOUTAMI *****
    'Dystans pomiedzy kolejnymi arkuszami
    distanceBetween = ThisDrawing.Utility.GetDistance(, "Podaj odstêp: ")
    
    'Podaj iloœc layoutów do stworzenia
    nOfLayouts = ThisDrawing.Utility.GetInteger("Podaj iloœæ layoutów do stworzenia: ")
    
    '*****  KASOWANIE ISTNIEJ¥CYCH LAYOUTÓW *****
    MsgBox "Kasowanie istniej¹cych layoutów", vbInformation
    
    'Skasowanie wszystkich layoutów innych niz 1
    For Each Layout In ThisDrawing.Layouts
        If Layout.Name <> "1" And Layout.Name <> "Model" Then
            Layout.Delete
        End If
    Next Layout
    
    '*****  TWORZENIE NOWYCH LAYOUTÓW *****
    MsgBox "Tworzenie nowych layoutów", vbInformation
    
    'Zaczynamy od konca i iterujemy do poczatku
    For i = nOfLayouts To 1 Step -1
        'Kopiujemy layout numer 1
        ThisDrawing.SendCommand "_layout _C" & vbCr & "1" & vbCr & CStr(i) & vbCr
        
    Next i
    
    Dim pSpacePt1_X As String, pSpacePt1_Y As String
    Dim pSpacePt2_X As String, pSpacePt2_Y As String
    
    pSpacePt1_X = ThisDrawing.Utility.RealToString(pSpacePoint_1(0), acDecimal, 2)
    pSpacePt1_Y = ThisDrawing.Utility.RealToString(pSpacePoint_1(1), acDecimal, 2)
    pSpacePt2_X = ThisDrawing.Utility.RealToString(pSpacePoint_2(0), acDecimal, 2)
    pSpacePt2_Y = ThisDrawing.Utility.RealToString(pSpacePoint_2(1), acDecimal, 2)

    Dim aPVPort As AcadPViewport
    Dim aEnt As AcadEntity
    
    Dim lowerLeftPt(2) As Double, upperRightPt(2) As Double
    
    lowerLeftPt(0) = mSpacePoint_1(0)
    lowerLeftPt(1) = mSpacePoint_1(1)
    lowerLeftPt(2) = 0
    
    upperRightPt(0) = mSpacePoint_2(0)
    upperRightPt(1) = mSpacePoint_2(1)
    upperRightPt(2) = 0
    
    lowerLeftPt(1) = mSpacePoint_1(1)
    lowerLeftPt(2) = 0
    
    upperRightPt(1) = upperRightPt(1)
    upperRightPt(2) = upperRightPt(2)
    
    For i = nOfLayouts To 1 Step -1
        'Ustawiamy nowo utworzony layout jako aktywny
        ThisDrawing.ActiveLayout = ThisDrawing.Layouts(CStr(i))
        'Wyslanie komendy alignspace
        
        For Each aEnt In ThisDrawing.ActiveLayout.Block
            If TypeOf aEnt Is AcadPViewport Then
                Set aPVPort = aEnt
                Exit For
            End If
        Next aEnt
        
        lowerLeftPt(0) = mSpacePoint_1(0) + (i - 1) * distanceBetween
        upperRightPt(0) = mSpacePoint_2(0) + (i - 1) * distanceBetween

        
        Debug.Print lowerLeftPt(0), lowerLeftPt(1), upperRightPt(0), upperRightPt(1)
        
        ThisDrawing.MSpace = True
        ThisDrawing.Application.ZoomWindow lowerLeftPt, upperRightPt
        ThisDrawing.MSpace = False
        
    Next i
    
    ThisDrawing.SetVariable "LAYOUTREGENCTL", layoutRegenCTL
    
End Sub

Sub AlignItems()
    Dim refPoints() As Variant
    Dim pointPicked As Variant
    Dim cnt As Long
    ThisDrawing.Utility.InitializeUserInput 128, "[K]"
    ReDim Preserve refPoints(2, 0)
    
    On Error Resume Next
    cnt = 0
    Do
        pointPicked = Null
        pointPicked = ThisDrawing.Utility.GetPoint(prompt:="Podaj punkt odniesienia lub [Koniec]")
        
        If Err Then
            Err.Clear
            Dim uInput As String
            uInput = ThisDrawing.Utility.GetInput
            If uInput = "K" Then
                Exit Do
            End If
        Else
            ReDim Preserve refPoints(2, cnt)

            refPoints(0, cnt) = pointPicked(0)
            refPoints(1, cnt) = pointPicked(1)
            refPoints(2, cnt) = pointPicked(2)
            cnt = cnt + 1
        End If
    Loop While (True)
    
    If IsEmpty(refPoints) Then
        Exit Sub
    End If
    
    Dim startPoint As Variant
    Dim distanceBetween As Double
    Dim distance_X As Double
    Dim distance_Y As Double
    
    startPoint = ThisDrawing.Utility.GetPoint(prompt:="Podaj punkt pocz¹tkowy: ")
    distanceBetween = ThisDrawing.Utility.GetDistance(prompt:="Podaj odstêp: ")
    distance_X = ThisDrawing.Utility.GetDistance(startPoint, "Podaj d³ugoœæ na osi X zakresu obiektów: ")
    distance_Y = ThisDrawing.Utility.GetDistance(startPoint, "Podaj d³ugoœæ na osi Y zakresu obiektów: ")
    
    Dim i As Integer

    Dim selectionSet As AcadSelectionSet
    
    Set selectionSet = ThisDrawing.SelectionSets.Item("SS1")
    If Err Then
        Set selectionSet = ThisDrawing.SelectionSets.Add("SS1")
        Err.Clear
    End If
    
    Dim ref_X As Double
    Dim ref_Y As Double
    Dim ref_Z As Double
    
    Dim moveTo_Point(2) As Double
    Dim moveFrom_Point(2) As Double
    Dim slctPolyPts(0 To 11) As Double
    Const diff As Double = 0.01
    
    For i = 0 To UBound(refPoints, 2)
        selectionSet.Clear
        ref_X = refPoints(0, i)
        ref_Y = refPoints(1, i)
        ref_Z = refPoints(2, i)
        
        
        slctPolyPts(0) = ref_X - diff: slctPolyPts(1) = ref_Y - diff: slctPolyPts(2) = 0
        slctPolyPts(3) = ref_X + distance_X + diff: slctPolyPts(4) = ref_Y - diff: slctPolyPts(5) = 0
        slctPolyPts(6) = ref_X + distance_X + diff: slctPolyPts(7) = ref_Y + distance_Y + diff: slctPolyPts(8) = 0
        slctPolyPts(9) = ref_X - diff: slctPolyPts(10) = ref_Y + distance_Y + diff: slctPolyPts(11) = 0
        
        selectionSet.SelectByPolygon acSelectionSetWindowPolygon, slctPolyPts
        
        moveFrom_Point(0) = ref_X: moveFrom_Point(1) = ref_Y: moveFrom_Point(2) = ref_Z
        moveTo_Point(0) = startPoint(0) + i * distanceBetween: moveTo_Point(1) = startPoint(1): moveTo_Point(2) = startPoint(2)
        
        Dim aEnt As AcadEntity
        For Each aEnt In selectionSet
            aEnt.Move moveFrom_Point, moveTo_Point
        Next aEnt
        
    Next i
End Sub

Private Function GetLayoutByName(num As Integer) As AcadLayout
    Dim aLay As AcadLayout
    Set GetLayoutByName = Null
    For Each aLay In ThisDrawing.Layouts
        If aLay.Name = CStr(num) Then
            Set GetLayoutByName = aLay
            Exit Function
        End If
    Next aLay
End Function

Sub RenumberLayouts()
    Dim aLayout As AcadLayout
    Dim num As Integer
    Dim j As Integer
    For j = ThisDrawing.Layouts.count To 1 Step -1
        Set aLayout = ThisDrawing.Layouts(j)
        If aLayout.Name <> "Model" And aLayout.Name <> "1" Then
            aLayout.Name = CStr((CInt(aLayout.Name) + 1))
        End If
    Next j
End Sub

Function PrintArray(arr As Variant)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Debug.Print "Array[" & CStr(i) & "] = ", CStr(arr(i))
    Next i
End Function


Sub ArrangeLayouts_A4()
    '*****  DEKLARACJA ZMIENNYCH *****
    Dim i As Integer, nOfLayouts As Integer
    Dim distanceBetween As Double
    Dim Layout As AcadLayout
    Dim mSpacePoint_1(0 To 2) As Double
    Dim mSpacePoint_2(0 To 2) As Double
    Dim pSpacePoint_1(0 To 2) As Double
    Dim pSpacePoint_2(0 To 2) As Double
    
    On Error GoTo 0
    
    Dim layoutRegenCTL As Integer
    layoutRegenCTL = ThisDrawing.GetVariable("LAYOUTREGENCTL")
    
    ThisDrawing.SetVariable "LAYOUTREGENCTL", 0
    
    '*****  PUNKTY NA MODELSPACE *****
    'Lewy dolny punkt
    mSpacePoint_1(0) = 0: mSpacePoint_1(1) = 0
    'Prawy gorny punkt
    mSpacePoint_2(0) = 180: mSpacePoint_2(1) = 287
    
    '*****  PUNKTY NA PAPERSPACE *****
    'Lewy dolny punkt
    pSpacePoint_1(0) = -31: pSpacePoint_1(1) = 21
    'Prawy gorny punkt
    pSpacePoint_2(0) = 149: pSpacePoint_2(1) = 308
    
    '*****  DYSTANS MIÊDZY LAYOUTAMI *****
    
    'Dystans pomiedzy kolejnymi arkuszami
    distanceBetween = ThisDrawing.Utility.GetDistance(, "Podaj odstêp: ")
    
    'Podaj iloœc layoutów do stworzenia
    nOfLayouts = ThisDrawing.Utility.GetInteger("Podaj iloœæ layoutów do stworzenia: ")
    
    
    ThisDrawing.Utility.InitializeUserInput 2, "T N"
    
    Dim deleteLayouts As String
    deleteLayouts = ThisDrawing.Utility.GetString(1, prompt:="Usun¹æ wszystkie layouty [Tak/Nie]?")
    
    Dim lastLayoutNumber As Long

    Select Case (deleteLayouts)
    
        Case "T"
            '*****  KASOWANIE ISTNIEJ¥CYCH LAYOUTÓW *****
            MsgBox "Kasowanie istniej¹cych layoutów", vbInformation
    
            'Skasowanie wszystkich layoutów innych niz 1
            For Each Layout In ThisDrawing.Layouts
                If Layout.Name <> "1" And Layout.Name <> "Model" Then
                    Layout.Delete
                End If
            Next Layout
            
            lastLayoutNumber = 1
        Case "N"
            lastLayoutNumber = CInt(ThisDrawing.Layouts.Item(ThisDrawing.Layouts.count - 2).Name)
            
            If (nOfLayouts <= lastLayoutNumber) Then
                MsgBox "Podana iloœæ layoutów do stworzenia jest mniejsza ni¿ liczba istniej¹cych layoutów", vbCritical
                Exit Sub
            End If
        Case Else
    End Select
    
    '*****  TWORZENIE NOWYCH LAYOUTÓW *****
    MsgBox "Tworzenie nowych layoutów", vbInformation
    
    Dim aPVPort As AcadPViewport
    Dim aEnt As AcadEntity
    Dim lowerLeftPt(2) As Double, upperRightPt(2) As Double
        
    lowerLeftPt(0) = mSpacePoint_1(0): lowerLeftPt(1) = mSpacePoint_1(1): lowerLeftPt(2) = 0
    
    upperRightPt(0) = mSpacePoint_2(0): upperRightPt(1) = mSpacePoint_2(1): upperRightPt(2) = 0
    
    lowerLeftPt(1) = mSpacePoint_1(1): lowerLeftPt(2) = 0
    
    upperRightPt(1) = upperRightPt(1): upperRightPt(2) = upperRightPt(2)
           
    'Zaczynamy od konca i iterujemy do poczatku
    For i = nOfLayouts To lastLayoutNumber Step -1
        'Kopiujemy layout numer 1
        ThisDrawing.SendCommand "_layout _C" & vbCr & "1" & vbCr & CStr(i) & vbCr

        ThisDrawing.ActiveLayout = ThisDrawing.Layouts(CStr(i))
        
        For Each aEnt In ThisDrawing.ActiveLayout.Block
            If TypeOf aEnt Is AcadPViewport Then
                Set aPVPort = aEnt
                Exit For
            End If
        Next aEnt
        
        lowerLeftPt(0) = mSpacePoint_1(0) + (i - 1) * distanceBetween
        upperRightPt(0) = mSpacePoint_2(0) + (i - 1) * distanceBetween
        
        ThisDrawing.MSpace = True
        ThisDrawing.Application.ZoomWindow lowerLeftPt, upperRightPt
        ThisDrawing.MSpace = False
    Next i
    
    ThisDrawing.SetVariable "LAYOUTREGENCTL", layoutRegenCTL
End Sub

Sub AutoNumberHookUps()

    '*****  DEKLARACJA ZMIENNYCH *****
    
    Dim blockName As String
    Dim selectionSet As AcadSelectionSet
    Dim blockRefObj As AcadBlockReference

    'Do przechowywania atrybutów
    Dim varAttributes As Variant
    Dim attRef As AcadAttributeReference
    
    'Do przechowywania referencji bloków
    Dim blockRefArray() As AcadBlockReference

    Dim blockCount As Integer, i As Integer, j As Integer, pageCount As Integer
    Dim strAttributes As String
    strAttributes = ""
    
    '*****  WYBÓR BLOKU *****
    Dim pickedBlockRef As AcadBlockReference
    
    Set pickedBlockRef = PickBlockReference
    If Not TypeOf pickedBlockRef Is AcadBlockReference Then
        MsgBox "Z³y wybór!", vbCritical
        Exit Sub
    Else
        blockName = pickedBlockRef.EffectiveName
    End If
    
    '*****  WYBÓR TRYBU (M / E) *****
    
    'Zainicjalizuj s³owa kluczowe M i E
    ThisDrawing.Utility.InitializeUserInput Bits:=1, keywordList:="M E"
    
    Dim numberMode As String
    numberMode = UCase(ThisDrawing.Utility.GetString(0, vbLf & "Mechaniczne/Elektryczne [M/E]?"))
    
    'Zakoñcz jeœli wybór jest niepoprawny
    If numberMode <> "M" And numberMode <> "E" Then
        MsgBox "B³êdna wartoœæ.", vbCritical
        Exit Sub
    End If
    
    '*****  LICZBA STARTOWA *****
    'Podaj liczbê startow¹ do numeracji
    pageCount = ThisDrawing.Utility.GetInteger(prompt:="Podaj liczbê startow¹ do numeracji: ")
    If pageCount <= 0 Then
        MsgBox "Niepoprawna wartoœæ startowa do numeracji!", vbCritical
    End If
    
    i = 0
    j = 0
    blockCount = 0
    
    
    '*****  FILTROWANIE *****
    Set selectionSet = GetSelectionSet
    selectionSet.Clear
    selectionSet.SelectOnScreen
    
    '*****  DODANIE DO TABLICY *****

    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = AddToBlockRefArray(selectionSet, blockName)

    '*****  SORTOWANIE BLOKÓW *****
    
    Dim attNum As Long: attNum = 0
    
    'Sortowanie po rosn¹cym punkcie wstawienia X
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, SortOrder:=acAscending)
    
    '*****  ITEROWANIE PO BLOKACH DYNAMICZNYCH *****
    
    Dim dybprop As Variant
    Dim blockObj As AcadBlockReference
    'Zmiana atrybutu numer 0 w bloku ( = ARKUSZ)
    For j = LBound(blockRefArray) To UBound(blockRefArray)
        Set blockObj = blockRefArray(j)
        If blockObj.IsDynamicBlock Then
            dybprop = blockObj.GetDynamicBlockProperties
            For i = LBound(dybprop) To UBound(dybprop)
                If dybprop(i).PropertyName = "Visibility1" Then
                    Select Case numberMode
                        Case "M"
                            If dybprop(i).value = "Mechaniczny" Then
                                varAttributes = blockObj.GetAttributes
                                
                                Set attRef = varAttributes(attNum)
                                attRef.textString = "M" & Format(CStr(pageCount), "000")
                                pageCount = pageCount + 1
                            End If
                        Case "E"
                            If dybprop(i).value = "Elektryczny" Then
                                varAttributes = blockObj.GetAttributes
                                
                                Set attRef = varAttributes(attNum)
                                attRef.textString = "E" & Format(CStr(pageCount), "000")
                                pageCount = pageCount + 1
                            End If
                    End Select
                End If
            Next i
        End If
    Next j
    
    '*****  INFORMACJA KOÑCOWA *****
    'Wyœwietlenie liczby stron do zmiany w bloku
    MsgBox "Zakoñczono numerowanie stron." & vbCrLf & _
           "Liczba stron: " & CStr(pageCount - 1)
End Sub

Public Sub AlignItemsVertically()
    Dim blockName As String
    Dim selectionSet As AcadSelectionSet
    Dim blockRefObj As AcadBlockReference
    Dim distanceBetween As Double
    
    Dim blockRefArray() As AcadBlockReference
    Dim blockCount As Integer, i As Integer, j As Integer
    Dim refPoint As Variant
    Dim newPoint As Variant
    
    Dim ent As AcadEntity
    
    If ThisDrawing.SelectionSets.count = 0 Then
        ThisDrawing.SelectionSets.Add ("SS1")
    End If
    
    Set selectionSet = ThisDrawing.SelectionSets.Item(0)
    selectionSet.Clear
    selectionSet.SelectOnScreen
    
    refPoint = ThisDrawing.Utility.GetPoint(, "Podaj punkt pocz¹tkowy:")
    distanceBetween = ThisDrawing.Utility.GetDistance(, "Podaj odstêp:")
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = AddToBlockRefArray(blockRefArray, selectionSet)
    'Sortowanie po rosn¹cym punkcie wstawienia X
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, SortOrder:=acAscending)
    
    newPoint = refPoint
    
    For i = LBound(blockRefArray) To UBound(blockRefArray)
        newPoint(1) = refPoint(1) - i * distanceBetween
        blockRefArray(i).insertionPoint = newPoint
    Next i
End Sub

Sub AutoNumberAttributes_X()
    Call AutoNumberAttributes(False)
End Sub

Sub AutoNumberAttributes_Y()
    Call AutoNumberAttributes(True, acDescending)
End Sub

Sub AutoNumberAttributes(Optional SortByY As Boolean = False, Optional SortOrder As Integer = acAscending)

    '*****  DEKLARACJA ZMIENNYCH *****
    Dim blockName As String
    
    Dim selectionSet As AcadSelectionSet                'Zmienna do zestawu wybranych elementów
    Dim blockRefObj As AcadBlockReference
    Dim varAttributes As Variant                        'Do przechowywania atrybutów
    
    Dim blockRefArray() As AcadBlockReference           'Do przechowywania referencji bloków
    ReDim blockRefArray(0)                              'Tablica dynamiczna referencji bloków
    
    '*****  WYBÓR BLOKU *****
    
    Dim blockObj As AcadBlockReference
    Set blockObj = PickBlockReference
    
    If blockObj Is Nothing Then
        MsgBox "Nie wybrano bloku!", vbCritical
        Exit Sub
    Else
        blockName = blockObj.EffectiveName
    End If
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(1) As Variant, filterData(1) As Variant
    
    'Filtrowanie po typie obiektu - AcadBlockReference
    filterType(0) = acFilterByObjectType: filterData(0) = acBlockRef
    'Filtrowanie po nazwie obiektu - BlockName
    filterType(1) = acFilterByObjectName: filterData(1) = blockName
    'Wybor selection setem z filtrem
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    '*****  DODANIE DO TABLICY *****
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = GetItemsFromSelectionSet(selectionSet)

    '*****  SORTOWANIE BLOKÓW *****

    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)

    '*****  WYBÓR NUMERU ATRYBUTU *****
    'Wybranie numeru atrybutu z bloku do numeracji
    Dim tagNumber As Integer
    tagNumber = GetTagListAndSelect(blockRefArray)
    
    '*****  POBRANIE TABLICY TAGÓW O TEJ SAMEJ NAZWIE *****
    Dim tagArray() As Integer
    tagArray = GetTagNumberArray(blockRefArray(0), tagNumber)
    
    '*****  POBRANIE WARTOŒCI STARTOWEJ I INKREMENTACJI *****
    Dim startingValue As Integer, incrementValue As Integer
    ThisDrawing.Utility.InitializeUserInput 1 + 2
    startingValue = ThisDrawing.Utility.GetInteger("Podaj wartoœæ pocz¹tkow¹ numeracji: ")  'Podaj wartoœc pocz¹tkow¹
    incrementValue = ThisDrawing.Utility.GetInteger("Podaj wartoœæ inkrementacji: ")        'Podaj wartoœæ inkrementacji
    
    '*****  ITEROWANIE PO ATRYBUTOWACH *****
    Dim nextNumber As Integer, lastNumber As Integer, tagVarNumber As Variant
    nextNumber = startingValue
    
    Dim attRef As AcadAttributeReference, i As Integer
    'Zmiana atrybutu w bloku
    For i = LBound(blockRefArray) To UBound(blockRefArray)
        lastNumber = nextNumber                                                    'Zmienna do uzycia w prompcie, przechowuje ostatnia przypisana wartoœæ
        varAttributes = blockRefArray(i).GetAttributes                             'Pobranie zestawu atrybutów dla bloku[i]
        
        For Each tagVarNumber In tagArray
            Set attRef = varAttributes(tagVarNumber)
            attRef.textString = CStr(nextNumber)                                   'Przypisanie do atrybutu nowej wartoœci
        Next tagVarNumber
                                     
        nextNumber = nextNumber + incrementValue                                   'Okreœlenie kolejnej wartoœci
    Next i
    
    '*****  WIADOMOŒÆ KOÑCOWA *****
    Dim promptMsg As String
    promptMsg = "Zakoñczono numeracjê." & vbNewLine
    promptMsg = promptMsg & "Wartoœæ pocz¹tkowa: " & vbTab & startingValue & vbNewLine
    promptMsg = promptMsg & "Inkrementacja:      " & vbTab & incrementValue & vbNewLine
    promptMsg = promptMsg & "Wartoœæ koñcowa:    " & vbTab & lastNumber
    MsgBox prompt:=promptMsg, Buttons:=vbInformation, Title:="Zakoñczono"
End Sub

Sub ExportAllTables()
    Dim acTable As acadTable
    Dim tRows As Integer, tCols As Integer
    Dim arr() As Variant
    Dim selectionSet As AcadSelectionSet
    Dim tableCount As Integer, mTextCount As Integer
    
    Dim tableArray() As acadTable                       'Do przechowywania tabelek
    Dim mtextArray() As AcadMText                       'Do przechowywania MText
    
    Dim ent As AcadEntity                               'Do iterowania po obiektach w selection secie
    Dim blockCount As Integer                           'Do zliczenia iloœci obiektów w selection secie
    
    Dim temp As AcadEntity
    Dim i As Integer, j As Integer
    
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWksht As Excel.Worksheet
    Dim xlRange As Excel.Range
    
    Set selectionSet = ThisDrawing.ActiveSelectionSet
    selectionSet.Clear
    selectionSet.SelectOnScreen
    
    tableCount = 0
    mTextCount = 0
    
    ThisDrawing.Utility.InitializeUserInput 1, "M E PEBX"
    Dim answer As String
    Dim layerName As String
    
    answer = ThisDrawing.Utility.GetString(1, "Mechaniczny/Elektryczny/Skrzynka [M/E/PEBX]")
    
    Select Case True
        Case answer = "M"
            layerName = "TabelkaZTagamiMechaniczne"
        Case answer = "E"
            layerName = "TabelkaZTagamiElektryczne"
        Case answer = "PEBX"
            layerName = "TabelkaZTagamiPEBX"
        Case Else
            MsgBox "B³êdne dane!", vbCritical
            Exit Sub
    End Select
    
    For Each ent In selectionSet
        If ent.Layer = layerName Then                'Jesli obiekt w selection secie jest blokiem, to dodaj do tablicy
            If TypeOf ent Is acadTable Then
                ReDim Preserve tableArray(tableCount)       'Dynamiczna alokacja wymiarów tablicy
                Set tableArray(tableCount) = ent            'Przypisanie do elementu tablicy kolejnego obiektu
                tableCount = tableCount + 1                 'Inkrementacja iloœci obiektów
            ElseIf TypeOf ent Is AcadMText Then
                ReDim Preserve mtextArray(mTextCount)       'Dynamiczna alokacja wymiarów tablicy
                Set mtextArray(mTextCount) = ent            'Przypisanie do elementu tablicy kolejnego obiektu
                mTextCount = mTextCount + 1                 'Inkrementacja iloœci obiektów
            End If
        End If
    Next ent
    
    
    For i = LBound(tableArray) To UBound(tableArray)
        For j = LBound(tableArray) To (UBound(tableArray) - i - 1)
            If tableArray(j).insertionPoint(0) > tableArray(j + 1).insertionPoint(0) Then
                Set temp = tableArray(j)
                Set tableArray(j) = tableArray(j + 1)
                Set tableArray(j + 1) = temp
            End If
        Next j
    Next i
    
    For i = LBound(mtextArray) To UBound(mtextArray)
        For j = LBound(mtextArray) To (UBound(mtextArray) - i - 1)
            If mtextArray(j).insertionPoint(0) > mtextArray(j + 1).insertionPoint(0) Then
                Set temp = mtextArray(j)
                Set mtextArray(j) = mtextArray(j + 1)
                Set mtextArray(j + 1) = temp
            End If
        Next j
    Next i
    
    Set xlApp = New Excel.Application
    Set xlWb = xlApp.Workbooks.Add()
    Set xlWksht = xlWb.Worksheets(1)
    
    Dim c As Integer
    Dim text As String
    c = 0
    For c = 0 To UBound(tableArray)
    
        tRows = tableArray(c).Rows
        tCols = tableArray(c).Columns
        Set xlWksht = xlWb.Worksheets.Add()
        xlWksht.Name = mtextArray(c).textString
        
        For i = 0 To tRows
            For j = 0 To tCols
                text = tableArray(c).GetText(i, j)
                If InStr(1, text, ";", vbBinaryCompare) <> 0 Then
                    text = Split(text, ";")(1)
                End If
                xlWksht.Cells(i + 1, j + 1) = text
            Next j
            
            If i = 0 Then
                xlWksht.Cells(i + 1, tCols + 1) = "DWG"
            ElseIf i < tRows Then
                xlWksht.Cells(i + 1, tCols + 1) = mtextArray(c).textString
            End If
            
        Next i
        
        With xlWksht.UsedRange.Cells
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .EntireColumn.AutoFit
            .EntireRow.AutoFit
        End With
    Next c
    
    xlApp.Visible = True
    xlWksht.Columns.AutoFit
    
    Dim allWkSht As Excel.Worksheet
    Dim wkSheets As Excel.Worksheet
    
    Dim lastRow As Integer, lastCol As Integer
    lastRow = 0
    
    Set allWkSht = xlWb.Worksheets.Add()
    allWkSht.Name = "Zbiorcze"
    
    Excel.Application.DisplayAlerts = False
    
    For Each xlWksht In xlWb.Worksheets
        If (xlWksht.Name <> allWkSht.Name) Then
        
            lastRow = allWkSht.Cells.SpecialCells(xlCellTypeLastCell).Row
            xlWksht.Range("A1").CurrentRegion.Copy Destination:=allWkSht.Cells(lastRow + 1, 1)
            
        End If
    Next xlWksht
    
    Excel.Application.DisplayAlerts = True
    
    With allWkSht.UsedRange.Cells
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .EntireColumn.AutoFit
            .EntireRow.AutoFit
    End With

End Sub

Sub FillTable()
    Dim aTable As acadTable
    Dim tRows As Integer, tCols As Integer
    Dim acEnt As AcadEntity

    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook, thisWb As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    
    Dim classDict As New Scripting.Dictionary
    Dim classDataDict As New Scripting.Dictionary
    Dim wkShtName As String
    
    Dim className As String
    
    ThisDrawing.Utility.GetEntity acEnt, Array(0, 0, 0), "Select table: "
    If TypeOf acEnt Is acadTable Then
        Set aTable = acEnt
    Else
        MsgBox "Z³y wybór"
    End If
    
    className = ThisDrawing.Utility.GetString(1, "Podaj klasê: ")
    className = UCase(Excel.WorksheetFunction.Clean(className))
    
    If className = "" Then
        MsgBox "Wrong input"
        Exit Sub
    End If
    
    wkShtName = ""
    Set xlApp = GetObject(, "Excel.Application")
    
    For Each xlWb In xlApp.Workbooks
        If xlWb.Name = wkShtName Then
            Set thisWb = xlWb
        End If
    Next xlWb
    
    Dim data(0 To 1) As String
    Dim cell As Excel.Range
    Dim i As Integer, j As Integer
    Dim key As String
    For Each xlSht In thisWb.Worksheets
        If xlSht.Name = className Then
            For i = 1 To xlSht.Range("A1").CurrentRegion.Cells.SpecialCells(xlCellTypeLastCell).Row
                key = xlSht.Cells(i, 1).text
                data(0) = xlSht.Cells(i, 2).text
                data(1) = xlSht.Cells(i, 3).text
                If key <> "" Then
                    classDataDict.Add key:=key, Item:=data
                Else
                    Exit For
                End If
                
            Next i
        End If
    Next xlSht
    
    tRows = aTable.Rows
    tCols = aTable.Columns
    Dim tag As Variant

    For i = 1 To tRows
        tag = aTable.GetCellValue(i, 0)
        If Not IsEmpty(tag) Then
            If classDataDict.Exists(tag) = True Then
                aTable.SetCellValue i, 1, classDataDict(tag)(0)
                aTable.SetCellValue i, 2, classDataDict(tag)(1)
            Else
                MsgBox "Brak tagu: " & tag
            End If
        End If
    Next i
    
End Sub

Sub FillTagTable()

    Dim acEnt As AcadEntity
    Dim aTable As acadTable
    
    ThisDrawing.Utility.GetEntity acEnt, Array(0, 0, 0), "Select table: "
    If TypeOf acEnt Is acadTable Then
        Set aTable = acEnt
    Else
        MsgBox "Z³y wybór"
    End If
    
    Dim wkShtName As String
    wkShtName = ""
    
    Dim xlApp As Excel.Application
    Set xlApp = GetObject(, "Excel.Application")
    
    Dim xlWb As Excel.Workbook, thisWb As Excel.Workbook
    For Each xlWb In xlApp.Workbooks
        If xlWb.Name = wkShtName Then
            Set thisWb = xlWb
        End If
    Next xlWb
    
    Dim data(0 To 1) As String
    Dim cell As Excel.Range
    Dim i As Integer, j As Integer
    Dim key As String
    Dim xlSht As Excel.Worksheet
    Dim ActiveSheet As Excel.Worksheet
    
    Set ActiveSheet = thisWb.Worksheets("NEW")
    
    Dim tRows As Integer, tCols As Integer
    tRows = aTable.Rows
    tCols = aTable.Columns
    
    Dim rowNum As Long
    Dim pid As Variant, tag As Variant, linia As Variant, urzadzenie As Variant, uwagi As Variant
    Dim cellTextHeight As Double, cellHeight As Double
    
    cellHeight = 9.43
    cellTextHeight = 1.8
    
    With Excel.WorksheetFunction
        For i = 1 To (tRows - 1)
            tag = aTable.GetCellValue(i, 1)
            rowNum = .Match(tag, ActiveSheet.Range("tblData[TAG]"), 0#)
            
            pid = .Index(ActiveSheet.Range("tblData[PID]"), rowNum)
            linia = .Index(ActiveSheet.Range("tblData[Nowy Nr ruroci¹gu]"), rowNum)
            urzadzenie = .Index(ActiveSheet.Range("tblData[Urz¹dzenie]"), rowNum)
            uwagi = "-"
            
            aTable.SetCellValue i, 2, pid(1)
            aTable.SetCellTextHeight i, 2, cellTextHeight

            
            aTable.SetCellValue i, 3, linia(1)
            aTable.SetCellTextHeight i, 3, cellTextHeight

            aTable.SetCellValue i, 4, urzadzenie(1)
            aTable.SetCellTextHeight i, 4, cellTextHeight

            
            aTable.SetCellValue i, 5, uwagi
            aTable.SetCellTextHeight i, 5, cellTextHeight
            aTable.SetRowHeight i, cellHeight
        Next i
    End With
    
End Sub

Sub InsertData()
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook, thisWb As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    
    Dim classDict As New Scripting.Dictionary
    Dim classDataDict As New Scripting.Dictionary
    Dim wkShtName As String
    
    Dim className As String
    
    className = ThisDrawing.Utility.GetString(1, "Podaj klasê: ")
    className = Excel.WorksheetFunction.Clean(className)
    
    If className = "" Then
        MsgBox "Wrong input"
        Exit Sub
    End If
    
    wkShtName = ""
    Set xlApp = GetObject(, "Excel.Application")
    
    For Each xlWb In xlApp.Workbooks
        If xlWb.Name = wkShtName Then
            Set thisWb = xlWb
        End If
    Next xlWb
    
    Dim data(0 To 1) As String
    Dim cell As Excel.Range
    Dim i As Integer, j As Integer
    
    For Each xlSht In thisWb.Worksheets
        If xlSht.Name = className Then
            For i = 1 To xlSht.Cells.SpecialCells(xlCellTypeLastCell).Row
                data(0) = xlSht.Cells(i, 2).text
                data(1) = xlSht.Cells(i, 3).text
                classDataDict.Add key:=xlSht.Cells(i, 1).text, Item:=data
            Next i
        End If
    Next xlSht
    
    Dim key As Variant
    For Each key In classDataDict.Keys
        MsgBox key & " : " & classDataDict(key)(0)
    Next key
    
    
End Sub

Private Function AddToBlockRefArray(selectionSet As AcadSelectionSet, Optional blockName As String = "") As Variant
    Dim ent As AcadEntity                               'Do iterowania po obiektach w selection secie
    Dim blockCount As Integer                           'Do zliczenia iloœci obiektów w selection secie
    
    Dim temp() As AcadBlockReference
    ReDim Preserve temp(0)
    Select Case (blockName)
        Case ""
            For Each ent In selectionSet
                If TypeOf ent Is AcadBlockReference Then        'Jesli obiekt w selection secie jest blokiem, to dodaj do tablicy
                    ReDim Preserve temp(blockCount)   'Dynamiczna alokacja wymiarów tablicy
                    Set temp(blockCount) = ent        'Przypisanie do elementu tablicy kolejnego obiektu
                    blockCount = blockCount + 1                 'Inkrementacja iloœci obiektów
                End If
            Next ent
        Case Else
            For Each ent In selectionSet
                If TypeOf ent Is AcadBlockReference Then        'Jesli obiekt w selection secie jest blokiem, to dodaj do tablicy
                    If ent.EffectiveName = blockName Then
                        ReDim Preserve temp(blockCount)   'Dynamiczna alokacja wymiarów tablicy
                        Set temp(blockCount) = ent        'Przypisanie do elementu tablicy kolejnego obiektu
                        blockCount = blockCount + 1                 'Inkrementacja iloœci obiektów
                    End If
                End If
            Next ent
    End Select

    
    AddToBlockRefArray = temp
End Function

Private Function BubbleSortByInsertionPoint(ArrayToBeSorted() As AcadBlockReference, Optional SortOrder As Integer = acAscending, Optional SortByY As Boolean = False)
    'Algorytm sortowania b¹belkowego, sortujemy malej¹co po pozycji X bloku
    Dim temp As AcadBlockReference
    Dim i As Integer, j As Integer
    
    BubbleSortByInsertionPoint = Array(" ")
    
    Dim insNumPt As Integer
    
    Select Case SortByY
        Case True
            insNumPt = 1
        Case False
            insNumPt = 0
    End Select
    
    Select Case SortOrder
        Case acAscending:
            'Sortowanie rosn¹ce
            For i = LBound(ArrayToBeSorted) To UBound(ArrayToBeSorted)
                For j = LBound(ArrayToBeSorted) To (UBound(ArrayToBeSorted) - i - 1)
                    If ArrayToBeSorted(j).insertionPoint(insNumPt) > ArrayToBeSorted(j + 1).insertionPoint(insNumPt) Then
                        Set temp = ArrayToBeSorted(j)
                        Set ArrayToBeSorted(j) = ArrayToBeSorted(j + 1)
                        Set ArrayToBeSorted(j + 1) = temp
                    End If
                Next j
            Next i
        Case acDescending:
            'Sortowanie malej¹ce
                For i = LBound(ArrayToBeSorted) To UBound(ArrayToBeSorted)
                    For j = LBound(ArrayToBeSorted) To (UBound(ArrayToBeSorted) - i - 1)
                        If ArrayToBeSorted(j).insertionPoint(insNumPt) < ArrayToBeSorted(j + 1).insertionPoint(insNumPt) Then
                            Set temp = ArrayToBeSorted(j)
                            Set ArrayToBeSorted(j) = ArrayToBeSorted(j + 1)
                            Set ArrayToBeSorted(j + 1) = temp
                        End If
                    Next j
                Next i
        Case Else:
            'Inna wartoœæ ni¿ acAscending lub acDescending
        Exit Function
    End Select
    'Zwrócenie posortowanej tablicy
    BubbleSortByInsertionPoint = ArrayToBeSorted
End Function


Private Function BubbleSortByInsertionPoint_MText(ArrayToBeSorted() As AcadMText, Optional SortOrder As Integer = acAscending, Optional SortByY As Boolean = False)
    'Algorytm sortowania b¹belkowego, sortujemy malej¹co po pozycji X bloku
    Dim temp As AcadMText
    Dim i As Integer, j As Integer
    
    Dim insNumPt As Integer
    
    Select Case SortByY
        Case True
            insNumPt = 1
        Case False
            insNumPt = 0
    End Select
    
    Select Case SortOrder
        Case acAscending:
            'Sortowanie rosn¹ce
            For i = LBound(ArrayToBeSorted) To UBound(ArrayToBeSorted)
                For j = LBound(ArrayToBeSorted) To (UBound(ArrayToBeSorted) - i - 1)
                    If ArrayToBeSorted(j).insertionPoint(insNumPt) > ArrayToBeSorted(j + 1).insertionPoint(insNumPt) Then
                        Set temp = ArrayToBeSorted(j)
                        Set ArrayToBeSorted(j) = ArrayToBeSorted(j + 1)
                        Set ArrayToBeSorted(j + 1) = temp
                    End If
                Next j
            Next i
        Case acDescending:
            'Sortowanie malej¹ce
                For i = LBound(ArrayToBeSorted) To UBound(ArrayToBeSorted)
                    For j = LBound(ArrayToBeSorted) To (UBound(ArrayToBeSorted) - i - 1)
                        If ArrayToBeSorted(j).insertionPoint(insNumPt) < ArrayToBeSorted(j + 1).insertionPoint(insNumPt) Then
                            Set temp = ArrayToBeSorted(j)
                            Set ArrayToBeSorted(j) = ArrayToBeSorted(j + 1)
                            Set ArrayToBeSorted(j + 1) = temp
                        End If
                    Next j
                Next i
        Case Else:
            'Inna wartoœæ ni¿ acAscending lub acDescending
			Exit Function
    End Select
    'Zwrócenie posortowanej tablicy
    BubbleSortByInsertionPoint_MText = ArrayToBeSorted
End Function


Private Function GetTagListAndSelect(blockRefArray() As AcadBlockReference) As Integer
    Dim promptString As String                                      'Do przechowania monitu u¿ytkownika przy wybieraniu numeru tagu
    Dim tags As Variant                                             'Do przechowywania tablicy tagów
    Dim tagNumber As Integer                                        'Do przechowywania wybranego numeru tagów
    Dim i As Integer                                                'Jako zmienna iteruj¹ca
    
    On Error Resume Next
    
    Dim attRef As AcadAttributeReference
    promptString = "Podaj numer tagu: " & vbLf
    tags = blockRefArray(0).GetAttributes                           'Pobranie tablicy tagów z pierwszego elementu tablicy
    If Not Err Then
        For i = LBound(tags) To UBound(tags)
            Set attRef = tags(i)
            promptString = promptString & "[" & i & "]-" & attRef.tagString & vbLf
        Next i
    End If
    MsgBox promptString
    
    GetTagListAndSelect = ThisDrawing.Utility.GetInteger("Podaj numer tagu: ")
End Function

Private Function GetTagNumberArray(blockRef As AcadBlockReference, attNum As Integer) As Integer()
    Dim attRef As AcadAttributeReference
    Dim attTagString As String
    Dim tempArray() As Integer
    Dim count As Integer, i As Integer
    count = 0
    If blockRef.HasAttributes Then
        attTagString = blockRef.GetAttributes(attNum).tagString
        For i = LBound(blockRef.GetAttributes) To UBound(blockRef.GetAttributes)
            Set attRef = blockRef.GetAttributes(i)
            If attRef.tagString = attTagString Then
                ReDim Preserve tempArray(count)
                tempArray(count) = i
                count = count + 1
            End If
        Next i
    End If
    GetTagNumberArray = tempArray
End Function

Sub GetCaptionsAboveDWGS()
    '*****  DEKLARACJA ZMIENNYCH *****
    Dim selBlock As AcadBlockReference
    
    Dim blockName As String
    
    Dim selectionSet As AcadSelectionSet
    Dim blockRefObj As AcadBlockReference
    
    'Do przechowywania atrybutów
    Dim varAttributes As Variant

    'Do przechowywania opisów
    Dim textArray() As String
    'Tablica dynamiczna
    ReDim textArray(0)
    
    Dim blockCount As Integer, i As Integer
    Dim txtCount As Integer
    

    i = 0
    blockCount = 0
    txtCount = 0
    
    '*****  WYBÓR BLOKU *****
    Dim blockObj As AcadBlockReference
    Set blockObj = PickBlockReference
    
    If blockObj Is Nothing Then
        MsgBox "Nie wybrano bloku!", vbCritical
        Exit Sub
    Else
        blockName = blockObj.EffectiveName
    End If
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(1) As Variant, filterData(1) As Variant
    
    'Filtrowanie po typie obiektu - AcadBlockReference
    filterType(0) = acFilterByObjectType: filterData(0) = acBlockRef
    'Filtrowanie po nazwie obiektu - BlockName
    filterType(1) = acFilterByObjectName: filterData(1) = blockName
    'Wybor selection setem z filtrem
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    '*****  DODANIE DO TABLICY *****
    'Do przechowywania referencji bloków
    Dim blockRefArray() As AcadBlockReference
    ReDim blockRefArray(0)
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = GetItemsFromSelectionSet(selectionSet)

    '*****  SORTOWANIE BLOKÓW *****

    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)
    
    
    '*****  WYBÓR NUMERU ATRYBUTU *****
    'Wybranie numeru atrybutu z bloku do numeracji
    Dim tagNumber As Integer
    tagNumber = GetTagListAndSelect(blockRefArray)
    
    Dim attRef As AcadAttributeReference
    
    '*****  UTWORZENIE TABLICY Z TEKSTAMI *****
    
    For i = LBound(blockRefArray) To UBound(blockRefArray)
        Set blockObj = blockRefArray(i)
        ReDim Preserve textArray(txtCount)
        If blockObj.HasAttributes Then
            varAttributes = blockRefArray(i).GetAttributes
            If tagNumber <= UBound(varAttributes) Then
                Set attRef = varAttributes(tagNumber)
                textArray(txtCount) = attRef.textString
            Else
                textArray(txtCount) = ""
            End If
        Else
        textArray(txtCount) = ""
        End If
        txtCount = txtCount + 1
    Next i
    
    '*****  POBRANIE PUNKTU WSTAWIENIA I INKREMENTACJI *****
    Dim textObj As AcadMText
    Dim textString As Variant
    Dim insertionPoint As Variant
    
    Dim increment As Integer
    Dim textHeight As Double
    Dim textWidth As Double
    
    insertionPoint = ThisDrawing.Utility.GetPoint(, "Podaj punkt wstawienia: ")
    increment = ThisDrawing.Utility.GetDistance(, "Podaj odstêp: ")
    textWidth = ThisDrawing.Utility.GetDistance(, "Podaj szerokoœæ tekstu: ")
    textHeight = ThisDrawing.Utility.GetInteger("Podaj wysokoœæ tekstu: ")
    
    Dim j As Integer
    j = 0
    
    '*****  DODANIE MTEXTÓW DO RYSUNKU *****


    Dim captionLayer As AcadLayer
    Set captionLayer = GetLayer("Captions Above", False)
    
    For Each textString In textArray
        Set textObj = ThisDrawing.ModelSpace.AddMText(insertionPoint, textWidth, VBA.UCase(textString))
        
        textObj.height = textHeight
        textObj.AttachmentPoint = acAttachmentPointMiddleCenter
        textObj.insertionPoint = insertionPoint
        textObj.Layer = captionLayer.Name
        textObj.Update
        
        insertionPoint(0) = insertionPoint(0) + increment
    Next textString
    
End Sub

Sub ApplyFormat()

    Dim blockName As String
    
    Dim selectionSet As AcadSelectionSet                'Zmienna do zestawu wybranych elementów
    Dim blockRefObj As AcadBlockReference
    Dim ent As AcadEntity                               'Do sortowania po zestawie wybranych elementów
    Dim varAttributes As Variant                        'Do przechowywania atrybutów
    Dim blockRefArray() As AcadBlockReference           'Do przechowywania referencji bloków

    ReDim blockRefArray(0)                              'Tablica dynamiczna referencji bloków
    
    On Error Resume Next
    Set selectionSet = ThisDrawing.SelectionSets.Item("SS1")
    
    If Err Then
        Set selectionSet = ThisDrawing.SelectionSets.Add("SS1")
        Err.Clear
    End If
    
    selectionSet.Clear
    selectionSet.SelectOnScreen
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = AddToBlockRefArray(blockRefArray, selectionSet)
    
    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)
    
    'Wybranie numeru atrybutu z bloku do numeracji
    Dim tagNumber As Integer                            'Do przechowywania numeru tagu, który bêdzie numerowany
    tagNumber = GetTagListAndSelect(blockRefArray)
    
    'Pobranie tablicy tagów o tej samej nazwie
    Dim tagArray() As Integer
    tagArray = GetTagNumberArray(blockRefArray(0), tagNumber)
    
    'Okreœlenie wartoœci do procedury numeracji
    Dim startingValue As Integer, incrementValue As Integer
    ThisDrawing.Utility.InitializeUserInput 1 + 2
    
    Dim attRef As AcadAttributeReference, i As Integer
    Dim tagVarNumber As Variant
    'Zmiana atrybutu w bloku
    For i = LBound(blockRefArray) To UBound(blockRefArray)
    
        varAttributes = blockRefArray(i).GetAttributes                             'Pobranie zestawu atrybutów dla bloku[i]
        For Each tagVarNumber In tagArray
            Set attRef = varAttributes(tagVarNumber)
            attRef.textString = Format(CStr(attRef.textString), "00")                            'Przypisanie do atrybutu nowej wartoœci
        Next tagVarNumber
                                    
    Next i
    
    'Wiadomoœæ koñcowa
    Dim promptMsg As String
    promptMsg = "Zakoñczono numeracjê." & vbNewLine
End Sub

Sub ApplyConstantTag()

    Dim blockName As String
    
    Dim selectionSet As AcadSelectionSet                'Zmienna do zestawu wybranych elementów
    Dim blockRefObj As AcadBlockReference
    Dim ent As AcadEntity                               'Do sortowania po zestawie wybranych elementów
    Dim varAttributes As Variant                        'Do przechowywania atrybutów
    Dim blockRefArray() As AcadBlockReference           'Do przechowywania referencji bloków

    ReDim blockRefArray(0)                              'Tablica dynamiczna referencji bloków
    
    On Error Resume Next
    Set selectionSet = ThisDrawing.SelectionSets.Item("SS1")
    
    If Err Then
        Set selectionSet = ThisDrawing.SelectionSets.Add("SS1")
        Err.Clear
    End If
    
    
    Dim blockObj As AcadBlockReference, pickPts As Variant
    ThisDrawing.Utility.GetEntity blockObj, pickPts, "Podaj blok: "
    
    blockName = blockObj.EffectiveName
    
    selectionSet.Clear
    selectionSet.SelectOnScreen
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = AddToBlockRefArray(selectionSet, blockName)
    
    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)
    
    'Wybranie numeru atrybutu z bloku do numeracji
    Dim tagNumber As Integer                            'Do przechowywania numeru tagu, który bêdzie numerowany
    tagNumber = GetTagListAndSelect(blockRefArray)
    
    'Pobranie tablicy tagów o tej samej nazwie
    Dim tagArray() As Integer
    tagArray = GetTagNumberArray(blockRefArray(0), tagNumber)
    
    Dim constantVal As String
    constantVal = ThisDrawing.Utility.GetString(1, "Podaj sta³¹: ")
    Dim attRef As AcadAttributeReference, i As Integer
    Dim tagVarNumber As Variant
    'Zmiana atrybutu w bloku
    For i = LBound(blockRefArray) To UBound(blockRefArray)
    
        varAttributes = blockRefArray(i).GetAttributes                             'Pobranie zestawu atrybutów dla bloku[i]
        For Each tagVarNumber In tagArray
            Set attRef = varAttributes(tagVarNumber)
            attRef.textString = constantVal                                        'Przypisanie do atrybutu nowej wartoœci
        Next tagVarNumber
                                    
    Next i
    
    'Wiadomoœæ koñcowa
    Dim promptMsg As String
    promptMsg = "Zakoñczono numeracjê." & vbNewLine
End Sub

Sub deleteLayouts()
    Dim lastLayoutNr As Long
    
    ThisDrawing.Utility.InitializeUserInput 1 + 2
    lastLayoutNr = ThisDrawing.Utility.GetInteger("Podaj numer ostatniego layoutu: ")
    
    Dim i As Long
    Dim Layout As AcadLayout
    For Each Layout In ThisDrawing.Layouts
        If Layout.Name <> "Model" Then
            If CInt(Layout.Name) > lastLayoutNr Then
                Layout.Delete
            End If
        End If
    Next Layout
End Sub

Sub BreakLineInParts()
    Dim startPt As Variant
    Dim obj As AcadObject
    Dim pickedPts As Variant
    
    ThisDrawing.Utility.GetEntity obj, pickedPts, "Podaj liniê: "
    
    If Not TypeOf obj Is acadLine And Not TypeOf obj Is AcadLWPolyline Then
        MsgBox "B³êdny wybór!"
        Exit Sub
    End If
    
    Dim dist As Double
    dist = ThisDrawing.Utility.GetDistance(prompt:="Podaj odstêp: ")
    
    Dim lineCount As Long
    lineCount = ThisDrawing.Utility.GetInteger(prompt:="Podaj iloœæ linii: ")
    
    
    Dim acadLine As acadLine
    Dim angle As Double
    Set acadLine = obj
    
    startPt = acadLine.startPoint
    angle = acadLine.angle
    
    Dim aLayerName As String
    aLayerName = acadLine.Layer
    Dim aLayer As AcadLayer
    Set aLayer = ThisDrawing.Layers(aLayerName)
    
    Dim acadLineWeight As AcLineWeight
    acadLineWeight = aLayer.Lineweight
    
    Dim acadLineType As String
    acadLineType = acadLine.Linetype
    
    Dim acadLineTypeScale As Double
    acadLineTypeScale = acadLine.LinetypeScale
    
    acadLine.Delete
    
    Dim firstPt As Variant
    Dim nextPt As Variant
    
    
    firstPt = startPt
    
    Dim i As Integer
    For i = 1 To lineCount
       nextPt = ThisDrawing.Utility.PolarPoint(firstPt, angle, dist)
       Set acadLine = ThisDrawing.ModelSpace.AddLine(firstPt, nextPt)
       acadLine.Layer = aLayerName
       
       acadLine.Lineweight = acadLineWeight
       acadLine.Linetype = acadLineType
       acadLine.LinetypeScale = acadLineTypeScale
       
       firstPt = nextPt
    Next i
    
    
End Sub

Sub LoopThroughAtts()
    Dim textObj As aCadText
    Dim aEnt As AcadEntity
    Dim sset As AcadSelectionSet
    
    Dim i As Integer
    Dim attNum As Integer
    Dim texts() As aCadText
    
    Dim attRef As AcadAttributeReference
    Dim attVals As Variant
    Dim blockRef As AcadBlockReference
    
    ThisDrawing.Utility.GetEntity blockRef, Array(1#, 1#, 1#), "Podaj blok: "
    If blockRef.HasAttributes Then
        attNum = UBound(blockRef.GetAttributes) - 1
    End If
    
    Do
        ReDim Preserve texts(i)
        ThisDrawing.Utility.GetEntity aEnt, Array(1#, 1#, 1#), "Podaj text: "
        Set textObj = aEnt
        Set texts(i) = textObj
        i = i + 1
    Loop While i <= attNum
    

    
    If blockRef.HasAttributes Then
        attVals = blockRef.GetAttributes
        For i = LBound(texts) To UBound(texts)
            Set attRef = attVals(i)
            attRef.textString = texts(i).textString
        Next i
    End If
    
End Sub


Sub CapsAtt()
    Dim acadEnt As AcadEntity
    Dim attArray As Variant
    Dim attRef As AcadAttributeReference
    
    ThisDrawing.Utility.GetEntity acadEnt, Array(1#, 1#, 1#), "Podaj blok: "
    
    If Not TypeOf acadEnt Is AcadBlockReference Then
        Exit Sub
    End If
    
    Dim blockName As String
    Dim blockRef As AcadBlockReference
    blockName = acadEnt.Name
    
    Set blockRef = acadEnt
    
    If blockRef.HasAttributes Then
        attArray = blockRef.GetAttributes
        Set attRef = attArray(0)
        attRef.textString = StrConv(attRef.textString, vbUpperCase)
    End If
     
     MsgBox attRef.textString
    
    
End Sub

Sub ListLayers()
    Dim aEnt As AcadEntity
    Dim sset As AcadSelectionSet
    Set sset = GetSelectionSet
    
    sset.SelectOnScreen
    
    For Each aEnt In sset
        Debug.Print aEnt.ObjectName, aEnt.Layer
    Next aEnt
End Sub

Sub GetTextFromExcel()
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    
    On Error Resume Next
    Set xlApp = GetObject("Excel.Application")
    If Err Then
        Set xlApp = CreateObject("Excel.Application")
        Err.Clear
    End If
    
    Dim path As String
    path = xlApp.GetSaveAsFilename
    Set xlWb = xlApp.Workbooks.Open(FileName:=path)
    xlApp.Visible = True
    
    Dim xlRange As Excel.Range
    Set xlRange = xlApp.InputBox(prompt:="Wska¿ zakres do wklejenia do AutoCADa", Type:=8)
    
    Dim textArray() As String
    Dim i As Long: i = 0
    Dim xlCell As Excel.Range
    For Each xlCell In xlRange
        ReDim Preserve textArray(i)
        textArray(i) = xlCell.value
        i = i + 1
    Next xlCell
    
    xlWb.Close SaveChanges:=False
    xlApp.Quit
    
    Dim insPt As Variant
    insPt = ThisDrawing.Utility.GetPoint(prompt:="Podaj punkt wstawienia: ")
    
    Dim increment As Double
    increment = ThisDrawing.Utility.GetDistance(prompt:="Podaj odstêp: ")
    
    Dim txtObj As AcadMText
    Dim txt As Variant
    
    i = 0
    For Each txt In textArray
        Set txtObj = ThisDrawing.ModelSpace.AddMText(insPt, 200, txt)
        insPt(1) = insPt(1) - increment
        i = i + 1
    Next txt
    
End Sub

Private Function GetSelectionSet()
    Dim sset As AcadSelectionSet
    'Jeœli b³¹d, to ju¿ istnieje
    On Error Resume Next
    Set sset = ThisDrawing.SelectionSets.Add("SS1")
    If Err Then
        Err.Clear
        Set sset = ThisDrawing.SelectionSets.Item("SS1")
    End If
    
    Set GetSelectionSet = sset
End Function


Private Function GetFilteredSelectionSet(filterType As Variant, filterData As Variant) As AcadSelectionSet
    Dim selSet As AcadSelectionSet
    Set selSet = GetSelectionSet
    selSet.Clear
    
    On Error Resume Next
    Debug.Print "SelectionSetFilters()" & vbNewLine
    
    Dim filterT() As Integer
    ReDim filterT(UBound(filterType) - LBound(filterType))
    
    Dim i As Integer
    For i = LBound(filterType) To UBound(filterType)
        filterT(i) = CInt(filterType(i))
    Next i
    
    selSet.SelectOnScreen filterT, filterData
    Set GetFilteredSelectionSet = selSet
End Function

Private Function GetItemsFromSelectionSet(selectionSet As AcadSelectionSet) As Variant
    Dim ent As AcadEntity
    Dim blockRefArray() As AcadEntity
    Dim blockCount As Long: blockCount = 0
    For Each ent In selectionSet
        'Dynamiczna alokacja wymiarów tablicy
        ReDim Preserve blockRefArray(blockCount)
        Set blockRefArray(blockCount) = ent
        blockCount = blockCount + 1
    Next ent
    GetItemsFromSelectionSet = blockRefArray
End Function

Private Function PickBlockReference() As AcadBlockReference
    Dim aEnt As AcadEntity: Set aEnt = Nothing
    Dim blockRef As AcadBlockReference: Set blockRef = Nothing
    Set PickBlockReference = Nothing
    'Ustaw K jako s³owo kluczowe
    ThisDrawing.Utility.InitializeUserInput 128, "K"
    
    On Error Resume Next
    'Warunek wyjscia z petli - wybor slowa kluczowego K lub wybor bloku
    Do
        'Wybor bloku
        ThisDrawing.Utility.GetEntity aEnt, Array(1#, 1#, 1#), "Wybierz blok lub [Koniec]: "
        'Wybor akcji w zaleznosci od numeru bledu
        Select Case (Err.Number)
            Case acErrorBlankSpaceClicked
                Err.Clear
            Case acErrorKeywordSelected
                Dim userInput As String
                userInput = ThisDrawing.Utility.GetInput
                If userInput = "K" Then
                    Exit Function
                End If
                Err.Clear
        End Select
        
        If Not aEnt Is Nothing Then
            If TypeOf aEnt Is AcadBlockReference Then
                Set blockRef = aEnt
                Exit Do
            End If
        End If
        
    Loop While (True)
    
    Set PickBlockReference = blockRef

End Function

Sub DeleteSpacesInMText()
    Dim sset As AcadSelectionSet
    Dim filterType(3) As Variant, filterData(3) As Variant
    
    filterType(0) = -4
    filterData(0) = "<or"
    filterType(1) = 0
    filterData(1) = "TEXT"
    filterType(2) = 0
    filterData(2) = "MTEXT"
    filterType(3) = -4
    filterData(3) = "or>"
    
    Set sset = GetFilteredSelectionSet(filterType, filterData)
    
    Dim mText As AcadMText, txt As aCadText
    Dim text As String
    Dim aEnt As AcadEntity
    
    For Each aEnt In sset
        Select Case (True)
         Case TypeOf aEnt Is AcadMText
            Set mText = aEnt
            text = mText.textString
            text = VBA.Replace(text, " ", "")
            mText.textString = text
         Case TypeOf aEnt Is aCadText
            Set txt = aEnt
            text = txt.textString
            text = VBA.Replace(text, " ", "")
            txt.textString = text
        End Select
    Next aEnt
    Set sset = Nothing
End Sub

Sub MTextsIntoTags()

    '*****  DEKLARACJA ZMIENNYCH *****
    Dim selBlock As AcadBlockReference
    Dim blockName As String

    Dim blockRefObj As AcadBlockReference

    'Do przechowywania opisów
    Dim textArray() As String
    'Tablica dynamiczna
    ReDim textArray(0)
    
    Dim blockCount As Integer, i As Integer
    Dim txtCount As Integer
    
    '*****  CZÊŒÆ ZWI¥ZANA Z TEKSTAMI *****
    MsgBox "Podaj teksty: ", vbInformation
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(3) As Variant, filterData(3) As Variant
    '
    filterType(0) = acFilterByLogicalExp: filterData(0) = "<or"
    'Filtrowanie po typie obiektu - AcadText
    filterType(1) = acFilterByObjectType: filterData(1) = acText
    'Filtrowanie po typie obiektu - AcadMText
    filterType(2) = acFilterByObjectType: filterData(2) = acMText
    '
    filterType(3) = acFilterByLogicalExp: filterData(3) = "or>"
    

    'Wybor selection setem z filtrem
    Dim selectionSet As AcadSelectionSet
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    Dim mTextObjArray() As AcadMText
    ReDim mTextObjArray(0)
    
    'Pobierz mTexty z selection setu
    mTextObjArray = GetItemsFromSelectionSet(selectionSet)
    'Sortuj mTexty po punkcie wstawienia
    mTextObjArray = BubbleSortByInsertionPoint_MText(mTextObjArray, acAscending)

    '*****  CZÊŒÆ ZWI¥ZANA Z BLOKAMI *****
    MsgBox "Podaj bloki: ", vbInformation
    

    '*****  DODANIE DO TABLICY *****
    'Do przechowywania referencji bloków
    Dim blockRefArray() As AcadBlockReference
    ReDim blockRefArray(0)
    
    i = 0
    blockCount = 0
    txtCount = 0
    
    '*****  WYBÓR BLOKU *****
    Dim blockObj As AcadBlockReference
    Set blockObj = PickBlockReference
    
    If blockObj Is Nothing Then
        MsgBox "Nie wybrano bloku!", vbCritical
        Exit Sub
    Else
        blockName = blockObj.EffectiveName
    End If
    
    '*****  CZÊŒÆ ZWI¥ZANA Z BLOKAMI *****
    
    Dim filterData_Block(1)
    Dim filterType_Block(1)
    
    '*****  FILTROWANIE *****
    'Filtrowanie po typie obiektu - AcadBlockReference
    filterType_Block(0) = acFilterByObjectType: filterData_Block(0) = acBlockRef
    'Filtrowanie po nazwie obiektu - BlockName
    filterType_Block(1) = acFilterByObjectName: filterData_Block(1) = blockName
    'Wybor selection setem z filtrem
    Set selectionSet = GetFilteredSelectionSet(filterType_Block, filterData_Block)
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = GetItemsFromSelectionSet(selectionSet)

    '*****  SORTOWANIE BLOKÓW *****

    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)
    
    '*****  WYBÓR NUMERU ATRYBUTU *****
    'Wybranie numeru atrybutu z bloku do numeracji
    Dim tagNumber As Integer
    tagNumber = GetTagListAndSelect(blockRefArray)
    
    Dim attRef As AcadAttributeReference
End Sub

Sub GetTextToExcel()
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    
    On Error Resume Next
    Set xlApp = GetObject("Excel.Application")
    If Err Then
        Set xlApp = CreateObject("Excel.Application")
        Err.Clear
    End If
    On Error GoTo 0

    Set xlWb = xlApp.Workbooks.Add
    
    '*****  CZÊŒÆ ZWI¥ZANA Z TEKSTAMI *****
    MsgBox "Podaj teksty: ", vbInformation
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(3) As Variant, filterData(3) As Variant
    
    filterType(0) = acFilterByLogicalExp: filterData(0) = "<or"
    'Filtrowanie po typie obiektu - AcadText
    filterType(1) = acFilterByObjectType: filterData(1) = acText
    'Filtrowanie po typie obiektu - AcadMText
    filterType(2) = acFilterByObjectType: filterData(2) = acMText
    '
    filterType(3) = acFilterByLogicalExp: filterData(3) = "or>"

    'Wybor selection setem z filtrem
    Dim selectionSet As AcadSelectionSet
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    Dim mTextObjArray() As AcadMText
    ReDim mTextObjArray(0)
    
    'Pobierz mTexty z selection setu
    mTextObjArray = GetItemsFromSelectionSet(selectionSet)
    'Sortuj mTexty po punkcie wstawienia
    mTextObjArray = BubbleSortByInsertionPoint_MText(ArrayToBeSorted:=mTextObjArray, _
                                              SortOrder:=acDescending, SortByY:=True)
                                              
    Dim xlRange As Excel.Range
    Set xlRange = xlWb.Worksheets(1).Range("A1")
    
    Dim i As Long
    For i = LBound(mTextObjArray) To UBound(mTextObjArray)
        xlRange.value = mTextObjArray(i).textString
        Set xlRange = xlRange.Offset(1)
    Next i
    
    xlApp.Visible = True
End Sub

Sub AttDefToMText()
    Dim attDef As AcadAttribute
    
    
    Dim sset As AcadSelectionSet
    Dim filterType(0) As Variant, filterData(0) As Variant
    filterType(0) = acFilterByObjectType: filterData(0) = acAttributeDef
    
    Set sset = GetFilteredSelectionSet(filterType, filterData)
    
    Dim insPt As Variant, tagString As String
    Dim lowerLeft As Variant, upperRight As Variant
    Dim width As Double
    
    Dim newText As AcadMText
    
    For Each attDef In sset
        insPt = attDef.insertionPoint
        tagString = attDef.tagString
        attDef.GetBoundingBox lowerLeft, upperRight
        width = (upperRight(0) - lowerLeft(0)) * 1.2
        Set newText = ThisDrawing.ModelSpace.AddMText(insPt, width, tagString)
        
        newText.AttachmentPoint = acAttachmentPointBottomLeft
        newText.insertionPoint = insPt
        newText.Update
        
        attDef.Delete
    Next attDef
End Sub

Function GetLayer(layerName As String, Optional printable As Boolean = False) As AcadLayer
    On Error Resume Next
    Set GetLayer = ThisDrawing.Layers(layerName)
    If Err Then
        Set GetLayer = ThisDrawing.Layers.Add(layerName)
        GetLayer.Plottable = printable
        Err.Clear
    End If
    On Error GoTo 0
End Function


Sub AddRevCloud()

    Dim sumInfo As AcadSummaryInfo
    Set sumInfo = ThisDrawing.SummaryInfo
    
    Dim rev As String
    
    On Error Resume Next
    sumInfo.GetCustomByKey "Rev", rev
    
    If Err Then
        Err.Clear
        rev = ThisDrawing.Utility.GetString(0, "Podaj numer rewizji: ")
        sumInfo.AddCustomInfo "Rev", rev
    End If
    
    Dim revTriangle As AcadBlock
    On Error Resume Next
    Set revTriangle = ThisDrawing.Blocks.Item("Rev_Triangle")
    If Err Then
        MsgBox "Zdefiniuj blok o nazwie Rev_Triangle"
        Exit Sub
    End If
    
    Dim blockCountBefore As Long, blockCountAfter As Long
    'Numer ostatniego elementu przed dodaniem linii
    blockCountBefore = ThisDrawing.ModelSpace.count
    
    ThisDrawing.SendCommand "revcloud" & vbCr
    
    Dim insPt As Variant, revTriangleBlockRef As AcadBlockReference
    insPt = ThisDrawing.Utility.GetPoint(prompt:="Podaj punkt wstawienia trójk¹ta rewizyjnego: ")
    Set revTriangleBlockRef = ThisDrawing.ModelSpace.InsertBlock(insPt, revTriangle.Name, 1, 1, 1, 0)
    
    
    Dim attRef As AcadAttributeReference, i As Integer
    Dim tagVarNumber As Variant, varAttributes As Variant
    tagVarNumber = 0

    varAttributes = revTriangleBlockRef.GetAttributes                            'Pobranie zestawu atrybutów dla bloku[i]
    
    Set attRef = varAttributes(tagVarNumber)
    attRef.textString = rev                                      'Przypisanie do atrybutu nowej wartoœci

    'Numer ostatniego elementu po dodaniu linii
    blockCountAfter = ThisDrawing.ModelSpace.count
    
    If blockCountAfter - blockCountBefore = 0 Then
        ThisDrawing.Utility.prompt ("Nie dodano ¿adnego elementu.") & vbCr
        Exit Sub
    End If
    
    Dim aEnt As AcadEntity
    Dim revLayer As AcadLayer
    Set revLayer = GetLayer("Revision " & rev, True)
    
    For i = blockCountBefore To blockCountAfter
        Set aEnt = ThisDrawing.ModelSpace.Item(i)
        aEnt.Layer = revLayer.Name
        aEnt.color = acRed
    Next i
    
End Sub

Sub AddRevCloudForCheckCopy()

    Dim blockCountBefore As Long, blockCountAfter As Long
    'Numer ostatniego elementu przed dodaniem linii
    blockCountBefore = ThisDrawing.ModelSpace.count
    
    ThisDrawing.SendCommand "revcloud" & vbCr
    
    'Numer ostatniego elementu po dodaniu linii
    blockCountAfter = ThisDrawing.ModelSpace.count
    
    If blockCountAfter - blockCountBefore = 0 Then
        ThisDrawing.Utility.prompt ("Nie dodano ¿adnego elementu.") & vbCr
        Exit Sub
    End If
    
    Dim aEnt As AcadEntity
    'Odwo³anie siê do nowododanego elementu - ostatniego na liœcie itemów w rysunku
    Set aEnt = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.count - 1)
    
    Dim revLayer As AcadLayer
    Set revLayer = GetLayer("Check Copy", False)
    'Ustawienie odpowiedniej warstwy i koloru
    aEnt.Layer = revLayer.Name
    aEnt.color = acGreen
End Sub

Sub AddAuxLine()
    Dim blockCountBefore As Long, blockCountAfter As Long
    'Numer ostatniego elementu przed dodaniem linii
    blockCountBefore = ThisDrawing.ModelSpace.count
    
    ThisDrawing.SendCommand "line" & vbCr
    
    'Numer ostatniego elementu po dodaniu linii
    blockCountAfter = ThisDrawing.ModelSpace.count - 1
    
    Dim linetypeName As String
    linetypeName = "ACAD_ISO02W100"
    Call LoadLinetype(linetypeName)
    
    Dim auxLayer As AcadLayer
    Set auxLayer = GetLayer("Auxiliary Lines", False)
    
    Dim aEnt As AcadEntity
    Dim i As Long
    For i = blockCountBefore To blockCountAfter
        Set aEnt = ThisDrawing.ModelSpace.Item(i)
        aEnt.Layer = auxLayer.Name
        aEnt.color = acRed
        aEnt.Linetype = linetypeName
        aEnt.LinetypeScale = 20#
    Next i
End Sub

Sub AddAnnotations()

    Dim blockCountBefore As Long, blockCountAfter As Long
    'Numer ostatniego elementu przed dodaniem linii
    blockCountBefore = ThisDrawing.ModelSpace.count

    ThisDrawing.Utility.InitializeUserInput 128, "TEXT LINIA PROSTOK¥T REVCLOUD MLEADER KONIEC"
    Dim userInput As String
    
    
    On Error Resume Next
    Do
    userInput = ThisDrawing.Utility.GetKeyword("[TEXT/LINIA/PROSTOK¥T/REVCLOUD/MLEADER/KONIEC]")
    
        Select Case (userInput)
            Case "TEXT"
                ThisDrawing.SendCommand "mtext" & vbCr
            Case "LINIA"
                ThisDrawing.SendCommand "line" & vbCr
            Case "PROSTOK¥T"
                ThisDrawing.SendCommand "rec" & vbCr
            Case "REVCLOUD"
                ThisDrawing.SendCommand "revcloud" & vbCr
            Case "MLEADER"
                ThisDrawing.SendCommand "mleader" & vbCr
            Case "KONIEC"
                Exit Do
        End Select
        
    Loop While (True)
    'Numer ostatniego elementu po dodaniu linii
    blockCountAfter = ThisDrawing.ModelSpace.count - 1
    
    Dim annotationLayer As AcadLayer
    Set annotationLayer = GetLayer("Annotations", False)
    
    Dim aEnt As AcadEntity
    Dim i As Long
    For i = blockCountBefore To blockCountAfter
        Set aEnt = ThisDrawing.ModelSpace.Item(i)
        aEnt.Layer = annotationLayer.Name
        aEnt.color = acYellow
    Next i
    
    On Error GoTo 0
End Sub

Sub LoadLinetype(linetypeName As String)
    On Error Resume Next
    ThisDrawing.Linetypes.Load linetypeName, "acadiso.lin"
    On Error GoTo 0
End Sub

Sub AddDivisionRectangle()
    Dim blockCountBefore As Long, blockCountAfter As Long
    'Numer ostatniego elementu przed dodaniem linii
    blockCountBefore = ThisDrawing.ModelSpace.count
    
    ThisDrawing.SendCommand "rec" & vbCr
    
    'Numer ostatniego elementu po dodaniu linii
    blockCountAfter = ThisDrawing.ModelSpace.count
    
    If blockCountAfter - blockCountBefore = 0 Then
        ThisDrawing.Utility.prompt ("Nie dodano ¿adnego elementu.") & vbCr
        Exit Sub
    End If
    
    Dim linetypeName As String
    linetypeName = "ACAD_ISO02W100"
    Call LoadLinetype(linetypeName)
    
    Dim divisionLayer As AcadLayer
    Dim aEnt As AcadEntity
    Set divisionLayer = GetLayer("Division Boxes", False)
    Set aEnt = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.count - 1)
    aEnt.Layer = divisionLayer.Name
    aEnt.color = acYellow
    aEnt.Linetype = linetypeName
    aEnt.LinetypeScale = 20#
    
End Sub

Sub ExcelTableToAutoCad()

    Dim aTable As acadTable
    Dim acEnt As AcadEntity
    
    Set aTable = PickTable()

    If aTable Is Nothing Then
        Exit Sub
    End If
    
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    
    On Error Resume Next
    Set xlApp = GetObject("Excel.Application")
    If Err Then
        Set xlApp = CreateObject("Excel.Application")
        Err.Clear
    End If
    
    Dim path As String
    path = xlApp.GetSaveAsFilename
    Set xlWb = xlApp.Workbooks.Open(FileName:=path)
    xlApp.Visible = True
    
    Dim xlRange As Excel.Range, vals As Variant
    Set xlRange = xlApp.InputBox(prompt:="Wska¿ zakres do wklejenia do AutoCADa", Type:=8)
    vals = xlRange.value
    
    Dim rowNum As Long, colNum As Long
    rowNum = xlRange.Rows.count
    colNum = xlRange.Columns.count

    Dim i As Long, j As Long
    For i = 1 To rowNum
        For j = 1 To colNum
            aTable.SetCellValue i, j, vals(i, j)
        Next j
    Next i

     xlApp.Quit
End Sub

Function PickTable() As acadTable
    Dim aEnt As AcadEntity: Set aEnt = Nothing
    Set PickTable = Nothing
    'Ustaw K jako s³owo kluczowe
    
    ThisDrawing.Utility.InitializeUserInput 128, "K"
    
    On Error Resume Next
    'Warunek wyjscia z petli - wybor slowa kluczowego K lub wybor bloku
    Do
        'Wybor bloku
        ThisDrawing.Utility.GetEntity aEnt, Array(1#, 1#, 1#), "Wybierz tabelê lub [Koniec]: "
        'Wybor akcji w zaleznosci od numeru bledu
        Select Case (Err.Number)
            Case acErrorBlankSpaceClicked
                Err.Clear
            Case acErrorKeywordSelected
                Dim userInput As String
                userInput = ThisDrawing.Utility.GetInput
                Err.Clear
                If userInput = "K" Then
                    Exit Function
                End If
        End Select
        
        If Not aEnt Is Nothing Then
            If TypeOf aEnt Is acadTable Then
                Set PickTable = aEnt
                Exit Do
            End If
        End If
    Loop While (True)
End Function

Sub BrowseDWG()
    Dim distance As Variant
    Dim lowerLeftPt As Variant, upperRightPt As Variant
    
    lowerLeftPt = ThisDrawing.Utility.GetPoint(, "Podaj lewy dolny róg")
    
    upperRightPt = ThisDrawing.Utility.GetPoint(, "Podaj prawy górny róg okna do zoomowania")
    
    distance = ThisDrawing.Utility.GetDistance(, "Podaj odleg³oœæ: ")
    ThisDrawing.Application.ZoomWindow lowerLeftPt, upperRightPt
    ThisDrawing.Utility.InitializeUserInput 2, "R M K"
    
    Dim mode As String
    mode = ThisDrawing.Utility.GetKeyword("Wybierz tryb: [Rosn¹co (w prawo)/Malej¹co (w lewo)/Koniec]")
    
    Dim userInput As String
    On Error Resume Next
    Do
    

        ThisDrawing.Utility.InitializeUserInput 2, "D K T"
        userInput = ThisDrawing.Utility.GetKeyword("Wybierz akcjê [Dalej/Koniec/Tryb]")
    
        Select Case (userInput)
            Case "D"
                Select Case (mode)
                    Case "R"
                        lowerLeftPt(0) = lowerLeftPt(0) + distance
                        upperRightPt(0) = upperRightPt(0) + distance
                    Case "M"
                        lowerLeftPt(0) = lowerLeftPt(0) - distance
                        upperRightPt(0) = upperRightPt(0) - distance
                End Select
            Case "K"
                Exit Sub
            Case "T"
                ThisDrawing.Utility.InitializeUserInput 2, "R M K"
                mode = ThisDrawing.Utility.GetKeyword("Wybierz tryb: [Rosn¹co (w prawo)/Malej¹co (w lewo)/Koniec]")
                Select Case (mode)
                    Case "R"
                        ThisDrawing.Utility.prompt "Zmieniono tryb na rosn¹co"
                    Case "M"
                        ThisDrawing.Utility.prompt "Zmieniono tryb na malej¹co"
                    Case "K"
                        Exit Sub
                End Select
         End Select
         
        ThisDrawing.Application.ZoomWindow lowerLeftPt, upperRightPt
    Loop While (True)
End Sub

Sub AddRevCloudForClarification()

    Dim blockCountBefore As Long, blockCountAfter As Long
    'Numer ostatniego elementu przed dodaniem linii
    blockCountBefore = ThisDrawing.ModelSpace.count
    
    ThisDrawing.SendCommand "revcloud" & vbCr
    
    'Numer ostatniego elementu po dodaniu linii
    blockCountAfter = ThisDrawing.ModelSpace.count
    
    If blockCountAfter - blockCountBefore = 0 Then
        ThisDrawing.Utility.prompt ("Nie dodano ¿adnego elementu.") & vbCr
        Exit Sub
    End If
    
    Dim aEnt As AcadEntity
    'Odwo³anie siê do nowododanego elementu - ostatniego na liœcie itemów w rysunku
    Set aEnt = ThisDrawing.ModelSpace.Item(ThisDrawing.ModelSpace.count - 1)
    
    Dim clarificationLayer As AcadLayer
    Set clarificationLayer = GetLayer("TO BE CLARIFIED", False)
    'Ustawienie odpowiedniej warstwy i koloru
    aEnt.Layer = clarificationLayer.Name
    aEnt.color = acMagenta
End Sub

Sub BrowseThroughQSelectSelectionSet()
    ThisDrawing.SendCommand "qselect" & vbCr
    
    Dim aEnt As AcadEntity
    Dim selSet As AcadSelectionSet
    Dim counter As Long
    
    Set selSet = ThisDrawing.ActiveSelectionSet
    
    Dim lowerLeftPt As Variant, upperRightPt As Variant
    Dim zoomInScale As Double, zoomOutScale As Double
    zoomInScale = 1.5
    zoomOutScale = 0.5
    
    
    ThisDrawing.Utility.InitializeUserInput 2, "Dalej Oddal Przybli¿ Koniec"
    
    counter = 0
    Set aEnt = selSet.Item(counter)
    aEnt.GetBoundingBox lowerLeftPt, upperRightPt
    ThisDrawing.Application.ZoomWindow lowerLeftPt, upperRightPt
    ThisDrawing.Application.ZoomScaled zoomOutScale, acZoomScaledRelative
    
    Do
        Dim userInput As String
        userInput = ThisDrawing.Utility.GetKeyword("[Dalej/Oddal/Przybli¿/Koniec]")
        Select Case (userInput)
            Case "Dalej"
                counter = counter + 1
                If counter < selSet.count Then
                    Set aEnt = selSet.Item(counter)
                    aEnt.GetBoundingBox lowerLeftPt, upperRightPt
                    ThisDrawing.Application.ZoomWindow lowerLeftPt, upperRightPt
                    ThisDrawing.Application.ZoomScaled zoomOutScale, acZoomScaledRelative
                Else
                    MsgBox "Osi¹gniêto ostatni element w zbiorze elementów. Koniec!"
                    counter = 1
                End If
                
            Case "Oddal"
                ThisDrawing.Application.ZoomScaled zoomOutScale, acZoomScaledRelative
            Case "Przybli¿"
                ThisDrawing.Application.ZoomScaled zoomInScale, acZoomScaledRelative
            Case "Koniec"
                Exit Sub
        End Select
    Loop While (True)
    
End Sub

Sub FindGivenValves()
    Dim aEnt As AcadEntity
    Dim blkRef As AcadBlockReference
    Dim tagArray As Variant
    Dim atts As Variant
    Dim tagAttRef As AcadAttributeReference
    tagArray = Array("...", "...", "...")     'tablica z tagami do znalezienia

    Dim counter As Long
    Dim tagString As String
    For Each aEnt In ThisDrawing.ModelSpace
        If TypeOf aEnt Is AcadBlockReference Then
            Set blkRef = aEnt
            If blkRef.EffectiveName = "..." Or blkRef.EffectiveName = "..." Then   'szukanie tylko w blokach o danej nazwie
                atts = blkRef.GetAttributes
                Set tagAttRef = atts(0)
                tagString = VBA.Replace(tagAttRef.textString, "\W0.8000;", "")
                If IsInArray(tagString, tagArray) Then
                    counter = counter + 1
                Else
                    blkRef.Delete
                End If
            End If
        End If
    Next aEnt
    MsgBox counter
End Sub

Function IsInArray(text As String, arr As Variant) As Boolean
    IsInArray = False
	
    If IsArray(arr) = False Then
        Exit Function
    End If
    
    Dim element As Variant
    For Each element In arr
        If (VBA.Trim(VBA.UCase(text))) = (VBA.Trim(VBA.UCase(element))) Then
            IsInArray = True
            Exit Function
        End If
    Next element
    
End Function

Sub ExportXAndYOfBLocks()

    '*****  DEKLARACJA ZMIENNYCH *****
    Dim blockName As String
    
    Dim selectionSet As AcadSelectionSet                'Zmienna do zestawu wybranych elementów
    Dim blockRefObj As AcadBlockReference
    Dim varAttributes As Variant                        'Do przechowywania atrybutów
    
    Dim blockRefArray() As AcadBlockReference           'Do przechowywania referencji bloków
    ReDim blockRefArray(0)                              'Tablica dynamiczna referencji bloków
    
    '*****  WYBÓR BLOKU *****
    
    Dim blockObj As AcadBlockReference
    Set blockObj = PickBlockReference
    
    If blockObj Is Nothing Then
        MsgBox "Nie wybrano bloku!", vbCritical
        Exit Sub
    Else
        blockName = blockObj.EffectiveName
    End If
    
    '*****  FILTROWANIE *****
    'Dane do filtru
    Dim filterType(0) As Variant, filterData(0) As Variant
    
    'Filtrowanie po typie obiektu - AcadBlockReference
    filterType(0) = acFilterByObjectType: filterData(0) = acBlockRef
    'Wybor selection setem z filtrem
    Set selectionSet = GetFilteredSelectionSet(filterType, filterData)
    
    '*****  DODANIE DO TABLICY *****
    
    'Przeiterowanie po selection secie i dodanie referencji bloków do tablicy
    blockRefArray = GetItemsFromSelectionSet(selectionSet)

    '*****  SORTOWANIE BLOKÓW *****

    'Algorytm sortowania b¹belkowego, sortujemy rosn¹co po pozycji X bloku
    blockRefArray = BubbleSortByInsertionPoint(blockRefArray, acAscending)


    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    
    On Error Resume Next
    Set xlApp = GetObject("Excel.Application")
    If Err Then
        Set xlApp = CreateObject("Excel.Application")
        Err.Clear
    End If
    On Error GoTo 0

    Set xlWb = xlApp.Workbooks.Add
    
    Dim xlRange As Excel.Range
    Set xlRange = xlWb.Worksheets(1).Range("A1")
    Dim i As Long
    
    Dim varAtts As Variant
    For i = LBound(blockRefArray) To UBound(blockRefArray)
    
        varAtts = blockRefArray(i).GetAttributes
        
        xlRange.value = varAtts(0).textString
        xlRange.Offset(0, 1).value = blockRefArray(i).insertionPoint(0)
        xlRange.Offset(0, 2).value = blockRefArray(i).insertionPoint(1)
        Set xlRange = xlRange.Offset(1, 0)
    Next i

    xlApp.Visible = True
End Sub
