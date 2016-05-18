Attribute VB_Name = "SAPMakro"
'++ EJ160204
'*--------------------------------------------------------------------------------*
'* Copyright © MAGNA STEYR Fahrzeugtechnik AG & Co KG.                            *
'* Alle Rechte vorbehalten                                                        *
'*--------------------------------------------------------------------------------*
'* Objektname:                                                                    *
'* Referenzierte Objekte:                                                         *
'* Beschreibung:                                                                  *
'*                                                                                *
'* Änderungshistorie                                                              *
'*--------------------------------------------------------------------------------*
'* Anforderer   Entwickler  Kürzel      Beschreibung                              *
'*--------------------------------------------------------------------------------*
'* Brandstätter Mundprecht  MH          Urversion                                 *
'* Brandstätter Edlinger    EJ160204    Excel-Planungsdatei                       *
'*                                      MUC_EI_Kleinprojekte.xlsx                 *
'*                                      wurde von 2010-2018 auf                   *
'*                                      2010-2026 erweitert.                      *
'*                                      Neuer Parameter LastPlanningYear variabel *
'* Bruckmeyer   Edlinger    EJ160216    FC-Monate sollen nur selektiv hochgeladen *
'*                                      werden, mit "No Upload" markierte nicht!  *
'* Kammann      Edlinger    EJ160301    Prüfg eingebaut, ob Projekt als Worksheet *
'*                                      in der Datei exitstiert.                  *
'* Bruckmeyer   Mundprecht  HM160413    Planwerte löschen vor Upload              *
'* Edlinger     Mundprecht  HM160518    Leere Jahre beim Löschen ignorieren       *
'*--------------------------------------------------------------------------------*
'++ EJ160204

' EJ160204: Konstanten für Zeilen und Spalten in allen Projekt-Sheets (inkl. Sheet "Vorlage"
Const c_DSC As Integer = 14     ' EJ160204: Date Start Column; FC-Spalte des erste Monats (Jänner)
'Const c_DEC As Integer = 228   '-- EJ160204: Es wird stattdessen ein Parameter (Sheet Parameter) für das letzte Planungsjahr eingeführt (aDEC)
Const c_ELC As Integer = 10     ' EJ160204: Element Column; Spalte J, "ext", "int",...
Const c_DR As Integer = 13      ' EJ160204: Date Row; Zeile mit der Datums-Angabe
Const c_NoUplR As Integer = 15  '++ EJ160216: In dieser Zeile werden FC-Spalten markiert, die nicht nach SAP geladen werden sollen.
Const c_HSR As Integer = 17     ' EJ160204: Hours Start Row
Const c_HER As Integer = 99     ' EJ160204: Hours End Row
Const c_CSR As Integer = 101    ' EJ160204: COS Start Row
Const c_CER As Integer = 199    ' EJ160204: COS End Row

Dim aDEC As Integer    ' EJ160204: Date End Column; FC-Spalte des letzten Monats (Dezember), um 8 Jahre = 192 Spalten mehr

Sub SAP_PS_COC_PostPrimCost()
On Error GoTo SAP_PS_COC_PostPrimCost_Err
Dim aPlanElemMapping As New PlanElemMapping
Dim aPlanCOSPrice As New PlanCOSPrice
Dim aPlanHours As New PlanHours
Dim aPlanCOS As New PlanCOS
'++ HM160413
Dim aClearCOS As New PlanCOS
Dim aClearCOSRec As Variant
Dim aClearData As Collection
Dim aClearMsg As MessageCollection
'++ HM160413
Dim aPlanCOSRec As Variant
Dim aCOArea As String
Dim aVers As String
Dim aCurt As String
Dim aProject As String
Dim aFWBS As String
Dim aPWS As Worksheet
Dim aWS As Worksheet
'++ EJ160204
Dim aLastPlanY As Integer
'++ EJ160204
Dim i As Integer
Dim aData As Collection
Dim aKey As String
Dim aSAP_PC_Plan As New SAP_PC_Plan
Dim aSAPFormat As New SAPFormat
Dim aValue As Double

    Dim aYear As Integer
    Dim aOldYear As Integer
    Dim aMonth As Integer
    Dim aDR As Range
    Dim aHR As Range
    Dim aFName As String
    Dim Index As Integer
    Dim j As Integer
    Dim aWBS As String
    Dim aCostElem As String
    Dim aHRName As String
    Dim aDRName As String
    Dim aRRName As String
    Dim aWSName As String
        
  Worksheets("Parameter").Activate
  Set aPWS = ActiveWorkbook.Worksheets("Parameter") ' EJ160204: Worksheet "Parameter" mit KoKrs, Planversion, Währungsart, Kontenliste und Liste der Fakt.-Elemente
  aCOArea = aPWS.Cells(2, 2)
  aVers = aPWS.Cells(3, 2)
  aCurt = aPWS.Cells(4, 2)
  aLastPlanY = aPWS.Cells(5, 2)   '++ EJ160204
  
'++ EJ160204
  If IsNull(aPWS.Cells(5, 1)) Or aPWS.Cells(5, 1) = "" Then
    MsgBox "Parameter <Last Planning Year> missing in line 5!", vbCritical + vbOKOnly
    Exit Sub
  End If
'++ EJ160204
  
'-- EJ160204
'  If IsNull(aCOArea) Or aCOArea = "" Or _
'     IsNull(aVers) Or aVers = "" Or _
'     IsNull(aCurt) Or aCurt = "" Then
'    MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
'    Exit Sub
'  End If
'-- EJ160204

'++ EJ160204
  If IsNull(aCOArea) Or aCOArea = "" Or _
     IsNull(aVers) Or aVers = "" Or _
     IsNull(aCurt) Or aCurt = "" Or _
     IsNull(aLastPlanY) Or aLastPlanY = 0 Then
    MsgBox "Please fill all obligatory fields in the parameter sheet!", vbCritical + vbOKOnly
    Exit Sub
  End If
  If aLastPlanY < 2018 Or aLastPlanY > 2026 Then
    MsgBox "Last Planning Year must be within 2018 and 2026!", vbCritical + vbOKOnly
    Exit Sub
  End If
  aDEC = c_DSC + (aLastPlanY - 2009) * 24 - 2 ' EJ160204: Berechnung Dezember-FC-Spalte des letzten Planungsjahres
  i = 2
  Do ' EJ160204: Für die Sheets aller Projekte der Projektliste (Sheet Parameter)
    aProject = CStr(aPWS.Cells(i, 5))
'++ EJ160301
    If Not WorksheetExists(aProject) Then
      MsgBox "Worksheet for Project " & aProject & " does not exist!", vbCritical + vbOKOnly
      Worksheets("Parameter").Activate
      Cells(i, 5).Activate
      Exit Sub
    End If
'++ EJ160301
    If Worksheets(aProject).Cells(c_DR - 1, aDEC + 1).Value <> "Gesamt" Then ' EJ160204: Prüfung, ob Sheet auch bis zum angegebenen Jahr ausgebaut ist.
      MsgBox "Last Planning Year for Project " & aProject & " not correct!", vbCritical + vbOKOnly
      Exit Sub
    End If
'++ EJ160216: Prüfung, ob die Zellen für "No Upload" auch gültige Werte enthalten.
    For j = c_DSC To aDEC Step 2
        If Worksheets(aProject).Cells(c_NoUplR, j).Value <> "" And UCase(Worksheets(aProject).Cells(c_NoUplR, j).Value) <> "NO UPLOAD" Then
          MsgBox "Wrong Entry in No Upload Cell, Row " & c_NoUplR & ", Col " & j & " for Project " & aProject & "!" & Chr(13) & "Valid is '' (empty) or 'No Upload'.", vbCritical + vbOKOnly
          Worksheets(aProject).Activate
          Cells(c_NoUplR, j).Activate
          Exit Sub
        End If
    Next j
'++ EJ160216
    i = i + 1
  Loop While CStr(aPWS.Cells(i, 5)) <> ""
'++ EJ160204

  aRet = SAPCheck() ' EJ160204: SAP Verbindung aufbauen
  If Not aRet Then
    MsgBox "Connection to SAP failed!", vbCritical + vbOKOnly
    Exit Sub
  End If
  
  Set aPlanElemMapping = getElemMapping()   ' EJ160204: Liest die Element / Account Mapping Tabelle ab Zeile 10 nach unten aus Sheet "Parameter" ein
  Set aPlanCOSPrice = getCOSPrices()        ' EJ160204: Liest die Zeile 4 aus Sheet "ACCT" für die Jahre aus Zeile 2 ein
  
  Application.Cursor = xlWait
  i = 2
  Do ' EJ160204: Für jedes Projekt/Fakt.-PSP der Tabelle in Sheet "Parameter"
    Set aPlanHours = New PlanHours
    Set aPlanCOS = New PlanCOS
    Set aData = New Collection
    aProject = CStr(aPWS.Cells(i, 5))                   ' EJ160204: Projekt aus Tabelle in Sheet "Parameter"
    aFWBS = CStr(aPWS.Cells(i, 6))                      ' EJ160204: Fakt.-PSP aus Tabelle in Sheet "Parameter"
    Set aWS = ActiveWorkbook.Worksheets(aProject)       ' EJ160204: Verweis auf Worksheet des Projekts
    Set aPlanHours = collectHours(aWS)                  ' EJ160204: Einlesen der Stunden für ein Projekt: Monatliche Werte "FC" jährlich aufsummiert für alle Stunden-Zeilen
    Set aPlanCOS = collectCOS(aWS, aPlanElemMapping)    ' EJ160204: Einlesen der COS für ein Projekt: Monatliche Werte "FC" jährlich aufsummiert für alle COS-Zeile
    Set aClearCOS = collectClearData(aWS, aPlanElemMapping) ''++ HM160413: Get all year/cost-elememt combinations to clear
    aPlanCOS.addOtherCOS aPlanHours, aPlanCOSPrice, aPlanElemMapping
    
'   we need plan records per year -> transform to collection of SAP_PC_Plan
    For Each aPlanCOSRec In aPlanCOS.aPlanCOS
        aKey = aPlanCOSRec.aYear.Value & "-" & aPlanCOSRec.aElem.Value
        aValue = CDbl(aPlanCOSRec.aValue.Value)
        If contains(aData, aKey, "obj") Then
            Set aSAP_PC_Plan = aData(aKey)
            aValue = aValue + CDbl(aSAP_PC_Plan.ValVar("VAR_VAL_PER12").Value)
            aSAP_PC_Plan.ValVar("VAR_VAL_PER12").Value = CStr(aValue)
        Else
            Set aSAP_PC_Plan = aSAP_PC_Plan.create(aCOArea, aPlanCOSRec.aYear.Value, "12", "12", aVers, aCurt)
            aSAP_PC_Plan.setOValues "WBS_ELEMENT", aFWBS
            aSAP_PC_Plan.CostElement = aSAPFormat.unpack(aPlanCOSRec.aElem.Value, "10")
            aSAP_PC_Plan.addVValue "VAR_VAL_PER12", CStr(Format$(aValue, "0.00"))
            aData.add aSAP_PC_Plan, aKey
        End If
    Next
'++ HM160413
'   build aClearData in SAP_PC_Plan form
    Set aClearData = New Collection
    For Each aClearCOSRec In aClearCOS.aPlanCOS
        aKey = aClearCOSRec.aYear.Value & "-" & aClearCOSRec.aElem.Value
        If Not contains(aClearData, aKey, "obj") Then
            Set aSAP_PC_Plan = aSAP_PC_Plan.create(aCOArea, aClearCOSRec.aYear.Value, "1", "12", aVers, aCurt)
            aSAP_PC_Plan.setOValues "WBS_ELEMENT", aFWBS
            aSAP_PC_Plan.CostElement = aSAPFormat.unpack(aClearCOSRec.aElem.Value, "10")
            aClearData.add aSAP_PC_Plan, aKey
        End If
    Next
'++ HM160413
'   post the data to SAP
'   This could be done for all object in one year at once
    Dim aSAPData As Collection
    Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning
    Dim aSAPRet As String
'++ HM160413
    Dim aSAPClearRet As String '++ HM160413
    Set aClearMsg = New MessageCollection
    For Each aSAP_PC_Plan In aClearData
        aSAPClearRet = ""
        Set aSAPData = New Collection
        aSAPData.add aSAP_PC_Plan
        aSAPRet = aSAPCostActivityPlanning.PostPrimCostClear(aSAPData)
        If aSAPRet <> "Success" Then
            aSAPClearRet = aSAPClearRet & aSAPRet
        End If
        aClearMsg.addMsg aSAPClearRet, aSAP_PC_Plan.Object("WBS_ELEMENT").Value
    Next
    aSAPRet = ""
    For Each aSAP_PC_Plan In aData
        aSAPClearRet = aClearMsg.aMsgCol(aSAP_PC_Plan.Object("WBS_ELEMENT").Value).Value
        If aSAPClearRet = "" Then
            Set aSAPData = New Collection
            aSAPData.add aSAP_PC_Plan
            aSAPRet = aSAPRet & aSAPCostActivityPlanning.PostPrimCostPer(aSAPData)
        Else
            aSAPRet = "Clear: " & aSAPClearRet
            Exit For
        End If
    Next
'++ HM160413
    aPWS.Cells(i, 7).Value = aSAPRet
    i = i + 1
  Loop While CStr(aPWS.Cells(i, 5)) <> ""
  MsgBox "SAP-Upload finished, please check the messages!", vbInformation + vbOKOnly
  Application.Cursor = xlDefault
  Exit Sub
SAP_PS_COC_PostPrimCost_Err:
  Application.ScreenUpdating = True
  Application.Cursor = xlDefault
  MySAPErr.MSGProt "SAPMakro", "SAP_PS_COC_PostPrimCost", "", err.Number, err.Description
End Sub

Public Function IsRangeName(mySh As Worksheet, RangeName As String) As Boolean
On Error Resume Next
    IsRangeName = False
    IsRangeName = Len(mySh.Range(RangeName).Name) <> 0
End Function

Sub unhide_all()
    Dim aWS As Worksheet
    For Each aWS In ActiveWorkbook.Worksheets
        aWS.Visible = xlSheetVisible
    Next
End Sub

Function getCOSPrices() As PlanCOSPrice
Dim aWS As Worksheet
Dim i As Integer
Dim aPlanCOSPrice As New PlanCOSPrice

Set aWS = ActiveWorkbook.Worksheets("ACCT")
j = 3
Do
    aPlanCOSPrice.addPrice CStr(aWS.Cells(2, j).Value), CStr(aWS.Cells(4, j).Value)
    j = j + 1
Loop While CStr(aWS.Cells(4, j)) <> ""
Set getCOSPrices = aPlanCOSPrice
End Function

Function getElemMapping() As PlanElemMapping
Dim aWS As Worksheet
Dim i As Integer
Dim aPlanElemMapping As New PlanElemMapping

Set aWS = ActiveWorkbook.Worksheets("Parameter")

i = 10
Do
    aPlanElemMapping.addElem CStr(aWS.Cells(i, 1).Value), CStr(aWS.Cells(i, 2).Value)
    i = i + 1
Loop While CStr(aWS.Cells(i, 1)) <> ""
Set getElemMapping = aPlanElemMapping
End Function

Function collectHours(aWS As Worksheet) As PlanHours
Dim aPlanHours As New PlanHours
Dim aYear As Integer
Dim aValue As Double
Dim j As Integer
Dim i As Integer
    Dim a As New PlanHours
    For i = c_HSR To c_HER
'       For j = c_DSC To c_DEC Step 2   '-- EJ160204
        For j = c_DSC To aDEC Step 2    '++ EJ160204
            aYear = Year(aWS.Cells(c_DR, j - 1).Value)
'            aValue = CDbl(aWS.Cells(i, j).Value)   '-- EJ160216
'++ EJ160216
            If UCase(aWS.Cells(c_NoUplR, j).Value) <> "NO UPLOAD" Then ' EJ160216: Ist die FC-Spalte mit "No Upload" markiert, so werden die Werte nicht berücksichtigt
                aValue = CDbl(aWS.Cells(i, j).Value)
            Else
                aValue = 0
            End If
'++ EJ160216
            If aValue <> 0 Then
                aPlanHours.addHours aYear, aValue
            End If
        Next j
    Next i
    Set collectHours = aPlanHours
End Function

Function collectClearData(aWS As Worksheet, aPlanElemMapping As PlanElemMapping) As PlanCOS
Dim aPlanCOS As New PlanCOS
Dim aYear As Integer
Dim aElem As String
Dim aCElem As String
Dim aValue As Double
Dim j As Integer
Dim i As Integer

    For i = c_CSR To c_CER
        aElem = UCase(aWS.Cells(i, c_ELC))
        If aElem <> "" Then
            aCElem = aPlanElemMapping.getCElem(aElem)
            For j = c_DSC To aDEC Step 2
                If aWS.Cells(c_DR, j - 1).Value <> "" Then               'HM160518
                    aYear = Year(aWS.Cells(c_DR, j - 1).Value)
                    aPlanCOS.addCOS aYear, aCElem, 0
                End If                                                   'HM160518
            Next j
        End If
    Next i
'   add the other COS Cost element
    aCElem = aPlanElemMapping.getCElem("COS")
    For j = c_DSC To aDEC Step 2
        If aWS.Cells(c_DR, j - 1).Value <> "" Then                       'HM160518
            aYear = Year(aWS.Cells(c_DR, j - 1).Value)
            aPlanCOS.addCOS aYear, aCElem, 0
        End If                                                           'HM160518
    Next j
    Set collectClearData = aPlanCOS
End Function

Function collectCOS(aWS As Worksheet, aPlanElemMapping As PlanElemMapping) As PlanCOS
Dim aPlanCOS As New PlanCOS
Dim aYear As Integer
Dim aElem As String
Dim aCElem As String
Dim aValue As Double
Dim j As Integer
Dim i As Integer

    For i = c_CSR To c_CER
        aElem = UCase(aWS.Cells(i, c_ELC))
        If aElem <> "" Then
            aCElem = aPlanElemMapping.getCElem(aElem)
'           For j = c_DSC To c_DEC Step 2   '-- EJ160204
            For j = c_DSC To aDEC Step 2    '++ EJ160204
                aYear = Year(aWS.Cells(c_DR, j - 1).Value)
'                aValue = CDbl(aWS.Cells(i, j).Value)   '-- EJ160216
'++ EJ160216
                If UCase(aWS.Cells(c_NoUplR, j).Value) <> "NO UPLOAD" Then ' EJ160216: Ist die FC-Spalte mit "No Upload" markiert, so werden die Werte nicht berücksichtigt
                    aValue = CDbl(aWS.Cells(i, j).Value)
                Else
                    aValue = 0
                End If
'++ EJ160216
                If aValue <> 0 Then
                    aPlanCOS.addCOS aYear, aCElem, aValue
                End If
            Next j
        End If
    Next i
    Set collectCOS = aPlanCOS
End Function

Public Function contains(col As Collection, Key As Variant, Optional aType As String = "var") As Boolean
Dim obj As Object
Dim var As Variant
On Error GoTo err
    contains = True
    If aType = "obj" Then
        Set obj = col(Key)
    Else
        var = col(Key)
    End If
    Exit Function
err:
    contains = False
End Function

'++ EJ160301
Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    Dim ret As Boolean
    ret = False
    wsName = UCase(wsName)
    For Each ws In ActiveWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
End Function
'++ EJ160301
