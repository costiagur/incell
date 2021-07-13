Attribute VB_Name = "Module1"
Option Explicit

Function incell(analyzed_cell As Range)

Dim Regex As Object, eachcell, rangeobj As Object
Dim indcol As New Collection
Dim operators As Object, operator
Dim i As Integer, j As Integer
Dim maindict As Object, arraydict As Object
Dim midstr As String, refmidstr As String, midtxt As String
Dim pivMatches As Object, pivMatch As Object
Dim nextoperator As Integer, leftstr As String

Set maindict = CreateObject("Scripting.Dictionary")
Set arraydict = CreateObject("Scripting.Dictionary")

Set Regex = CreateObject("VBScript.RegExp")
Regex.pattern = "[\+\-=/\*\(\),<>]"
Regex.Global = True

Set operators = Regex.Execute(analyzed_cell.Formula)

For Each operator In operators
    indcol.Add operator.firstindex
Next

If indcol.count = 0 Then indcol.Add 0 'in case there is nothing except the first "=", assume it to be 0

If indcol(indcol.count) < Len(analyzed_cell.Formula) Then indcol.Add Len(analyzed_cell.Formula)

For i = 1 To indcol.count - 1
    midstr = Mid(analyzed_cell.Formula, indcol(i) + 1, indcol(i + 1) - indcol(i))
    refmidstr = Mid(midstr, 2, Len(midstr) - 1)
    midtxt = ""
    
    If InStr(1, refmidstr, "GETPIVOTDATA") Then
        Regex.pattern = "[\+\-=/\*<>]"
        Regex.Global = False
        
        leftstr = Mid(analyzed_cell.Formula, indcol(i) + 2, Len(analyzed_cell.Formula))

        If Regex.test(leftstr) Then
            
            Set pivMatches = Regex.Execute(leftstr)
            
            For Each pivMatch In pivMatches
                nextoperator = pivMatch.firstindex
            Next
        Else
            nextoperator = Len(analyzed_cell.Formula)
        End If
        
        leftstr = Mid(analyzed_cell.Formula, indcol(i) + 2, nextoperator)
        midtxt = getpivotdataval(leftstr)
        
        'now we need to skip all the mathches inside GETPIVOTDATA function
        Regex.pattern = "[\+\-=/\*\(\),<>]"
        Regex.Global = True
        
        Set pivMatches = Regex.Execute(leftstr)
            
        j = 0
        For Each pivMatch In pivMatches
            j = j + 1
        Next
        
        i = i + j 'move i forward by num of mathches inside GETPIVOTDATA
        
        Set pivMatches = Nothing
        GoTo nexti
    End If
    
    If InStr(1, refmidstr, "[@") > 0 Then 'if reference is part of list
        If analyzed_cell.ListObject Is Nothing Then 'if analyzed cell is not in listobject
            midtxt = Replace(refmidstr, "@", "[#Headers],") 'get address of the table header above the referenced cell
        Else
            midtxt = Replace(refmidstr, "[@", analyzed_cell.ListObject.Name & "[[#Headers],") 'get address of the table header above the referenced cell
        End If
        
        refmidstr = Range(midtxt).Offset(analyzed_cell.Row() - Range(midtxt).Row(), 0).Address 'get the address relative to the analyzed cell
        midtxt = ""

    End If
    
     If InStr(1, refmidstr, "[#Totals]") > 0 Then 'if reference is part of totals of the list
        j = InStr(indcol(i) + 1, analyzed_cell.Formula, "]]")
        midstr = Mid(analyzed_cell.Formula, indcol(i) + 1, j - indcol(i) + 1)
        refmidstr = Trim(Mid(midstr, 2, Len(midstr) - 1))

        refmidstr = Range(refmidstr).Address
        
        i = i + 1 'move i forward to skip next [

    End If                       
                            
    On Error Resume Next
        Set rangeobj = Range(refmidstr)
        
        If rangeobj Is Nothing Then
            midtxt = refmidstr
        Else
        
            If IsArray(rangeobj.Value) Then
                j = 1
                For Each eachcell In rangeobj
                    Select Case eachcell.NumberFormat
                        Case "General"
                            arraydict.Add j, eachcell.Value
                        Case Else
                            arraydict.Add j, format(eachcell.Value, eachcell.NumberFormat)
                    End Select
                    
                    j = j + 1
                Next
                
                midtxt = Join(arraydict.items, ",")
                arraydict.RemoveAll
            
            Else
                
                Select Case rangeobj.NumberFormat
                    Case "General"
                        midtxt = rangeobj.Value
                    Case Else
                        midtxt = format(rangeobj.Value, rangeobj.NumberFormat)
                End Select
            
            End If
            

        End If

nexti:

    maindict.Add i, left(midstr, 1) & midtxt
    Set rangeobj = Nothing 'required to clear existing, otherwise will not return to nothing in the set above
    Err.Clear

Next i

incell = Join(maindict.items, "")

Set rangeobj = Nothing
Set maindict = Nothing
Set arraydict = Nothing
Set Regex = Nothing
    
End Function

Function getpivotdataval(inistring)
Dim resstring As String
Dim Regex As Object
Dim MatchRes As Object
Dim eachMatch As Object

Set Regex = CreateObject("VBScript.RegExp")

With Regex
    .pattern = "(\(.*\)$)"
    .Global = True
    .IgnoreCase = True
End With

Set MatchRes = Regex.Execute(inistring)

resstring = ""

For Each eachMatch In MatchRes

    resstring = eachMatch.Value

Next

getpivotdataval = Evaluate("=GetPivotData" & resstring)

End Function
