Attribute VB_Name = "Module1"
Option Explicit

Function incell(analyzed_cell As Range)
Dim regex As Object, eachcell, rangeobj As Object
Dim indcol As New Collection
Dim operators As Object, operator
Dim i As Integer, j As Integer
Dim mainict As Object, arraydict As Object
Dim midstr As String, refmidstr As String, midtxt As String

j = 1

Set mainict = CreateObject("Scripting.Dictionary")
Set arraydict = CreateObject("Scripting.Dictionary")

Set regex = CreateObject("VBScript.RegExp")
regex.Pattern = "[^a-zA-Z0-9!:\[\]_]"
regex.Global = True

Set operators = regex.Execute(analyzed_cell.Formula)

For Each operator In operators
    indcol.Add operator.firstindex
Next

If indcol(indcol.Count) < Len(analyzed_cell.Formula) Then indcol.Add Len(analyzed_cell.Formula)

For i = 1 To indcol.Count - 1
    midstr = Mid(analyzed_cell.Formula, indcol(i) + 1, indcol(i + 1) - indcol(i))
    refmidstr = Mid(midstr, 2, Len(midstr) - 1)
    midtxt = ""
    If refmidstr = "" Then GoTo textval
    Err.Clear
    
    On Error Resume Next
        Set rangeobj = Range(refmidstr)
        
        If IsArray(rangeobj.Value) Then
            j = 1
            For Each eachcell In rangeobj
                arraydict.Add j, IIf(eachcell.NumberFormat = "General", eachcell.Value, Format(eachcell.Value, eachcell.NumberFormat))
                j = j + 1
            Next
            
            midtxt = Join(arraydict.items, ",")
            arraydict.RemoveAll
        
        Else
            midtxt = IIf(rangeobj.NumberFormat = "General", rangeobj.Value, Format(rangeobj.Value, rangeobj.NumberFormat))
        
        End If
                
textval:
        If Err.Number = 1004 Then
            midtxt = refmidstr
        End If
        
    mainict.Add i, Left(midstr, 1) & midtxt

Next i

incell = Join(mainict.items, "")

Set rangeobj = Nothing
Set analyzed_cell = Nothing
Set mainict = Nothing
Set arraydict = Nothing
Set regex = Nothing

End Function
