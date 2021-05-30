Attribute VB_Name = "Module1"
Option Explicit

Function incell(analyzed_cell As Range)

Dim regex As Object, eachcell, rangeobj As Object
Dim indcol As New Collection
Dim operators As Object, operator
Dim i As Integer, j As Integer
Dim maindict As Object, arraydict As Object
Dim midstr As String, refmidstr As String, midtxt As String

Set maindict = CreateObject("Scripting.Dictionary")
Set arraydict = CreateObject("Scripting.Dictionary")

Set regex = CreateObject("VBScript.RegExp")
regex.Pattern = "[^a-zA-Z0-9!:\[\]_]"
regex.Global = True

Set operators = regex.Execute(analyzed_cell.Formula)

For Each operator In operators
    indcol.Add operator.firstindex
Next

If indcol.count = 0 Then indcol.Add 0 'in case there is nothing except the first "=", assume it to be 0
                    
If indcol(indcol.Count) < Len(analyzed_cell.Formula) Then indcol.Add Len(analyzed_cell.Formula)

For i = 1 To indcol.count - 1
    midstr = Mid(analyzed_cell.Formula, indcol(i) + 1, indcol(i + 1) - indcol(i))
    refmidstr = Mid(midstr, 2, Len(midstr) - 1)
    midtxt = ""
    
    On Error Resume Next
        Set rangeobj = Range(refmidstr) 'in case of error, will return Nothing or previous value (which is set to nothing in code below)
        
        If rangeobj Is Nothing Then
            midtxt = refmidstr
        Else
        
            If IsArray(rangeobj.Value) Then
                j = 1
                For Each eachcell In rangeobj
                    arraydict.Add j, IIf(eachcell.NumberFormat = "General", eachcell.Value, format(eachcell.Value, eachcell.NumberFormat))
                    j = j + 1
                Next
                
                midtxt = Join(arraydict.items, ",")
                arraydict.RemoveAll
            
            Else
                midtxt = IIf(rangeobj.NumberFormat = "General", rangeobj.Value, format(rangeobj.Value, rangeobj.NumberFormat))
            
            End If
        End If
        
    maindict.Add i, left(midstr, 1) & midtxt
    Set rangeobj = Nothing 'required to clear existing, otherwise will not return to nothing in the set above
    Err.Clear

Next i

incell = Join(maindict.items, "")

Set rangeobj = Nothing
Set maindict = Nothing
Set arraydict = Nothing
Set regex = Nothing
    
End Function
