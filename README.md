# Divide-Excel-and-create-dinamic-table
VBA for dividing a spreadsheet in tabs and creating different files with dinamic tables 

Sub DivisãoTransp()
'
' DivisãoTransp Macro
' Divisão das transportadoras por abas
'

'
Dim n As Long
n = Range("A1").CurrentRegion.Rows.Count


    ActiveSheet.Range("$A$1:$V" & n).AutoFilter Field:=3, Criteria1:= _
        "CEVA BR"
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Planilha2").Select
    Sheets("Planilha2").Name = "Ceva"
    Sheets("part-1").Select
    
    ActiveSheet.Range("$A$1:$V" & n).AutoFilter Field:=3, Criteria1:= _
        "Imile Brazil"
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Planilha2").Select
    ActiveSheet.Paste
    Sheets("Planilha3").Select
    Sheets("Planilha3").Name = "Imile"
    Sheets("part-1").Select

    ActiveSheet.Range("$A$1:$V" & n).AutoFilter Field:=3, Criteria1:="Pegaki"
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Planilha4").Select
    Sheets("Planilha4").Name = "Pegaki"
    Sheets("part-1").Select


    ActiveSheet.Range("$A$1:$V" & n).AutoFilter Field:=3, Criteria1:= _
        "SEQUOIA BR"
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Planilha5").Select
    Sheets("Planilha5").Name = "Sequoia"
    ActiveSheet.Paste
    Sheets("part-1").Select

    ActiveSheet.Range("$A$1:$V" & n).AutoFilter Field:=3, Criteria1:= _
        "Splog Brasil"
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Planilha6").Select
    Sheets("Planilha6").Name = "Anjun"
    Sheets("part-1").Select

End Sub
