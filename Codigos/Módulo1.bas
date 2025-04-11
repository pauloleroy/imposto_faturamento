Attribute VB_Name = "Módulo1"
Option Explicit

Function nfFaltante(ano As Integer) As String
    Dim idColumn As Range, i As Integer, nr As Long, maiorNF As Integer
    maiorNF = ThisWorkbook.Sheets("filelist").Cells(2 + ano - 2024, 3)
    Set idColumn = ThisWorkbook.Sheets("NFs").Range("A:A")
    nfFaltante = ""
    For i = 1 To maiorNF
        nr = findByID(idColumn, ano & "/" & i)
        If nr = 0 Then nfFaltante = nfFaltante & ano & "/" & i & " - "
    Next
    If Len(nfFaltante) > 3 Then nfFaltante = Left(nfFaltante, Len(nfFaltante) - 3)
    If nfFaltante = "" Then nfFaltante = "Numeração Correta"
End Function

Function findByID(idColumn As Range, id As String) As Long
    Dim found As Range
    Set found = idColumn.Find(id, lookat:=xlWhole)
    If Not found Is Nothing Then
        findByID = found.Row
    Else
        findByID = 0
    End If
End Function
Sub tteste()
    Call nfFaltante(2025)
End Sub
Function renameNF(nfName As String) As String
    Dim ano As String, numero As String, cont As Integer, i As Integer
    cont = 0
    ano = Left(nfName, 4)
    For i = 1 To Len(nfName)
        cont = cont + 1
        If Mid(nfName, 4 + cont, 1) <> "0" Then Exit For
    Next
    renameNF = Left(nfName, 4) & "/" & Right(nfName, 12 - cont)
End Function
Function getNumber(nfName As String) As Integer
    Dim i As Integer, cont As Integer
    For i = 1 To Len(nfName)
        cont = cont + 1
        If Mid(nfName, 4 + cont, 1) <> "0" Then Exit For
    Next
    getNumber = CInt(Right(nfName, Len(nfName) - 3 - cont) * 1)
End Function
Function maiorNF(nfName As String) As Integer
    Dim nNF As Integer, anoNF As Integer
    nNF = getNumber(nfName)
    anoNF = Left(nfName, 4)
    If nNF > ThisWorkbook.Sheets("filelist").Cells(2 + anoNF - 2024, 3) Then ThisWorkbook.Sheets("filelist").Cells(2 + anoNF - 2024, 3) = nNF
End Function
Function formatarData(dateString As String)
    Dim formatedDate As Date, ano As Integer, mes As Integer, dia As Integer
    ano = Left(dateString, 4)
    mes = Mid(dateString, 6, 2)
    dia = Mid(dateString, 9, 2)
    
    formatarData = DateSerial(ano, mes, dia)
End Function
