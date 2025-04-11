Attribute VB_Name = "Main"
Option Explicit
Sub UseFileDialogOpen()
 
    Dim lngCount As Long
    Dim i As Integer
    
    i = 1
     
    'clear NFe and file list
    ThisWorkbook.Sheets("filelist").Columns("a:a").Clear
    
     
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Extensible Markup Language Files", "*.xml"
        .InitialView = msoFileDialogViewDetails
        .Show
        
        'save path files on filelist
        For lngCount = 1 To .SelectedItems.Count
            ThisWorkbook.Sheets("filelist").Cells(i, 1) = .SelectedItems(lngCount)
            i = i + 1
        Next lngCount
 
    End With
    
    Call Main

    MsgBox "NFs carregadas"
    ThisWorkbook.Application.CalculateFull
End Sub


Sub carregarListaNf(docpath As String)
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim ans
    

    ' Create a new XML Document object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    

    ' Configure properties
    xmlDoc.async = False
    xmlDoc.validateOnParse = False

    ' Load the XML file
    If Not xmlDoc.Load(docpath) Then
        MsgBox "Failed to load XML file. Exiting."
        Exit Sub
    End If
    
    ' Set namespace
    Dim XMLNamespaces As String
    XMLNamespaces = "xmlns:s='http://www.abrasf.org.br/nfse.xsd'"
    xmlDoc.SetProperty "SelectionNamespaces", XMLNamespaces
    
    
    Dim nodeCancelamento As IXMLDOMNodeList, nodeNF As IXMLDOMNodeList, j As Integer
    j = ThisWorkbook.Sheets("filelist").Range("C1")
    
    Set nodeNF = xmlDoc.DocumentElement.SelectNodes("s:Nfse")
    Set nodeCancelamento = xmlDoc.DocumentElement.SelectNodes("s:NfseCancelamento")
    
    If findByID(ThisWorkbook.Sheets("NFs").Range("A:A"), renameNF(xmlDoc.SelectSingleNode("//s:Numero").Text)) <> 0 Then
        ans = MsgBox("Nota já foi lançada. Deseja substituir os dados?", vbYesNo)
        If ans = vbYes Then
            j = findByID(ThisWorkbook.Sheets("NFs").Range("A:A"), renameNF(xmlDoc.SelectSingleNode("//s:Numero").Text))
            ThisWorkbook.Sheets("filelist").Range("C1") = ThisWorkbook.Sheets("filelist").Range("C1") - 1
        Else
            Exit Sub
        End If
    End If
    
    Dim cnpjNota As Double, cnpjWB As Double
    cnpjNota = CDbl(xmlDoc.SelectSingleNode("//s:PrestadorServico/s:IdentificacaoPrestador/s:Cnpj").Text)
    cnpjWB = CDbl(ThisWorkbook.Sheets("Menu").Range("I1"))
    If cnpjNota <> cnpjWB Then
        MsgBox "Nota " & renameNF(xmlDoc.SelectSingleNode("//s:Numero").Text) & " não pertence a está empresa"
        Exit Sub
    End If
    
    If Not (xmlDoc.SelectSingleNode("//s:Numero")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 1) = renameNF(xmlDoc.SelectSingleNode("//s:Numero").Text)
    If Not (xmlDoc.SelectSingleNode("//s:DataEmissao")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 2) = formatarData(xmlDoc.SelectSingleNode("//s:DataEmissao").Text)
    If Not (xmlDoc.SelectSingleNode("//s:ValorServicos")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 3) = xmlDoc.SelectSingleNode("//s:ValorServicos").Text Else ThisWorkbook.Sheets("NFs").Cells(j, 3) = 0
    If Not (xmlDoc.SelectSingleNode("//s:ValorPis")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 4) = xmlDoc.SelectSingleNode("//s:ValorPis").Text Else ThisWorkbook.Sheets("NFs").Cells(j, 4) = 0
    If Not (xmlDoc.SelectSingleNode("//s:ValorCofins")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 5) = xmlDoc.SelectSingleNode("//s:ValorCofins").Text Else ThisWorkbook.Sheets("NFs").Cells(j, 5) = 0
    If Not (xmlDoc.SelectSingleNode("//s:ValorIr")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 6) = xmlDoc.SelectSingleNode("//s:ValorIr").Text Else ThisWorkbook.Sheets("NFs").Cells(j, 6) = 0
    If Not (xmlDoc.SelectSingleNode("//s:ValorCsll")) Is Nothing Then ThisWorkbook.Sheets("NFs").Cells(j, 7) = xmlDoc.SelectSingleNode("//s:ValorCsll").Text Else ThisWorkbook.Sheets("NFs").Cells(j, 7) = 0
    If nodeCancelamento.Length > 0 Then ThisWorkbook.Sheets("NFs").Cells(j, 8) = False Else ThisWorkbook.Sheets("NFs").Cells(j, 8) = True
    Call maiorNF(xmlDoc.SelectSingleNode("//s:Numero").Text)
    ThisWorkbook.Sheets("filelist").Range("C1") = ThisWorkbook.Sheets("filelist").Range("C1") + 1
    
    ' Release the XML Document object
    Set xmlDoc = Nothing

End Sub

Sub resetAPP()
    ThisWorkbook.Sheets("filelist").Columns("a:a").Clear
    ThisWorkbook.Sheets("filelist").Range("C1") = 2
    ThisWorkbook.Sheets("filelist").Range("C2:C11") = 0
    ThisWorkbook.Sheets("NFs").Rows("2:1048576") = ""
End Sub
Sub Main()
    Dim i As Integer
    i = 1
    Do Until ThisWorkbook.Sheets("filelist").Cells(i, 1) = ""
        Call carregarListaNf(ThisWorkbook.Sheets("filelist").Cells(i, 1))
        i = i + 1
    Loop
End Sub


