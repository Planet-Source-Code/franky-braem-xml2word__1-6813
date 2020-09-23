VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_oXML As MSXML.DOMDocument
Dim m_sSource As String
Dim m_sWordTemplate As String
Dim m_bFile As Boolean
Dim m_sErr As String
Dim m_sDocPath As String
Dim m_sDocPrefix As String

Public Property Let XMLFile(sFileName As String)

    m_sSource = sFileName
    m_bFile = True
    
End Property

Public Property Let XMLString(sStr As String)

    m_sSource = sStr
    m_bFile = False
    
End Property

Public Function Execute() As Boolean

    Dim bLoaded As Boolean
    Dim oWordDoc As Word.Document
    Dim oWordField As Word.MailMergeField
    Dim oWordTable As Word.Table
    Dim oWordFieldRange As Word.Range
    Dim oXMLDocNode As MSXML.IXMLDOMNode
    Dim oXMLFieldNode As MSXML.IXMLDOMNode
    Dim oXMLParentNode As MSXML.IXMLDOMNode
    Dim oXMLNextNode As MSXML.IXMLDOMNode
    Dim oXMLDoc As MSXML.IXMLDOMNode
    Dim nParentIndex As Integer
    Dim nRow As Integer
    Dim nCol As Integer
    Dim bRowColFound As Boolean
    Dim bTableField As Boolean
    Dim sFieldSplit() As String
    Dim n As Integer
    Dim i As Integer
    Dim sPath As String
    Dim bTablesProcessed() As Boolean
    Dim nTable As Integer
    
    oWord.Visible = True
    
    If m_oXML Is Nothing Then
        Set m_oXML = New MSXML.DOMDocument
    End If
    m_oXML.async = False
    
    If m_bFile Then
        bLoaded = m_oXML.Load(m_sSource)
    Else
        bLoaded = m_oXML.loadXML(m_sSource)
    End If
    
    If bLoaded Then
        
        If m_sDocPath = "" Then
            sPath = GetPath(m_sWordTemplate)
        Else
            sPath = m_sDocPath
        End If
        
        Set oXMLDoc = m_oXML.documentElement
        
        If oXMLDoc.hasChildNodes Then
            For Each oXMLDocNode In oXMLDoc.childNodes
           
                Set oWordDoc = oWord.Documents.Add(m_sWordTemplate, False)
                ReDim bTablesProcessed(oWordDoc.Tables.Count)
                For Each oWordField In oWordDoc.MailMerge.Fields
                    
                    oWordField.Select
                    Set oWordFieldRange = oWord.Selection.Range
                
                    bTableField = False
                    nTable = -1
                    For Each oWordTable In oWordDoc.Tables
                        nTable = nTable + 1
                        If oWordFieldRange.InRange(oWordTable.Range) Then
                            bTableField = True
                            Exit For
                        End If
                    Next oWordTable
                    
                    sFieldSplit = Split(oWordField.Code, " ")
                    sFieldSplit = Split(sFieldSplit(2), "_")
                    
                    Set oXMLFieldNode = oXMLDocNode
                    Set oXMLParentNode = Nothing
                    For n = LBound(sFieldSplit) To UBound(sFieldSplit)
                    
                        Set oXMLFieldNode = oXMLFieldNode.selectSingleNode(sFieldSplit(n))
                        If oXMLFieldNode Is Nothing Then
                            Exit For
                        End If
                        If bTableField And oXMLParentNode Is Nothing Then
                            Set oXMLNextNode = oXMLFieldNode.nextSibling
                            If Not oXMLNextNode Is Nothing Then
                                If oXMLNextNode.baseName = oXMLFieldNode.baseName Then
                                    Set oXMLParentNode = oXMLFieldNode
                                    nParentIndex = n
                                    Exit For
                                End If
                            End If
                        End If
                    Next n
                    
                    If bTableField And Not oXMLParentNode Is Nothing Then
                        nRow = 0
                        nCol = 0
                        Do While Not oXMLParentNode Is Nothing
                            Set oXMLFieldNode = oXMLParentNode
                            For n = nParentIndex + 1 To UBound(sFieldSplit)
                                Set oXMLFieldNode = oXMLFieldNode.selectSingleNode(sFieldSplit(n))
                                If oXMLFieldNode Is Nothing Then
                                    Exit For
                                End If
                            Next n
                            If oXMLFieldNode Is Nothing Then
                                oWordFieldRange.Text = "Unresolved fieldname"
                            Else
                                If nRow = 0 Then
                                    bRowColFound = False
                                    For nRow = 1 To oWordTable.Rows.Count
                                        For nCol = 1 To oWordTable.Columns.Count
                                            If oWordFieldRange.InRange(oWordTable.Cell(nRow, nCol).Range) Then
                                                bRowColFound = True
                                                Exit For
                                            End If
                                        Next nCol
                                        If bRowColFound Then
                                            Exit For
                                        End If
                                    Next nRow
                                    oWordFieldRange.Text = oXMLFieldNode.childNodes.Item(0).Text
                                Else
                                    nRow = nRow + 1
                                    If Not bTablesProcessed(nTable) Then
                                        If nRow <= oWordTable.Rows.Count Then
                                            oWordTable.Rows.Add oWordTable.Rows.Item(nRow)
                                        Else
                                            oWordTable.Rows.Add
                                        End If
                                    End If
                                    oWordTable.Cell(nRow, nCol).Range.Text = oXMLFieldNode.childNodes.Item(0).Text
                                End If
                                Set oXMLParentNode = oXMLParentNode.nextSibling
                            End If
                        Loop
                        bTablesProcessed(nTable) = True
                     Else
                         If oXMLFieldNode Is Nothing Then
                             oWordFieldRange.Text = "Unresolved fieldname"
                         Else
                             oWordFieldRange.Text = oXMLFieldNode.childNodes.Item(0).Text
                         End If
                    End If
                    
                Next oWordField
                
                i = i + 1
                oWordDoc.SaveAs sPath + m_sDocPrefix + CStr(i) + ".doc"
                oWordDoc.Close
                
            Next oXMLDocNode
            
        End If
    
    Else
    
    End If
    Set m_oXML = Nothing
    
End Function

Private Sub Class_Initialize()
    
    If nRef = 0 Then
        Set oWord = New Word.Application
    End If
    nRef = nRef + 1
        
End Sub

Private Sub Class_Terminate()

    Set m_oXML = Nothing
    nRef = nRef - 1
    If nRef = 0 Then
        Set oWord = Nothing
    End If

End Sub

Public Property Get LastError() As String

    LastError = m_sErr

End Property

Public Property Let WordTemplate(sWordTemplate As String)
    
    m_sWordTemplate = sWordTemplate
    
End Property

Public Property Let Path(sPath As String)

    m_sDocPath = sPath
    
End Property

Public Property Let DocumentPrefix(sPrefix As String)

    m_sDocPrefix = sPrefix
    
End Property
