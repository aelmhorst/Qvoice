Attribute VB_Name = "QueryDefSQLCreator"
Option Explicit

Public Sub CreateUpdateSQL(inDataBasePath As String, inStartDate As Date, inFileName As String)
    Dim db As Database
    Dim qDef As QueryDef
    Dim aXmlDoc As New MSXML2.DOMDocument
    Dim aQueryDefs As MSXML2.IXMLDOMElement
    Dim aElement As MSXML2.IXMLDOMElement
    Dim aChildElement As MSXML2.IXMLDOMElement
    Dim aParameter As MSXML2.IXMLDOMElement
    Dim aParam As Parameter
    
    '// Create the XML Doc
    aXmlDoc.loadXML "<QvoiceDB/>"
    Set aQueryDefs = aXmlDoc.createElement("QueryDefs")
    aXmlDoc.childNodes(0).appendChild aQueryDefs
    
    Set db = OpenDatabase(inDataBasePath)
    For Each qDef In db.QueryDefs
        If qDef.LastUpdated > inStartDate And Left$(qDef.Name, 1) <> "~" Then
            Set aElement = aXmlDoc.createElement("QueryDef")
            aQueryDefs.appendChild aElement
            
            aElement.setAttribute "Name", qDef.Name
            aElement.setAttribute "LastUpdated", qDef.LastUpdated
            AppendProperties aElement, qDef.Properties
            
            '// SQL
            Set aChildElement = aXmlDoc.createElement("SQL")
            aChildElement.Text = qDef.SQL
            aElement.appendChild aChildElement
            
            '// Properties
            Set aChildElement = aXmlDoc.createElement("Parameters")
            For Each aParam In qDef.Parameters
                Set aParameter = aXmlDoc.createElement("Parameter")
                aParameter.setAttribute "Name", aParam.Name
                aParameter.setAttribute "Type", aParam.Type
                aParameter.setAttribute "Direction", aParam.Direction
                AppendProperties aParameter, aParam.Properties
                aChildElement.appendChild aParameter
            Next
            aElement.appendChild aChildElement
        End If
    Next
    aXmlDoc.save inFileName
End Sub

Private Sub AppendProperties(inElement As IXMLDOMElement, inProperties As Properties)
    Dim aProp As Property
    On Error Resume Next
    For Each aProp In inProperties
        'inElement.setAttribute aProp.Name, aProp.Value
    Next
End Sub


Public Sub UpdateDatabaseQueryDefs(inDatabase As Database)
    Dim aDoc As New MSXML2.DOMDocument
    Dim aElement As MSXML2.IXMLDOMElement
    '//Dim aQueryDef As QueryDef
    Dim aQueryName As String
    Dim aDeleted As Boolean
    
    aDoc.Load App.Path & "\QueryDefs.xml"

    For Each aElement In aDoc.firstChild.selectSingleNode("QueryDefs").childNodes
        aDeleted = False
        aQueryName = aElement.getAttribute("Name")
        If Left$(aQueryName, 7) = "_Delete" Then
            aDeleted = True
            aQueryName = Trim$(Mid$(aQueryName, 8))
        End If
        
        On Error Resume Next
        inDatabase.QueryDefs.Delete aQueryName
        If Err.Number > 0 Then
            If Err.Number = 3265 Then
                Err.Clear
            Else
                MsgBox Err.Description, vbCritical, App.Title
                Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                Exit Sub
            End If
        End If
        On Error GoTo 0
        
        If Not aDeleted Then
            inDatabase.CreateQueryDef aQueryName, aElement.selectSingleNode("SQL").Text
        End If
    Next
End Sub
