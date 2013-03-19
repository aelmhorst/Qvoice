Attribute VB_Name = "ModData"
Option Explicit
Option Compare Text

Public Type IndexInfo
    Table As String
    IndexName As String
    Fields As String
    Unique As Boolean
    IgnoreNulls As Boolean
    Primary As Boolean
End Type

Public Sub CreateIndexDef(DatabasePath As String)
Dim dbsQvoice As Database
    Dim tdf As TableDef
    Dim idx As Index
    Dim aField As Field
    Dim aInfo() As IndexInfo
    Dim aInfoLength As Integer

    Set dbsQvoice = OpenDatabase(DatabasePath)
    ReDim aInfo(1 To 1000)
    
    aInfoLength = 0
    
    dbsQvoice.t
    
    For Each tdf In dbsQvoice.TableDefs
        If Not tdf.Name Like "MSys*" Then
            For Each idx In tdf.Indexes
                aInfoLength = aInfoLength + 1
                aInfo(aInfoLength).IndexName = idx.Name
                aInfo(aInfoLength).Table = tdf.Name
                aInfo(aInfoLength).Unique = idx.Unique
                aInfo(aInfoLength).IgnoreNulls = idx.IgnoreNulls
                aInfo(aInfoLength).Primary = idx.Primary
                For Each aField In idx.Fields
                    aInfo(aInfoLength).Fields = aInfo(aInfoLength).Fields & aField.Name & ","
                Next
                aInfo(aInfoLength).Fields = Left$(aInfo(aInfoLength).Fields, Len(aInfo(aInfoLength).Fields) - 1)
            Next
        End If
    Next
    dbsQvoice.Close
    
    ReDim Preserve aInfo(1 To aInfoLength)
    SaveInfos aInfo
    
End Sub

Public Sub RebuildIndexes(DatabasePath As String)
Dim dbsQvoice As Database
    Dim tdf As TableDef
    Dim idx As Index
    Dim aField As Field
    Dim aInfo() As IndexInfo
    Dim aCounter As Integer
    Dim x As Integer, y As Integer
    Dim aNames() As String
    Dim aIndexInfos As New Collection
    Dim aIndexfound As Boolean
    
    aInfo = GetInfos
    
    If UBound(aInfo) < 50 Then
        MsgBox "Index info file is missing indexes"
        Exit Sub
    End If
    
    Set dbsQvoice = OpenDatabase(DatabasePath)
    
    
    For Each tdf In dbsQvoice.TableDefs
        If Not tdf.Name Like "MSys*" Then
            frmMainDoctor.Status "Cleaning Table " & tdf.Name
            If Not tdf.Name Like "MSys*" Then
                '// Make a list of all of the indexes
                If tdf.Indexes.Count > 0 Then
                    ReDim aNames(1 To tdf.Indexes.Count)
                    aCounter = 0
                    For Each idx In tdf.Indexes
                        If Not (idx.Foreign Or idx.Primary) Then
                            aCounter = aCounter + 1
                            aNames(aCounter) = idx.Name
                        End If
                    Next
                    '// Remove all of the indexes
                    For x = 1 To aCounter
                        tdf.Indexes.Delete (aNames(x))
                        tdf.Indexes.Refresh
                    Next x
                End If
            End If
            tdf.Indexes.Refresh
        End If
    Next
            
    x = UBound(aInfo)
    For aCounter = 1 To x
        frmMainDoctor.Status "Creating Index " & aInfo(aCounter).IndexName & " in " & aInfo(aCounter).Table
        Set tdf = dbsQvoice.TableDefs(aInfo(aCounter).Table)
        aIndexfound = False
        For Each idx In tdf.Indexes
            If idx.Name = aInfo(aCounter).IndexName Then
                aIndexfound = True
                Exit For
            End If
        Next
        If Not aIndexfound Then
            Set idx = tdf.CreateIndex(aInfo(aCounter).IndexName)
            With idx
                .Unique = aInfo(aCounter).Unique
                .IgnoreNulls = aInfo(aCounter).IgnoreNulls
                .Primary = aInfo(aCounter).Primary
                aNames = Split(aInfo(aCounter).Fields, ",")
                For y = 0 To UBound(aNames)
                    .Fields.Append .CreateField(aNames(y))
                Next
            End With
            tdf.Indexes.Append idx
            tdf.Indexes.Refresh
        End If
    Next aCounter
    
    dbsQvoice.Close
    
    
    MsgBox aCounter & " Database Indexes have been created."
End Sub

Private Sub SaveInfos(inInfos() As IndexInfo)
    If Len(Dir$(GetIndexInfoPath)) > 0 Then
        Kill GetIndexInfoPath
    End If
    Open GetIndexInfoPath For Binary As #1
        Put #1, , inInfos
    Close #1
End Sub

Private Function GetIndexInfoPath() As String
    GetIndexInfoPath = App.Path & "\IndexInfo.data"
End Function


'//In this example, all we did was write three variables to the file. To read them back, use this:
Private Function GetInfos() As IndexInfo()
Dim aInfos() As IndexInfo
Dim x As Integer

    ReDim aInfos(1 To 1000)
    Open GetIndexInfoPath For Binary As #1
        Do Until EOF(1)
            x = x + 1
            Get #1, , aInfos(x)
        Loop
    Close #1
    ReDim Preserve aInfos(1 To x - 1)
    GetInfos = aInfos
End Function


Sub CreateIndexX()

    Dim dbsNorthwind As Database
    Dim tdfEmployees As TableDef
    Dim idxCountry As Index
    Dim idxFirstName As Index
    Dim idxLoop As Index

    Set dbsNorthwind = OpenDatabase("Northwind.mdb")
    Set tdfEmployees = dbsNorthwind!Employees

    With tdfEmployees
        ' Create first Index object, create and append Field
        ' objects to the Index object, and then append the
        ' Index object to the Indexes collection of the
        ' TableDef.
        Set idxCountry = .CreateIndex("CountryIndex")
        With idxCountry
            .Fields.Append .CreateField("Country")
            .Fields.Append .CreateField("LastName")
            .Fields.Append .CreateField("FirstName")
        End With
        .Indexes.Append idxCountry

        ' Create second Index object, create and append Field
        ' objects to the Index object, and then append the
        ' Index object to the Indexes collection of the
        ' TableDef.
        Set idxFirstName = .CreateIndex
        With idxFirstName
            .Name = "FirstNameIndex"
            .Fields.Append .CreateField("FirstName")
            .Fields.Append .CreateField("LastName")
        End With
        .Indexes.Append idxFirstName

        ' Refresh collection so that you can access new Index
        ' objects.
        .Indexes.Refresh

        Debug.Print .Indexes.Count & " Indexes in " & _
            .Name & " TableDef"

        ' Enumerate Indexes collection.
        For Each idxLoop In .Indexes
            Debug.Print "  " & idxLoop.Name
        Next idxLoop

        ' Print report.
        CreateIndexOutput idxCountry
        CreateIndexOutput idxFirstName

        ' Delete new Index objects because this is a
        ' demonstration.
        .Indexes.Delete idxCountry.Name
        .Indexes.Delete idxFirstName.Name
    End With

    dbsNorthwind.Close

End Sub

Function CreateIndexOutput(idxTemp As Index)

    Dim fldLoop As Field
    Dim prpLoop As Property

    With idxTemp
        ' Enumerate Fields collection of Index object.
        Debug.Print "Fields in " & .Name
        For Each fldLoop In .Fields
            Debug.Print "  " & fldLoop.Name
        Next fldLoop

        ' Enumerate Properties collection of Index object.
        Debug.Print "Properties of " & .Name
        For Each prpLoop In .Properties
            Debug.Print "  " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
        Next prpLoop
    End With

End Function


