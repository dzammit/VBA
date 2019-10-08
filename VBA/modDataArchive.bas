Attribute VB_Name = "modDataArchive"
Option Compare Database
Option Explicit

Function DATA_Archive(sTable As String, Optional sWhere As String, Optional sArchiveDB As String, Optional sArchiveTable As String)
' Function Notes:
' Archives data from table 'sTable' into destination database 'sDestDB'.
' If 'sWhere' is provided, only data meeting where clause is archived.
' Destination table name is 'sTable'_Archive by default, or 'sArchiveTable' is used if specified.
    Dim wrkDefault As Workspace
    Dim dbsTemp As Database
    Dim sSql As String

    If sArchiveTable = "" Then sArchiveTable = sTable & "_ARCHIVE"
    If sArchiveDB = "" Then sArchiveDB = CurrentDb.Name & "_archive.mdb"
    
    If Not FileExists(sArchiveDB) Then
        'Create a new database
        Set wrkDefault = DBEngine.Workspaces(0)
        Set dbsTemp = wrkDefault.CreateDatabase(sArchiveDB, dbLangGeneral)

        ' Delete the link to the temp table if it exists
        If TableExists(sArchiveTable) Then CurrentDb.TableDefs.DELETE sArchiveTable
    End If
    
    If Not TableExists(sArchiveTable) Then
        ' Create the temp table
        ' - if it's a linked table, copy from the dest link.
        'If Len(CurrentDb.TableDefs(sTable).connect) > 0 Then ' linked table
        '    Stop
        'End If
        
        sSql = "SELECT * into " & sArchiveTable & " in " & Q(sArchiveDB)
        sSql = sSql & " FROM " & sTable & " where 1=0"
        RunSql sSql
        
        'DoCmd.SetWarnings False
        'DoCmd.TransferDatabase acExport, "Microsoft Access", sArchiveDB, acTable, sTable, sArchiveTable, True
        'DoCmd.SetWarnings True

        LinkTable sArchiveTable, sArchiveDB
    End If


    sSql = "INSERT INTO " & sArchiveTable
    sSql = sSql & " Select * FROM " & sTable
    If Not (sWhere = "") Then
        sSql = sSql & " WHERE " & sWhere
    End If
    
    RunSql sSql

    sSql = "DELETE * FROM " & sTable
    If Not (sWhere = "") Then
        sSql = sSql & " WHERE " & sWhere
    End If
    
    RunSql sSql

End Function

Function DATA_Restore(sTable As String, Optional sWhere As String, Optional sArchiveDB As String, Optional sArchiveTable As String, Optional bClearData As Boolean = True)
    ' Function Notes:
    ' Archives data from table 'sTable' into destination database 'sDestDB'.
    ' If 'sWhere' is provided, only data meeting where clause is archived.
    ' Destination table name is 'sTable'_Archive by default, or 'sArchiveTable' is used if specified.
    Dim sSql As String
    
    If (sArchiveTable = "") Then sArchiveTable = sTable & "_ARCHIVE"
    If sArchiveDB = "" Then sArchiveDB = CurrentDb.Name & "_archive.mdb"
    
    If Not FileExists(sArchiveDB) Then
        err.Raise 1, "Unable to locate archive database " & Q(sArchiveDB), "Unable to locate archive database " & Q(sArchiveDB)
        Exit Function
    End If
    
    If Not TableExists(sArchiveTable) Then
        If Not (LinkTable(sArchiveTable, sArchiveDB)) Then
            err.Raise 1, "Unable to restore data - error in linking to archive database"
            Exit Function
        End If
    End If
    
    
    sSql = "INSERT INTO " & sTable
    sSql = sSql & " Select * FROM " & sArchiveTable
    If Not (sWhere = "") Then
        sSql = sSql & " WHERE " & sWhere
    End If
    
    RunSql sSql
    
    If bClearData Then
    sSql = "DELETE * FROM " & sArchiveTable
    If Not (sWhere = "") Then
        sSql = sSql & " WHERE " & sWhere
    End If
    
    RunSql sSql
    End If
End Function
