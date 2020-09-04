Attribute VB_Name = "mdl_MsAccess2MySql"
Option Compare Database

Sub Obj()
'' CONSULTA PARA LISTAR TABELAS DO ACCESS
''SELECT MSysObjects.name FROM MSysObjects WHERE (((MSysObjects.name)<>"tmpObj") AND ((MSysObjects.Type)=1) AND ((MSysObjects.Flags)=0));

'' RELAÇÃO DE TABELAS PARA ALTERAÇÕES
Dim rst As DAO.Recordset: Set rst = CurrentDb.OpenRecordset("Select nome from tmpObj")
Dim sql As String

    While Not rst.EOF
    
    sql = "ALTER TABLE " & rst.Fields("nome").Value & " add ID_EMPRESA long"
    Debug.Print sql
    
    DoCmd.RunSQL sql, 0
    
    rst.MoveNext
    
    Wend

CurrentDb.Close

End Sub


Private Sub ShowTableFields()
'' LISTA DE CAMPOS DE TABELAS
Dim db As Database
Dim tdf As TableDef
Dim x As Integer

Set db = CurrentDb

For Each tdf In db.TableDefs
   If Left(tdf.Name, 4) <> "MSys" Then ' Don't enumerate the system tables
      For x = 0 To tdf.Fields.Count - 1
      Debug.Print tdf.Name & "','" & tdf.Fields(x).Name
      Next x
   End If
Next tdf
End Sub


Private Sub ShowProcedure()
'' CRIAR MODELO DE PROCEDURES COM SYNTAXE MYSQL
Dim db As Database
Dim tdf As TableDef
Dim x As Integer
Dim sSQL As String

Dim sTmp As String
Dim sTmp2 As String

sTmp = "DROP PROCEDURE dpPROCEDURE" & vbNewLine
sTmp = sTmp & "CREATE PROCEDURE spPROCEDURE " & vbNewLine
sTmp = sTmp & "LANGUAGE sql" & vbNewLine
sTmp = sTmp & "NOT DETERMINISTIC" & vbNewLine
sTmp = sTmp & "CONTAINS sql" & vbNewLine
sTmp = sTmp & "SQL SECURITY DEFINER" & vbNewLine
sTmp = sTmp & "COMMENT ''" & vbNewLine

sTmp = sTmp & "BEGIN" & vbNewLine
sTmp = sTmp & "IF p_ID = 0 THEN " & vbNewLine
sTmp = sTmp & " INSERT INTO tbl_Tabela " & vbNewLine
sTmp = sTmp & "         ( " & vbNewLine
sTmp = sTmp & "         ID_EMPRESA  , " & vbNewLine
sTmp = sTmp & "         CNPJ_CPF " & vbNewLine
sTmp = sTmp & "         ) " & vbNewLine
sTmp = sTmp & "    VALUES  " & vbNewLine
sTmp = sTmp & "         ( " & vbNewLine
sTmp = sTmp & "         p_ID_EMPRESA, " & vbNewLine
sTmp = sTmp & "         trim(ucase(p_CNPJ_CPF)) ,  " & vbNewLine
sTmp = sTmp & "          " & vbNewLine
sTmp = sTmp & "         ); " & vbNewLine
sTmp = sTmp & "ELSEIF p_ID <> 0 THEN " & vbNewLine
sTmp = sTmp & " IF p_NOME IS NOT NULL THEN " & vbNewLine
sTmp = sTmp & "     UPDATE tbl_Tabela " & vbNewLine
sTmp = sTmp & "         SET  " & vbNewLine
sTmp = sTmp & "             CNPJ_CPF    =   trim(ucase(p_CNPJ_CPF)) ,  " & vbNewLine
sTmp = sTmp & "             NOME        =   trim(ucase(p_NOME))      " & vbNewLine
sTmp = sTmp & "         WHERE ID = p_ID; " & vbNewLine
sTmp = sTmp & " ELSE " & vbNewLine
sTmp = sTmp & "     DELETE FROM tbl_Tabela WHERE ID = p_ID; " & vbNewLine
sTmp = sTmp & " END IF; " & vbNewLine
sTmp = sTmp & "END IF;  " & vbNewLine

sTmp = sTmp & "END" & vbNewLine

Set db = CurrentDb

For Each tdf In db.TableDefs
    If Left(tdf.Name, 4) <> "MSys" Then ' Don't enumerate the system tables
        
        '' PARAMETROS
        sSQL = ""
        sSQL = sSQL & Replace(Replace(tdf.Name, "tbl", "sp"), "_", "") & "("
        For x = 0 To tdf.Fields.Count - 1
           sSQL = sSQL & "IN `" & tdf.Fields(x).Name & IIf(Left(tdf.Fields(x).Name, 2) = "ID", "` INT,", "` VARCHAR(50),")
        Next x
        sSQL = Left(sSQL, Len(sSQL) - 1) & ")"
'        sSQL = sSQL & vbNewLine
        
        '' CARREGAR LAYOUT
        sTmp2 = sTmp
        
        '' DROP PROCEDURE
        sTmp2 = Replace(sTmp2, "dpPROCEDURE", Replace(Replace(tdf.Name, "tbl", "sp"), "_", ""))
        
        '' CREATE PROCEDURE
        sTmp2 = Replace(sTmp2, "spPROCEDURE", sSQL)
        
        '' TABLE
        sTmp2 = Replace(sTmp2, "tbl_Tabela", tdf.Name)
                
        GerarSaida sTmp2, "saida.log"
        
    End If
Next tdf

db.Close

End Sub


Public Function GerarSaida(strConteudo As String, strArquivo As String)
'' GERAR ARQUIVO DE LOG
    
    Open Application.CurrentProject.Path & "\" & strArquivo For Append As #1
    Print #1, strConteudo
    Close #1

End Function

