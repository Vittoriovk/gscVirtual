<!--METADATA Type="typelib" uuid="{00000205-0000-0010-8000-00AA006D2EA4}"-->
<%

function ServizioRichiesto_leggiPrezziProdotto(IdServizioRichiesto,IdSessione,IdAccountCliente,IdProdotto,IdProdottoTemplate,IdFornitore,IdCompagnia,ImptBaseCalcolo,PrezzoServizio,ImptProvvigioni,Giorni) 
Dim Es,Desc,RetValue,TmpNum,TmpChar

on error resume next
    err.clear

    set Command = Server.CreateObject("ADODB.Command")
    Command.ActiveConnection = ConnMsde
    
    Es=0
    Desc=""
    RetValue=true
    
    command.CommandText = "ServizioRichiesto_leggiPrezziProdotto"
    command.CommandType = adCmdStoredProc

    TmpNum  = NumForDb(IdServizioRichiesto)
    set objParameter = command.CreateParameter ("@IdServizioRichiesto"  , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpChar = IdSessione
    set objParameter = command.CreateParameter ("@IdSessione"           , adVarChar, adParamInput, 100, TmpChar)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(IdAccountCliente)
    set objParameter = command.CreateParameter ("@IdAccountCliente"     , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(IdProdotto)
    set objParameter = command.CreateParameter ("@IdProdotto"           , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(IdProdottoTemplate)
    set objParameter = command.CreateParameter ("@IdProdottoTemplate"   , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(IdFornitore)
    set objParameter = command.CreateParameter ("@IdFornitore"          , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(IdCompagnia)
    set objParameter = command.CreateParameter ("@IdCompagnia"          , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(ImptBaseCalcolo)
    set objParameter = command.CreateParameter ("@ImptBaseCalcolo"      , adDouble , adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(PrezzoServizio)
    set objParameter = command.CreateParameter ("@PrezzoServizio"       , adDouble , adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    TmpNum  = NumForDb(ImptProvvigioni)
    set objParameter = command.CreateParameter ("@ImptProvvigioni"      , adDouble , adParamInput,    , TmpNum)
    command.Parameters.Append objParameter
 
    TmpNum  = NumForDb(giorni)
    set objParameter = command.CreateParameter ("@giorni"               , adInteger, adParamInput,    , TmpNum)
    command.Parameters.Append objParameter

    'execute per eseguire senza avere un recordset di ritorno
    command.Execute , , adExecuteNoRecords
    
	xx=writeTraceAttivita("ServizioRichiesto_leggiPrezziProdotto " & err.description ,"Serv_requ",IdServizioRichiesto)  
	
    If err.number<>0 then
        RetValue=false
        ReTCode=err.number
        RetDesc=err.description
    Else
        ReTCode = 0
        RetDesc = "OK"
    End If
    
    set Command = nothing
    err.clear
    ServizioRichiesto_leggiPrezziProdotto=RetValue
   
end function 

%>