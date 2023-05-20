	<input type="hidden" name="CallingPage"  Id="CallingPage"  value ="<%=NomePagina%>">
	<input type="hidden" name="Oper"         Id="Oper"         value ="">
	<input type="hidden" name="NameRangeD"   Id="NameRangeD"   value ="<%=NameRangeD%>">
	<input type="hidden" name="NameRangeN"   Id="NameRangeN"   value ="<%=NameRangeN%>">
	<input type="hidden" name="NameRangeDT"  Id="NameRangeDT"  value ="<%=NameRangeDT%>">
	<input type="hidden" name="NameLoaded"   Id="NameLoaded"   value ="<%=NameLoaded%>">
	<input type="hidden" name="DescLoaded"   Id="DescLoaded"   value ="<%=DescLoaded%>">
	<input type="hidden" name="Pagina"       Id="Pagina"       value ="<%=cPag%>">
	<input type="hidden" name="RowPagina"    Id="RowPagina"    value ="<%=PageSize%>">
	<input type="hidden" name="TimePage"     Id="TimePage"     value ="<%=TimeStamp%>">
	<input type="hidden" name="ItemToRemove" Id="ItemToRemove" value ="0">
	<input type="hidden" name="ItemToModify" Id="ItemToModify" value ="<%=ItemToModify%>">
	<input type="hidden" name="PaginaReturn" Id="PaginaReturn" value ="<%=PaginaReturn%>">
	<input type="hidden" name="IdParm"       Id="IdParm"       value ="">
	<input type="hidden" name="DataDiOggi"   Id="DataDiOggi"   value ="<%=Stod(Dtos())%>">
	<%
	Session("TimePageLoad")=Dtos() & TimeTos()
	%>
	<input type="hidden" name="TimePageLoad"      Id="TimePageLoad" value ="<%=Session("TimePageLoad")%>">
	<input type="hidden" name="hiddenVirtualPath" id="hiddenVirtualPath" value = "<%=VirtualPath%>">
	
	<!--#include virtual="/gscVirtual/utility/modalInfo.asp"-->
	<!--#include virtual="/gscVirtual/utility/modalConfirm.asp"-->