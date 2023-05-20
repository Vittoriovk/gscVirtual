<script language="JavaScript">

var sepDataF = "~~";
/* contiene l'elenco di tutte le mappe degli header */
var mapHeader = new Map();
/* contiene l'elenco di tutte le mappe dei della riga da scrivere a video */
var mapHeaderRow = new Map();
/* contiene l'elenco di tutte le mappe dei dati */
var mapHeaderField = new Map();

/* struttura da creare : se flagEmpty=true azzera tutti i campi */
function setupHeader(tabName,flagEmpty) {
	/* si verifica se è stata gia' caricata la mappa per la struttura */
	if (mapHeader.has(tabName)==false) {
		/* inserisco l'elenco dei campi della struttura */
		var myHeader      = new Map();
		var myHeaderRow   = new Map();
		var myHeaderField = new Map();
		var header = $("#"+tabName + "_Header").val();
		var allH   = header.split(";");

		for (i = 0; i < allH.length; i++) {
			oneH = allH[i].trim();
			if (oneH.length>0) {
				var alloneH   = oneH.split(",");
				var txtH = alloneH[0].trim();
				if (txtH.length>0) {
					myHeader.set(txtH,i);
					myHeaderField.set(i,"");
				}
				/* vedo se devo mostrare il campo */
				txtH = alloneH[1].trim();
				if (txtH.length>0) {
					myHeaderRow.set(txtH,i);
				}				
			}
		}
		mapHeader.set(tabName,myHeader);
		mapHeaderRow.set(tabName,myHeaderRow);
		mapHeaderField.set(tabName,myHeaderField);
	}
	if (flagEmpty==true) {
		myMap = mapHeaderField.get(tabName);
		for (var key of myMap.keys()) {
			myMap.set(key,"");
		}
		mapHeaderField.set(tabName,myMap);
	}	
}	

function readStructureRow(tabName,idxRow)
{
	/*legge la riga e mette i dati nella mappa dei campi */
	setupHeader(tabName,true);
	var myMap = mapHeaderField.get(tabName);
	var dataRow = "";
	/*se esiste la riga la carico nela variabile */
	if($("#"+tabName + "_Row_" + idxRow).length)
		dataRow = $("#"+tabName + "_Row_" + idxRow).val();
	
	if (dataRow.length>0) {
		var singleDataRow = dataRow.split("~~");
		for (var key of myMap.keys()) {
			myMap.set(key,singleDataRow[key].trim());
		}
	}
	else {
		for (var key of myMap.keys()) {
			myMap.set(key,"");
		}
	}
	mapHeaderField.set(tabName,myMap);
	return myMap;
}
function evaluateMinusPlus(tabName)
{
	var n = tabName + "_plusMinus";
	var act=$("#" + n).val();
	
	if (act=='+') {
		var act=$("#" + n).val('-');
		$("#" + tabName + "_Plus").hide();
		$("#" + tabName + "_Minus").show();
		}
	else 
		{
		var act=$("#" + n).val('+');
		$("#" + tabName + "_Plus").show();
		$("#" + tabName + "_Minus").hide();			
		}
}

/*popola i campi del form dalla riga caricata in precedenza */
function populateFormByRow(tabName)
{
	var prefix = $("#"+tabName + "_Prefix").val();
	var postfix= $("#"+tabName + "_Postfix").val();

	var idx = $("#" + tabName + "_Idx0").val();
	readStructureRow(tabName,idx);
	// recupero il vettore con i campi : la funzione in functionTable
	var myMap = mapHeader.get(tabName);
	for (var key of myMap.keys()) {
		var s = readItemFromStructure(tabName,key);
		var n = prefix + key + postfix;
		$("#" + n).val(s);
		$("#" + n).css("background", "white");
	}	
}
	
function readItemFromMap(mapName,itemName){
var retVal="";
	try {
		if (mapName.has(itemName));
			retVal=mapName.get(itemName);
	} catch (error) {
		retVal="";
	}
	return retVal;
}

/* legge il form e crea la struttura da scrivere */
function createStructureRow(tabName)
{
	var prefix = $("#"+tabName + "_Prefix").val();
	var postfix= $("#"+tabName + "_Postfix").val();
	
	var retStr="";
	/* elenco dei campi da leggere */
	var myHeader      = mapHeader.get(tabName);
	/* sequenza dei campi da scrivere */
	var myHeaderField = mapHeaderField.get(tabName);

	for (var key of myHeader.keys()) {

		var seq = myHeader.get(key);
		var cam = prefix + key + postfix;

		var val = $("#" + cam).val();
		myHeaderField.set(seq,val);
	}
	for (i = 0; i < 100; i++) {
		if (myHeaderField.has(i))
			retStr = retStr + myHeaderField.get(i) + sepDataF;
	}
	return retStr;
}

/* legge un campo dalle strutture  */
function readItemFromStructure(tabName,itemName){
var retVal="";
var mapH    = mapHeader.get(tabName);
var mapName = mapHeaderField.get(tabName);
	try {
		if (mapH.has(itemName)==false)
			return "";
		var idx = mapH.get(itemName);
		if (mapName.has(idx))
			retVal=mapName.get(idx);
	} catch (error) {
		retVal="";
	}
	return retVal;
}

</script>
