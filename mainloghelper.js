function getcurrenttimewithoffset(offset)
{
var currentdate = new Date(); 
currentdate.setDate(currentdate.getDate()-offset);
var datetime = currentdate.getFullYear() + "-"
			+ (currentdate.getMonth()+1)  + "-" 
			+ currentdate.getDate()   
            +  "T"  
            + currentdate.getHours() + ":"  
            + currentdate.getMinutes() + ":" 
            + currentdate.getSeconds();
 return datetime;
}
function eventshide(e)
{
e.parentNode.parentNode.childNodes[1].childNodes[1].style.display = 'none';
e.parentNode.childNodes[1].style.display = 'block';
e.style.display = 'none';
}
function eventsshow(e)
{
e.parentNode.parentNode.childNodes[1].childNodes[1].style.display = 'block';
e.parentNode.childNodes[0].style.display = 'block';
e.style.display = 'none';
}
function showwpcblogevents(siteurl,listname,wptitle,htmlid,myQuery,expand)
{
var myQueryOptions = "";
var i = 1;
var itmsexist = 0;
var tstartdate;
var tdate;
var cdate = new Date();
  $().SPServices({
    operation: "GetListItems",
    webURL: siteurl,
    async: false,
    listName: listname,
    CAMLViewFields: "<ViewFields><FieldRef Name='KpiDescription' /><FieldRef Name='Created' /><FieldRef Name='Title' /><FieldRef Name='Event_x0020_Category' /></ViewFields>",
	CAMLQuery: myQuery,
	CAMLRowLimit: 100,
    completefunc: function (xData, Status) {
     // alert(xData.responseText);
		var liHtml = "<table width='100%' class='ms-viewlsts' id='#myTable'><thead><tr style='background-color: #0072c6; color: white; font-weight: bold;'><td height='30px' colspan='2'><b>"+wptitle+"</b></td></tr></thead><tbody>"
		responseXML = $.parseXML(xData.responseText);
		$XML = $(responseXML);
		var resultsLength = $XML.find('z\\:row, row').length;
		var arrayResults = [];
		var count = 0;		
		$XML.find('z\\:row, row').each(function(){     		
			itmsexist = 1;
	       try
			{
		//	alert($(this).attr("ows_Created"));
			var tcreated = $(this).attr("ows_Created");
  			var ttitle = $(this).attr("ows_Title");
  			var tdescription = $(this).attr("ows_KpiDescription");	
  			var teventcategory = $(this).attr("ows_Event_x0020_Category");	
  			var turl = siteurl + "/Lists/eventslog/WPCB%20Event/displayifs.aspx?IsDlg=1&ID=" + $(this).attr("ows_ID");				
					arrayResults.push({
							createdcolumn: tcreated ,
							titlecolumn: ttitle ,
							descriptioncolumn: tdescription ,
							linkcolumn: turl ,
							eventcategory: teventcategory ,
						});		
						count = count+1;
            }
   		catch(err)
   		{
   			alert("error: "+err.message);
   		}         		         	    			
     });
     		arrayResults.sort(function(a, b) {  
     				a = a["titlecolumn"]+a["createdcolumn"];
					b = b["titlecolumn"]+b["createdcolumn"];
					return a == b ? 0 : (a < b ? -1 : 1)
     		});
     	   var currtitle = "";
     	   var titlecount =  1;
    		for(i=0;i<count;i++)
     		{
     		 if( currtitle == "" || currtitle != arrayResults[i]["titlecolumn"])
     		 {
     			 if( currtitle != "")
     			 {
     			   titlecount += 1;
     			   liHtml += "</ui></td></tr>";
     			 }
     			    currtitle = arrayResults[i]["titlecolumn"];
     			    if((titlecount%2) == 0)
						{
							if(expand)
							{
							liHtml += "<tr class='ms-alternatingstrong'>";
       						liHtml += "<td><a style='display:block' onclick='eventshide(this)'><b>-</b></a><a style='display:none' onclick='eventsshow(this)'><b>+</b></a></td><td><ul><a onclick=\"javascript:SP.UI.ModalDialog.ShowPopupDialog('"+arrayResults[i]["linkcolumn"]+"')\"><b><u>"+arrayResults[i]["titlecolumn"]+"  - <b>["+arrayResults[i]["eventcategory"]+"]</b></u></b></a></ul><ul style='display:block' class='csevents'><li>"  + arrayResults[i]["descriptioncolumn"] + " <b> -["+arrayResults[i]["createdcolumn"] + "]</b></li>" ;
							}
							else
							{
							liHtml += "<tr class='ms-alternatingstrong'>";
       						liHtml += "<td><a style='display:none' onclick='eventshide(this)'><b>-</b></a><a style='display:block' onclick='eventsshow(this)'><b>+</b></a></td><td><ul><a onclick=\"javascript:SP.UI.ModalDialog.ShowPopupDialog('"+arrayResults[i]["linkcolumn"]+"')\"><b><u>"+arrayResults[i]["titlecolumn"]+"  - <b>["+arrayResults[i]["eventcategory"]+"]</b></u></b></a></ul><ul style='display:none' class='csevents'><li>"  + arrayResults[i]["descriptioncolumn"] + " <b> -["+arrayResults[i]["createdcolumn"] + "]</b></li>" ;	
							}
        				}
       				else
       					{
       						if(expand)
							{
							liHtml += "<tr style='background-color: white; color: #284775;'>";
       						liHtml += "<td><a style='display:block' onclick='eventshide(this)'><b>-</b></a><a style='display:none' onclick='eventsshow(this)'><b>+</b></a></td><td><ul><a onclick=\"javascript:SP.UI.ModalDialog.ShowPopupDialog('"+arrayResults[i]["linkcolumn"]+"')\"><b><u>"+arrayResults[i]["titlecolumn"]+"  - <b>["+arrayResults[i]["eventcategory"]+"]</b></u></b></a></ul><ul style='display:block' class='csevents'><li>"  + arrayResults[i]["descriptioncolumn"] + " <b> -["+arrayResults[i]["createdcolumn"] + "]</b></li>" ;
							}
							else
							{
							liHtml += "<tr style='background-color: white; color: #284775;'>";
       						liHtml += "<td><a style='display:none' onclick='eventshide(this)'><b>-</b></a><a style='display:block' onclick='eventsshow(this)'><b>+</b></a></td><td><ul><a onclick=\"javascript:SP.UI.ModalDialog.ShowPopupDialog('"+arrayResults[i]["linkcolumn"]+"')\"><b><u>"+arrayResults[i]["titlecolumn"]+"  - <b>["+arrayResults[i]["eventcategory"]+"]</b></u></b></a></ul><ul style='display:none' class='csevents'><li>"  + arrayResults[i]["descriptioncolumn"] + " <b> -["+arrayResults[i]["createdcolumn"] + "]</b></li>" ;	
							}         				
						} 
     		 }
     		 else
     		 {
     	        liHtml += 	"<li>"  + arrayResults[i]["descriptioncolumn"] + " <b> - ["+arrayResults[i]["createdcolumn"] + "]</b></li>";
     	        }
     				
     		}
 			if(itmsexist == 0)
				{
  					liHtml +="<tr class='ms-alternatingstrong'><td colspan='2' align='center'>No Events</td></tr>";
				}
			$("#"+htmlid).slideUp('slow');
    	 	$("#"+htmlid).html("");
    	 	$("#"+htmlid).slideDown('slow');
			$("#"+htmlid).append(liHtml);		
    }
 });
}