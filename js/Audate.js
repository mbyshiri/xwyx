function getlink(action,name,link) {
	monthnames = new Array(
	"一月",
	"二月",
	"三月",
	"四月",
	"五月",
	"六月",
	"七月",
	"八月",
	"九月",
	"十月",
	"十一月",
	"十二月"); 
	var linkcount=0;
	linkdays = new Array();
	monthdays = new Array(12);
	monthdays[0]=31;
	monthdays[1]=28;
	monthdays[2]=31;
	monthdays[3]=30;
	monthdays[4]=31;
	monthdays[5]=30;
	monthdays[6]=31;
	monthdays[7]=31;
	monthdays[8]=30;
	monthdays[9]=31;
	monthdays[10]=30;
	monthdays[11]=31;
	todayDate=new Date();
	thisday=todayDate.getDay();
	thismonth=todayDate.getMonth();
	thisdate=todayDate.getDate();
	thisyear=todayDate.getYear();
	thisyear = thisyear % 100;
	thisyear = ((thisyear < 50) ? (2000 + thisyear) : (1900 + thisyear));
	if (((thisyear % 4 == 0) && !(thisyear % 100 == 0))||(thisyear % 400 == 0)) monthdays[1]++;
	startspaces=thisdate;
	while (startspaces > 7) startspaces-=7;
	startspaces = thisday - startspaces + 1;
	if (startspaces < 0) startspaces+=7;
	document.write("<table border=0 cellpadding=2 cellspacing=1");
	document.write("><font color=black>");
	document.write("<tr><td colspan=7><center><< " 
	+"今天是："+ thisyear 
	+"年 "+monthnames[thismonth]+" >></center></font></td></tr>");
	document.write("<tr>");
	document.write("<td align=center>日</td>");
	document.write("<td align=center>一</td>");
	document.write("<td align=center>二</td>");
	document.write("<td align=center>三</td>");
	document.write("<td align=center>四</td>");
	document.write("<td align=center>五</td>");
	document.write("<td align=center>六</td>"); 
	document.write("</tr>");
	document.write("<tr>");
	for (s=0;s<startspaces;s++) {
		document.write("<td>&nbsp</td>");
	}
	count=1;
	while (count <= monthdays[thismonth]) {
		for (b = startspaces;b<7;b++) {
			document.write("<td");
			if (count==thisdate) {
				document.write(" bgcolor=red");
			}
			document.write(">");
			if (count <= monthdays[thismonth]) {
				document.write("<a href='" + action + ".asp?ChannelID=" + name + "&AuthorName=" + link + "&Data=" + thisyear + "-" + (thismonth + 1) +"-"+ count + "'>" + count + "</a>" );
			}
			else {
				document.write("&nbsp");
			}
			document.write("</td>");
			count++;
		}
		document.write("</tr>");
		document.write("<tr>");
		startspaces=0;
	}
	document.write("</tr></table>");
}