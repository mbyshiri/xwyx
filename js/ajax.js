function createAjaxObj(){
 var httprequest=false
 if (window.XMLHttpRequest){ // if Mozilla, Safari etc
  httprequest=new XMLHttpRequest()
  if (httprequest.overrideMimeType)
   httprequest.overrideMimeType('text/xml');
 }
 else if (window.ActiveXObject){ // if IE
  try 
  {
   httprequest=new ActiveXObject("Msxml2.XMLHTTP");
  } 
  catch (e)
  {
   try
   {
    httprequest=new ActiveXObject("Microsoft.XMLHTTP");
   }
   catch (e){}
  }
 }
 return httprequest;
}


function load_Hits(iroot,arrID){
 var xmlhttp = createAjaxObj();
 var siteroot;
  if(iroot == ''){
        siteroot="/";
    }else{
        siteroot=iroot;
    }
 try
 {

  var params="HitsType=0&SoftID="+arrID;
  xmlhttp.abort(); 
  
  xmlhttp.open("get",siteroot+"/Soft/GetHits.asp?"+params,true);
 
  xmlhttp.setRequestHeader("Content-type", "text/html;charset=gb2312"); 
  
  xmlhttp.setRequestHeader("If-Modified-Since","0"); 
 
  xmlhttp.setRequestHeader("Content-length", params.length);
  
  xmlhttp.setRequestHeader("Connection", "close");
 
  xmlhttp.onreadystatechange=f

  xmlhttp.send(null); 
 
 }catch(ex){alert(ex)}
 function f()
 {
  
   if(xmlhttp.readyState!= 4 || xmlhttp.status!=200 )
    return ;
   var b= xmlhttp.responseText;
   document.getElementById('Hits').innerHTML="";
   document.getElementById('Hits').innerHTML=b;
 }
}
