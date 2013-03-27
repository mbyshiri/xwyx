// **********************
// PowerEasy Cms2006
// userlogin part
// code by nt2003
// **********************

var siteroot;
var userstat;
var username;
var userid = 0;
var userpass;
var showtype;
var popmessage;
var messagecur = 0;
var messageold = 0;
var alogin = 1;

function LoadUserLogin(iroot,itype,iusepop)
{
    if(iroot == ''){
        siteroot="/";
    }else{
        siteroot=iroot;
    }
    if(itype == ''){
        showtype = 0;
    }else{
        showtype = parseInt(itype);
    }
    if(iusepop == ''){
        popmessage = 0;
    }else{
        popmessage = parseInt(iusepop);
    }
    ShowUserLogin();
}

function ShowUserLogin()
{ 
    var url = siteroot + "User/User_ChkLoginStatXml.asp";
    var pars = "action=xmlstat";
    var myAjax = new Ajax.Request(url, {method: 'post', parameters: pars, onComplete: ShowLoginForm, onFailure: reportError});
}

function ShowLoginForm(originalRequest)
{
    var xml; 
    if(window.ActiveXObject){
        xml = new ActiveXObject("Microsoft.XMLDOM");
        xml.async=false;
    } else {
        $('UserLogin').innerHTML = "<IFRAME id=\"UserLogin\" src=\""+ siteroot + "UserLogin.asp?ShowType=" + (showtype+1) + "\" frameBorder=\"0\" width=\"170\" scrolling=\"no\" height=\"145\"></IFRAME>";
		return;
    }
    xml.load(originalRequest.responseXml);
    var root = xml.getElementsByTagName("body");
    if(xml.readyState != 4 || root.length == 0){
        userstat = "notlogin";
        username = "";
        userpass = "";
    }else{
        var loginstat = root.item(0).getElementsByTagName("checkstat").item(0).text;
        if(loginstat=='err'){
            userstat = "notlogin";
            username = root.item(0).getElementsByTagName("user").item(0).text;
            userpass = "";
            if(root.item(0).getElementsByTagName("errsource").item(0).text==''){
                var tempstr = "<div id=\"loginerr\" style=\"display: none;color: red;background:#55FF88;text-align: center;height: 20;border: 1px solid #000000;\"></div>";
            }else{
                var tempstr = "<div id=\"loginerr\" style=\"color: red;background:#55FF88;text-align: center;height: 20;border: 1px solid #000000;\">" + root.item(0).getElementsByTagName("errsource").item(0).text + "</div>";
            }
            if(showtype==0){
                tempstr += "<table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">";
                tempstr += "<table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" class=\"userbox\">";
                tempstr += "<tr><td height=\"25\" align=\"right\"><span class=\"userlog\">&#x7528;&#x6237;&#x540D;&#xFF1A;</span></td><td height=\"25\" colspan=\"2\"><input name=\"UserName\" type=\"text\" id=\"UserName\" size=\"16\" maxlength=\"20\" style=\"width:110px;\"></td></tr>";
                tempstr += "<tr><td height=\"25\" align=\"right\"><span class=\"userlog\">&#x5BC6;&#x3000;&#x7801;&#xFF1A;</span></td><td height=\"25\"  colspan=\"2\"><input name=\"UserPassword\" type=\"password\" id=\"UserPassword\" size=\"16\" maxlength=\"20\" style=\"width:110px;\"></td></tr>";
                if(root.item(0).getElementsByTagName("checkcode").item(0).text=='1'){
                   tempstr += "<tr><td height=\"25\" align=\"right\"><span class=\"userlog\">&#x9A8C;&#x8BC1;&#x7801;&#xFF1A;</span></td><td height=\"25\"><input name=\"CheckCode\" type=\"text\" id=\"CheckCode\" size=\"6\" maxlength=\"6\" style=\"width:49px;\"></td><td><a href=\"javascript:refreshimg()\" title=\"&#x770B;&#x4E0D;&#x6E05;&#x695A;&#xFF0C;&#x6362;&#x4E2A;&#x56FE;&#x7247;\"><img id=\"checkcode\" src=\"" + siteroot + "inc/checkcode.asp\" style=\"border: 1px solid #ffffff\"></a></td></tr>";
                }

                tempstr += "</table><table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">";	
                tempstr += "<tr><td colspan=\"2\" align=\"center\">";
                tempstr += "<table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"5\" cellpadding=\"0\">";	
                tempstr += "<tr><td colspan=\"2\" align=\"center\"><input name=\"Login\" type=\"image\" id=\"Login\" src=\""+ siteroot +"Images/logins_01.gif\" style=\"width:45px;height:39px;border:0px;\" align=\"middle\" value=\" &#x767B; &#x5F55; \" onclick=\"CheckUser(" + root.item(0).getElementsByTagName("checkcode").item(0).text + ");\">&#x3000;<input type=\"checkbox\" name=\"CookieDate\" value=\"3\">&#x6C38;&#x4E45;&#x767B;&#x5F55;";
                tempstr += "</td></tr></table>";
                tempstr += "<table border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\">";
                tempstr += "<tr><td rowspan=\"2\"><img src=\""+ siteroot +"Images/loginr_01.gif\" alt=\"\"></td><td><a href=\""+ siteroot +"Reg/User_Reg.asp\" target=\"_blank\"><img src=\""+ siteroot +"Images/loginr_02.gif\" alt=\"&#x65B0;&#x7528;&#x6237;&#x6CE8;&#x518C;\" border=\"0\"></a></td></tr>";
                tempstr += "<tr><td><a href=\""+ siteroot +"User/User_GetPassword.asp\" target=\"_blank\"><img src=\""+ siteroot +"Images/loginr_03.gif\" alt=\"&#x5FD8;&#x8BB0;&#x5BC6;&#x7801;&#xFF1F;\" border=\"0\"></a></td></tr>";
                tempstr += "</table>";
                tempstr += "</tr></table>";
            }else{
               if(showtype==3)
               {
                tempstr += "<table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">";
                tempstr += "<tr><td align=\"right\"><font color=\"#FFFFFF\">&#x7528;&#x6237;&#x540D;</font></td><td><input name=\"UserName\" type=\"text\" id=\"UserName\" size=\"8\" maxlength=\"20\" style=\"width:40px;\"></td>";
                tempstr += "<td align=\"right\"><font color=\"#FFFFFF\">&#x5BC6;&#x7801;</font></td><td><input name=\"UserPassword\" type=\"password\" id=\"Password\" size=\"8\" maxlength=\"20\" style=\"width:40px;\"></td>";
                if(root.item(0).getElementsByTagName("checkcode").item(0).text=='1'){
                    tempstr += "<td align=\"right\"><font color=\"#FFFFFF\">&#x9A8C;&#x8BC1;&#x7801;</font></td>";
                    tempstr += "<td><input name=\"CheckCode\" type=\"text\" id=\"CheckCode\" size=\"8\" maxlength=\"6\" style=\"width:40px;\"></td><td><a href=\"javascript:refreshimg()\" title=\"&#x770B;&#x4E0D;&#x6E05;&#x695A;&#xFF0C;&#x6362;&#x4E2A;&#x56FE;&#x7247;\"><img id=\"checkcode\" src=\"" + siteroot + "inc/checkcode.asp\" style=\"border: 1px solid #ffffff\"></a></td>";
                }
                tempstr += "<td><input type=\"checkbox\" name=\"CookieDate\" value=\"3\"><font color=\"#FFFFFF\">&#x6C38;&#x4E45;&#x767B;&#x5F55;&#x3000;</font>";
                tempstr += "</td><td><input name=\"Login\" type=\"image\" id=\"Login\" src=\"" + siteroot + "Images/toplogin.gif\" value=\"\" onclick=\"CheckUser(" + root.item(0).getElementsByTagName("checkcode").item(0).text + ");\" style=\"width:45px;height:18px;\"></td><td><a href=\"" + siteroot + "Reg/User_Reg.asp\" target=\"_blank\"><font color=\"#FFFFFF\">&#x6CE8;&#x518C;</font></a> <a href=\"" + siteroot + "User/User_GetPassword.asp\" target=\"_blank\"><font color=\"#FFFFFF\">&#x5FD8;&#x8BB0;&#x5BC6;&#x7801;</font></a></td></tr></table>";
               }
               else{
                tempstr += "<table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">";
                tempstr += "<tr><td height=\"25\" align=\"right\">&#x7528;&#x6237;&#x540D;&#xFF1A;</td><td height=\"25\"><input name=\"UserName\" type=\"text\" id=\"UserName\" size=\"16\" maxlength=\"20\" style=\"width:110px;\"></td>";
                tempstr += "<td height=\"25\" align=\"right\">&#x5BC6;&#x3000;&#x7801;&#xFF1A;</td><td height=\"25\"><input name=\"UserPassword\" type=\"password\" id=\"Password\" size=\"16\" maxlength=\"20\" style=\"width:110px;\"></td>";
                if(root.item(0).getElementsByTagName("checkcode").item(0).text=='1'){
                    tempstr += "<td height=\"25\" align=\"right\">&#x9A8C;&#x8BC1;&#x7801;&#xFF1A;</td>";
                    tempstr += "<td height=\"25\"><input name=\"CheckCode\" type=\"text\" id=\"CheckCode\" size=\"6\" maxlength=\"6\"><a href=\"javascript:refreshimg()\" title=\"&#x770B;&#x4E0D;&#x6E05;&#x695A;&#xFF0C;&#x6362;&#x4E2A;&#x56FE;&#x7247;\"><img id=\"checkcode\" src=\"" + siteroot + "inc/checkcode.asp\" style=\"border: 1px solid #ffffff\"></a></td>";
                }
                tempstr += "<td height=\"25\" colspan=\"2\" align=\"center\"><input type=\"checkbox\" name=\"CookieDate\" value=\"3\">&#x6C38;&#x4E45;&#x767B;&#x5F55;&#x3000;&#x3000;";
                tempstr += "<input name=\"Login\" type=\"submit\" id=\"Login\" value=\" &#x767B; &#x5F55; \" onclick=\"CheckUser(" + root.item(0).getElementsByTagName("checkcode").item(0).text + ");\"></td><td height='25'><a href=\"" + siteroot + "Reg/User_Reg.asp\" target=\"_blank\">&#x65B0;&#x7528;&#x6237;&#x6CE8;&#x518C;</a>&#x3000;<a href=\"" + siteroot + "User/User_GetPassword.asp\" target=\"_blank\">&#x5FD8;&#x8BB0;&#x5BC6;&#x7801;&#xFF1F;</a></td></tr></table>";
                }
            }
            $('UserLogin').innerHTML = tempstr;
        }else{
            userstat = "login";
            username = root.item(0).getElementsByTagName("user").item(0).text;
            userid = root.item(0).getElementsByTagName("userid").item(0).text;
            userpass = root.item(0).getElementsByTagName("userpass").item(0).text;
            var plus_day = new Date( );
            var plus_hr= plus_day.getHours( );
            var timehello="hello"; 
            if (( plus_hr >= 0 ) && (plus_hr < 6 ))
            timehello = "<font color=\"#FF00FF\">&#x51CC;&#x6668;&#x597D;!</font>";
            if (( plus_hr >= 6 ) && (plus_hr < 9))
            timehello = "<font color=\"#FF00FF\">&#x65E9;&#x4E0A;&#x597D;!</font>";
            if (( plus_hr >= 9 ) && (plus_hr < 12))
            timehello = "<font color=\"#FF00FF\">&#x4E0A;&#x5348;&#x597D;!</font>";
            if (( plus_hr >= 12) && (plus_hr <14))
            timehello = "<font color=\"#FF00FF\">&#x4E2D;&#x5348;&#x597D;!</font>";
            if (( plus_hr >= 14) && (plus_hr <17))
            timehello = "<font color=\"#FF00FF\">&#x4E0B;&#x5348;&#x597D;!</font>";
            if (( plus_hr >= 17) && (plus_hr <18))
            timehello = "<font color=\"#FF00FF\">&#x508D;&#x665A;&#x597D;!</font>";
            if ((plus_hr >= 18) && (plus_hr <23))
            timehello = "<font color=\"#FF00FF\">&#x665A;&#x4E0A;&#x597D;!</font>";

            if(showtype==0){
                var tempstr = "<div id=\"userlogined\">";
                tempstr += "<font color=\"green\"><b>" + username + "</b></font>&#xFF0C;" + timehello;
                tempstr += "</div><div id=\"userlogined\">&#x8D44;&#x91D1;&#x4F59;&#x989D;&#xFF1A; <b><font color=\"blue\">" + root.item(0).getElementsByTagName("balance").item(0).text + "</font></b> &#x5143;";
                tempstr += "</div><div id=\"userlogined\">&#x7ECF;&#x9A8C;&#x79EF;&#x5206;&#xFF1A; <b><font color=\"blue\">" + root.item(0).getElementsByTagName("exp").item(0).text + "</font></b> &#x5206;";
                tempstr += "</div><div id=\"userlogined\">&#x53EF;&#x7528;" + root.item(0).getElementsByTagName("point/pointname").item(0).text + "&#xFF1A; <b><font color=\"gray\">" + root.item(0).getElementsByTagName("point/userpoint").item(0).text + "</font></b> " + root.item(0).getElementsByTagName("point/unit").item(0).text
                if(root.item(0).getElementsByTagName("day").item(0).text!='noshow'){
                    tempstr += "</div><div id=\"userlogined\">&#x5269;&#x4F59;&#x5929;&#x6570;&#xFF1A; <b><font color=\"blue\">";
                    if(root.item(0).getElementsByTagName("day").item(0).text=='unlimit'){
                        tempstr += "&#x65E0;&#x9650;&#x671F;";
                    }else{
                        tempstr += root.item(0).getElementsByTagName("day").item(0).text;
                    }
                }
                tempstr += "</font></b>";
                tempstr += "</div><div id=\"userlogined\">&#x5F85;&#x7B7E;&#x6587;&#x7AE0;&#xFF1A; <b><font color=\"gray\">" + root.item(0).getElementsByTagName("article").item(0).text + "</font></b> &#x7BC7;";
                if(root.item(0).getElementsByTagName("unreadmessage/stat").item(0).text=='full'){
                    tempstr += "</div><div id=\"usermessage\" class=\"havemessage\" onmouseover=\"havemessage();\" onmouseout=\"hidemessage();\" onclick=\"Element.toggle('messagelist');\" style=\"cursor:hand;\">&#x5F85;&#x9605;&#x77ED;&#x4FE1;&#xFF1A; <b><font color=\"gray\">" + root.item(0).getElementsByTagName("message").item(0).text + "</font></b> &#x6761;";
                    tempstr += "</div><div id=\"messagelist\" style=\"display:none\";>";
                    var messageloop = root.item(0).getElementsByTagName("unreadmessage/item");
                    var openurl;
                    for(i=0;i<messageloop.length;i++){
                        tempstr += "<li><a href=\"" + siteroot + "User/User_Message.asp?Action=ReadInbox&MessageID=" + messageloop.item(i).getElementsByTagName("id").item(0).text + "\" title=\"&#x6765;&#x81EA;&#xFF1A;" + messageloop.item(i).getElementsByTagName("sender").item(0).text + "\n&#x65F6;&#x95F4;&#xFF1A;" + messageloop.item(i).getElementsByTagName("time").item(0).text + "\">" + messageloop.item(i).getElementsByTagName("title").item(0).text + "</a></li>";
                    }
                }else{
                    tempstr += "</div><div id=\"userlogined\">&#x5F85;&#x9605;&#x77ED;&#x4FE1;&#xFF1A; <b><font color=\"gray\">" + root.item(0).getElementsByTagName("message").item(0).text + "</font></b> &#x6761;";
                }
                tempstr += "</div><div id=\"userlogined\">&#x767B;&#x5F55;&#x6B21;&#x6570;&#xFF1A; <b><font color=\"blue\">" + root.item(0).getElementsByTagName("logined").item(0).text + "</font></b> &#x6B21;";
                tempstr += "</div><div id=\"userctrl\"><a href=\"" + siteroot + "User/Index.asp\" target=\"ControlPad\">&#x3010;&#x4F1A;&#x5458;&#x4E2D;&#x5FC3;&#x3011;</a> <a href='#' onclick=\"UserLogout();\">&#x3010;&#x6CE8;&#x9500;&#x767B;&#x5F55;&#x3011;</a></div>";
            }else{
                if(showtype==3){
                    var tempstr = "<table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"2\" class=\"userlog\"><tr><td><font color=\"red\"><b>" + username + "</b></font>&#xFF0C;<span style=\"color:#ffff00;\">" + timehello + "</span></td>";
                    tempstr += "<td>&#x5F85;&#x7B7E;&#x6587;&#x7AE0;&#xFF1A;<b><font color=\"#ffff00\">" + root.item(0).getElementsByTagName("article").item(0).text + "</font></b> &#x7BC7;</td>";
                    tempstr += "<td>&#x5F85;&#x9605;&#x77ED;&#x4FE1;&#xFF1A;<b><font color=\"#ffff00\">" + root.item(0).getElementsByTagName("message").item(0).text + "</font></b> &#x6761;</td>";
                    tempstr += "<td>&#x767B;&#x5F55;&#x6B21;&#x6570;&#xFF1A;<b><font color=\"#ffff00\">" + root.item(0).getElementsByTagName("logined").item(0).text + "</font></b> &#x6B21;</td>";
                    tempstr += "<td><a href=\"" + siteroot + "User/Index.asp\" target=\"ControlPad\" class=\"Channel\">&#x3010;&#x4F1A;&#x5458;&#x4E2D;&#x5FC3;&#x3011;</a> <a href='#' onclick=\"UserLogout();\" class=\"Channel\">&#x3010;&#x6CE8;&#x9500;&#x3011;</a></td></tr></table>";
                }else{
                    var tempstr = "<table align=\"center\" width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"2\" ><tr><td>&#x3000;<font color=\"green\"><b>" + username + "</b></font>&#xFF0C;" + timehello + "</td>";
                    tempstr += "<td>&#x53EF;&#x7528;" + root.item(0).getElementsByTagName("point/pointname").item(0).text + "&#xFF1A; <b><font color=\"blue\">" + root.item(0).getElementsByTagName("point/userpoint").item(0).text + "</font></b></td>";
                    tempstr += "<td>&#x5F85;&#x7B7E;&#x6587;&#x7AE0;&#xFF1A;<b><font color=\"gray\">" + root.item(0).getElementsByTagName("article").item(0).text + "</font></b> &#x7BC7;</td>";
                    tempstr += "<td>&#x5F85;&#x9605;&#x77ED;&#x4FE1;&#xFF1A;<b><font color=\"gray\">" + root.item(0).getElementsByTagName("message").item(0).text + "</font></b> &#x6761;</td>";
                    tempstr += "<td>&#x767B;&#x5F55;&#x6B21;&#x6570;&#xFF1A;<b><font color=\"blue\">" + root.item(0).getElementsByTagName("logined").item(0).text + "</font></b> &#x6B21;</td>";
                    tempstr += "<td><a href=\"" + siteroot + "User/Index.asp\" target=\"ControlPad\">&#x3010;&#x4F1A;&#x5458;&#x4E2D;&#x5FC3;&#x3011;</a> <a href='#' onclick=\"UserLogout();\">&#x3010;&#x6CE8;&#x9500;&#x767B;&#x5F55;&#x3011;</a></td></tr></table>";
                }
            }
            $('UserLogin').innerHTML = tempstr;

            if(alogin==0)
            {
                var myAPIUrls = getAPIUrls(root,username,userpass)
                for (var i=0; i<myAPIUrls.length; i++)					
                {
                    var ifrm1 = document.createElement("IFRAME");
                    ifrm1.src = myAPIUrls[i];
                    ifrm1.height = "1";
                    ifrm1.width = "1";
                    ifrm1.frameborder= "0";
                    document.body.insertBefore(ifrm1);
                }
                alogin = 1;
            }
            if(popmessage==1){
                if(root.item(0).getElementsByTagName("unreadmessage/stat").item(0).text=='full'){
                    var messageurl;
                    var messloop = root.item(0).getElementsByTagName("unreadmessage/item");
                        messageurl = siteroot + "User/User_ReadMessage.asp?MessageID=" + messloop.item(0).getElementsByTagName("id").item(0).text;
                        window.open (messageurl, 'newmessage', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no');
                }
            }else if(popmessage==2){
                if(root.item(0).getElementsByTagName("grouptype").item(0).text > 1){
                    new PeriodicalExecuter(GetNewMessage,20);
                }
            }
        }
    }
}

function CheckUser(checktype)
{
    alogin = 0;
    var UserName = $F('UserName');
    var Password = $F('UserPassword');
    var CheckCode = ''; 
    if(checktype=='1'){
        CheckCode = $F('CheckCode');
    }else{
        var CheckCode = 0;
    }
    var CookieDate = $F('CookieDate');
    if(UserName==''){
        $('loginerr').innerHTML = "&#x8BF7;&#x586B;&#x5199;&#x7528;&#x6237;&#x540D;!";
        Element.show('loginerr');
        Field.focus('UserName');
    }else{
        if(Password==''){
            $('loginerr').innerHTML = "&#x8BF7;&#x586B;&#x5199;&#x5BC6;&#x7801;!";
            Element.show('loginerr');
            Field.focus('UserPassword');
        }else{
            if(checktype=='1' && CheckCode==''){
                $('loginerr').innerHTML = "&#x8BF7;&#x586B;&#x5199;&#x9A8C;&#x8BC1;&#x7801;!";
                Element.show('loginerr');
                Field.focus('CheckCode');
            }else{
                $('UserLogin').innerHTML = "&#x9A8C;&#x8BC1;&#x4E2D;...";
                var checkurl = siteroot + "User/User_ChkLoginXml.asp";

                // creat user xml file
                var xml_dom = new ActiveXObject("Microsoft.XMLDOM");
                xml_dom.async=false;
                var xmlproperty = xml_dom.createProcessingInstruction("xml","version=\"1.0\" encoding=\"gb2312\"");  
                xml_dom.appendChild(xmlproperty); 
                var objRoot = xml_dom.createElement("root");
                var objField = xml_dom.createNode(1,"username",""); 
                objField.text = UserName;
                objRoot.appendChild(objField);
                objField = xml_dom.createNode(1,"password",""); 
                objField.text = Password;
                objRoot.appendChild(objField);
                objField = xml_dom.createNode(1,"checkcode",""); 
                objField.text = CheckCode;
                objRoot.appendChild(objField);
                objField = xml_dom.createNode(1,"cookiesdate","");
                if(CookieDate>0){
                    objField.text = CookieDate;
                }
                objRoot.appendChild(objField);
                xml_dom.appendChild(objRoot);

                // send to server
                var userhttp = getHTTPObject();
                userhttp.open("POST",checkurl,false);
                userhttp.onreadystatechange = function () 
                {
	            if (userhttp.readyState == 4 && userhttp.status==200){
                       ShowLoginForm(userhttp);	
                   }else{
                       reportError();
	            }
                }
                userhttp.send(xml_dom);
            }
        }
    }
}

function GetNewMessage()
{ 
    var url = siteroot + "User/User_ChkLoginStatXml.asp";
    var pars = "action=xmlstat";
    var myAjax = new Ajax.Request(url, {method: 'get', parameters: pars, onComplete: ShowNewMessage});
}

function ShowNewMessage(originalRequest)
{
    var xml2 = new ActiveXObject("Microsoft.XMLDOM");
    xml2.async = false;
    xml2.load(originalRequest.responseXml);
    var root2 = xml2.getElementsByTagName("body/unreadmessage");
    var msgstat2 = root2.item(0).getElementsByTagName("stat").item(0).text;
    var messageloop2 = root2.item(0).getElementsByTagName("item");
    messagecur = messageloop2.length;
    if(messagecur != messageold){
        messageold = messagecur;
        ShowLoginForm(originalRequest);
    }
}

function havemessage()
{
    $('usermessage').className='havemessaged';
}

function hidemessage()
{
    $('usermessage').className='havemessage';
}

function UserLogout()
{
    var strTempHTML="";
    var dtime = 0;
    var outurl = siteroot + "User/User_Logout.asp?action=xml";
    var userhttp = getHTTPObject();
    userhttp.open("POST",outurl,false);
    userhttp.onreadystatechange = function () 
    {
        if (userhttp.readyState == 4) {
            if (userhttp.status==200){
                var xml; 
                xml = new ActiveXObject("Microsoft.XMLDOM");
                xml.async=false;
                xml.load(userhttp.responseXml);

                var root = xml.getElementsByTagName("body");
                if(root.length == 1){
                    var syskey = root.item(0).getElementsByTagName("syskey");
                    if (syskey.length == 1) {
                        var iUrls = root.item(0).getElementsByTagName("apiurl");
                        for (var i=0; i<iUrls.length; i++){
                            dtime = dtime + 2000;
                            strTempHTML += "<iframe frameborder=\"0\" width=\"1\" height=\"1\" src=\"" + iUrls.item(i).text + "?syskey=" + syskey.item(0).text + "&username=" + username + "\" \/>";
                        }
                        if (iUrls.length > 0) $('UserLogin').innerHTML = "logouting..." + strTempHTML;
                    }
                }
                var dd = setTimeout("ShowUserLogin()",dtime);
            }else{
                reportError();
            }
        }
    }
    userhttp.send();
}

function reportError()
{
    $('UserLogin').innerHTML = "<a href=\"#\" onclick=\"ShowUserLogin();\">&#x9519;&#x8BEF;,&#x670D;&#x52A1;&#x5668;&#x65E0;&#x54CD;&#x5E94;!</a>";
}

function refreshimg(){
  document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();
}

var glabelid;
var gvalue;
var gurl;
var gtime;
var dstat=0;

// *****************
// dynapage part 
// *****************
function ShowDynaPage(labelid,ipage,tflash,rootdir,value)
{
    var pagename = "dyna_page_" + labelid;
    $(pagename).innerHTML = "updateing...";
    gurl = rootdir + "dyna_page.asp";

    glabelid = labelid;
    gtime = tflash;

    // creat send xml file
    var dy_dom = new ActiveXObject("Microsoft.XMLDOM");
    dy_dom.async=false;
    var xmlproperty = dy_dom.createProcessingInstruction("xml","version=\"1.0\" encoding=\"gb2312\"");  
    dy_dom.appendChild(xmlproperty); 
    var objRoot = dy_dom.createElement("root");
    var objField = dy_dom.createNode(1,"id",""); 
    objField.text = labelid;
    objRoot.appendChild(objField);
    objField = dy_dom.createNode(1,"rootdir",""); 
    objField.text = rootdir;
    objRoot.appendChild(objField);
    objField = dy_dom.createNode(1,"page",""); 
    objField.text = ipage;
    objRoot.appendChild(objField);
    objField = dy_dom.createNode(1,"value","");
    objField.text = value;
    objRoot.appendChild(objField);
    dy_dom.appendChild(objRoot);
    gvalue = dy_dom;
    // sent to server
    var dyhttp = getHTTPObject();
    dyhttp.open("POST",gurl,false);
    dyhttp.onreadystatechange = function () 
    {
	if (dyhttp.readyState == 4 && dyhttp.status==200)
	{
        //$("dyna_body_" + labelid).innerHTML = dyhttp.responseText
        DynaPageResponse(dyhttp,labelid,tflash);		
	}
    }
    dyhttp.send(dy_dom);
    if(parseInt(tflash)>9){
        if(dstat==0){
            dstat=1;
            new PeriodicalExecuter(reFlashDynaPage,parseInt(tflash));
        }
    }
}

function reFlashDynaPage()
{
    var pagename1 = "dyna_page_" + glabelid;
    $(pagename1).innerHTML = "updateing...";

    // sent to server
    var fdyhttp = getHTTPObject();
    fdyhttp.open("POST",gurl,false);
    fdyhttp.onreadystatechange = function () 
    {
        if (fdyhttp.readyState == 4 && fdyhttp.status==200){
            DynaPageResponse(fdyhttp,glabelid,gtime);
        }	
    }
    fdyhttp.send(gvalue);
}

function DynaPageResponse(pageRequest,rid,rflash)
{
    var xml = new ActiveXObject("Microsoft.XMLDOM");
    xml.async = false;
    xml.load(pageRequest.responseXml);
    var tempdom = xml.getElementsByTagName("stat");
    var stat = tempdom.item(0).text;    
    if(stat=='err'){
        $("dyna_body_" + rid).innerHTML = xml.getElementsByTagName("infomation");
    }else{
        tempdom = xml.getElementsByTagName("id");
        var tid = tempdom.item(0).text;
        if(tid!=''){
            var temprootdir = xml.getElementsByTagName("rootdir").item(0).text;
            if(temprootdir == ''){ temprootdir = '\\'; }
            var tempcontent = xml.getElementsByTagName("content").item(0).text;
            if(tempcontent!=''){
                $("dyna_body_" + tid).innerHTML = tempcontent;
            }
            var temptotalpage = xml.getElementsByTagName("totalpage").item(0).text;
            if(temptotalpage == ''){ temptotalpage = '1'; }
            var tempcurrentpage = xml.getElementsByTagName("currentpage").item(0).text;
            if(tempcurrentpage == ''){ tempcurrentpage = '1'; }
            var temptotalitem = xml.getElementsByTagName("totalitem").item(0).text;
            if(temptotalitem == ''){ temptotalitem = '0'; }
            var tempvalue = xml.getElementsByTagName("value").item(0).text;
            GetPageList(tid,temprootdir,temptotalpage,tempcurrentpage,temptotalitem,tempvalue,0,rflash);
        }
    }
}

function GetPageList(t1,d1,p1,p2,p3,v1,m1,rt1)
{
    if(parseInt(p2)<1){
        p2=1;
    }
    if(p1>1){
        var temppage;
        if(m1==0){
            if(parseInt(p2)>1){
                temppage = "<img src=\"" + d1 + "Skin/blue/first.gif\" style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + ",1," + rt1 + ",'" + d1 + "','" + v1 + "');\">";
                temppage += " <img src=\"" + d1 + "Skin/blue/prev.gif\" style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + "," + (parseInt(p2)-1) + "," + rt1 + ",'" + d1 + "','" + v1 + "');\">";
            }else{
                temppage = "<img src=\"" + d1 + "Skin/blue/first_d.gif\">";
                temppage += " <img src=\"" + d1 + "Skin/blue/prev_d.gif\">";
            }
            var beginlog;
            var endlog;
            if(parseInt(p2)>5){
                beginlog = parseInt(p2)-4;
                temppage = temppage + ".";
            }else{
                beginlog = 1;
            }
            if((parseInt(p2)+4)<=p1){
                endlog = parseInt(p2)+4;
            }else{
                endlog = p1;
            }
            for (var i = beginlog; i <= endlog; i++) {
                if(parseInt(p2)==i){
                    temppage += " [<b><font color=red>" + i + "</font></b>] ";
                }else{
                    temppage += " <b style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + "," + i + "," + rt1 + ",'" + d1 + "','" + v1 + "');\">" + i + "</b> ";
                }
            }
            if((parseInt(p2)+4)<p1){
                temppage = temppage + ".";
            }
            if(parseInt(p2)<parseInt(p1)){
                temppage += "<img src=\"" + d1 + "Skin/blue/next.gif\" style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + "," + (parseInt(p2)+1) + "," + rt1 + ",'" + d1 + "','" + v1 + "');\">";
                temppage += " <img src=\"" + d1 + "Skin/blue/end.gif\" style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + "," + p1 + "," + rt1 + ",'" + d1 + "','" + v1 + "');\">";
            }else{
                temppage += "<img src=\"" + d1 + "Skin/blue/next_d.gif\">";
                temppage += " <img src=\"" + d1 + "Skin/blue/end_d.gif\">";
            }
        }else{
            if(parseInt(p2)>1){
                temppage = " <img src=\"" + d1 + "Skin/blue/prev.gif\" style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + "," + (parseInt(p2)-1) + "," + rt1 + ",'" + d1 + "','" + v1 + "');\">";
            }else{
                temppage = " <img src=\"" + d1 + "Skin/blue/prev_d.gif\">";
            }
            if(parseInt(p2)<p1){
                temppage += "<img src=\"" + d1 + "Skin/blue/next.gif\" style=\"cursor:hand;\" onclick=\"ShowDynaPage(" + t1 + "," + (parseInt(p2)+1) + "," + rt1 + ",'" + d1 + "','" + v1 + "');\">";
            }else{
                temppage += "<img src=\"" + d1 + "Skin/blue/next_d.gif\">";
            }
        }
        $("dyna_page_" + t1).innerHTML = temppage;
    }else{
        Element.hide("dyna_page_" + t1);
    }
}

//***************************
// xmlHTTPinit
//***************************
function getHTTPObject(){
    var xmlhttp_request = false;
    try{
        if( window.ActiveXObject ){
            for( var i = 5; i; i-- ){
                try{
                    if( i == 2 ){
                        xmlhttp_request = new ActiveXObject( "Microsoft.XMLHTTP" );
                    }else{
                        xmlhttp_request = new ActiveXObject( "Msxml2.XMLHTTP." + i + ".0" );
                        xmlhttp_request.setRequestHeader("Content-Type","text/xml");
                        xmlhttp_request.setRequestHeader("Content-Type","gb2312");
                    }
                    break;
                }catch(e){
                    xmlhttp_request = false;
                }
            }
        }else if( window.XMLHttpRequest ){
            xmlhttp_request = new XMLHttpRequest();
            if (xmlhttp_request.overrideMimeType) {
                xmlhttp_request.overrideMimeType('text/xml');
            }
        }
    }catch(e){
        xmlhttp_request = false;
    }
    return xmlhttp_request ;
}

//***************************
//cont for Visitor part
//***************************
function addfangke(ibid,idir)
{
   // alert(username);
    if(userstat == 'login'){
        if(idir==0){
            var fangurl = "index.asp?action=addfang";
        }else{
            var fangurl = idir + "/index.asp?action=addfang";
        }
        var fang_dom = new ActiveXObject("Microsoft.XMLDOM");
        fang_dom.async=false;
        var pfang = fang_dom.createProcessingInstruction("xml","version=\"1.0\" encoding=\"gb2312\""); 
        fang_dom.appendChild(pfang); 
        var fangRoot = fang_dom.createElement("root");

        var fangField = fang_dom.createNode(1,"blogid",""); 
        fangField.text = ibid;
        fangRoot.appendChild(fangField);
        fangField = fang_dom.createNode(1,"username",""); 
        fangField.text = username;
        fangRoot.appendChild(fangField);
        fangField = fang_dom.createNode(1,"userid",""); 
        fangField.text = userid;
        fangRoot.appendChild(fangField);

        fang_dom.appendChild(fangRoot);

        var VHttp = getHTTPObject();
        VHttp.open("POST",fangurl,false);
        VHttp.send(fang_dom);
    }
}

//***************************
//PDOaip part
//***************************
function getAPIUrls(root,username,userpass){
    var strTempHTML = "";
    var iName,iPass;
    var syskey = root.item(0).getElementsByTagName("syskey").item(0).text;
    var savecookie = root.item(0).getElementsByTagName("savecookie").item(0).text;
    if (savecookie != "") {
        savecookie = "&savecookie=" + savecookie;
    }else{
        savecookie = "&savecookie=";
    }
    if (syskey != "" && username != "") {
        iName = "&username=" + username;
        if (userpass != "") {
            iPass = "&password=" + userpass;
        }else{
            iPass = "&password=";
        }
        var iUrls = root.item(0).getElementsByTagName("apiurl");
        for (var i=0; i<iUrls.length; i++){
            strTempHTML += iUrls.item(i).text + "?syskey=" + syskey + iName + iPass + savecookie + "|";
        }
    }
	var strTempHTML = strTempHTML.substr(0, strTempHTML.length-1);
	var strTempHTML = strTempHTML.split("|");
    return strTempHTML;
}