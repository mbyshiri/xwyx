function ObjectAD() {
  /* Define Variables*/
  this.ADID        = 0;
  this.ADType      = 0;
  this.ADName      = "";
  this.ImgUrl      = "";
  this.ImgWidth    = 0;
  this.ImgHeight   = 0;
  this.FlashWmode  = 0;
  this.LinkUrl     = "";
  this.LinkTarget  = 0;
  this.LinkAlt     = "";
  this.Priority    = 0;
  this.CountView   = 0;
  this.CountClick  = 0;
  this.InstallDir  = "";
  this.ADDIR       = "";
}

function CoupletZoneAD(_id) {
  /* Define Common Variables*/
  this.ID          = _id;
  this.ZoneID      = 0;
  this.ZoneName    = "";
  this.ZoneWidth   = 0;
  this.ZoneHeight  = 0;
  this.ShowType    = 1;
  this.DivName     = "";
  this.Div1        = null;
  this.Div2        = null;

  /* Define Unique Variables*/
  this.Left        = 10;
  this.Right       = 10;
  this.Top         = 50;
  this.Delta       = 0.15;

  /* Define Objects */
  this.AllAD       = new Array();
  this.ShowAD      = null;

  /* Define Functions */
  this.AddAD       = CoupletZoneAD_AddAD;
  this.GetShowAD   = CoupletZoneAD_GetShowAD;
  this.Show        = CoupletZoneAD_Show;
  this.Move        = CoupletZoneAD_Move;

}

function CoupletZoneAD_AddAD(_AD) {
  this.AllAD[this.AllAD.length] = _AD;
}

function CoupletZoneAD_GetShowAD() {
  if (this.ShowType > 1) {
    this.ShowAD = this.AllAD[0];
    return;
  }
  var num = this.AllAD.length;
  var sum = 0;
  for (var i = 0; i < num; i++) {
    sum = sum + this.AllAD[i].Priority;
  }
  if (sum <= 0) {return ;}
  var rndNum = Math.random() * sum;
  i = 0;
  j = 0;
  while (true) {
    j = j + this.AllAD[i].Priority;
    if (j >= rndNum) {break;}
    i++;
  }
  this.ShowAD = this.AllAD[i];
}

function CoupletZoneAD_Show() {
  if (!this.AllAD) {
    return;
  } else {
    this.GetShowAD();
  }

  if (this.ShowAD == null) return false;
  this.DivName = "CoupletZoneAD_Div" + this.ZoneID;
  if (!this.ShowAD.ImgWidth) this.ShowAD.ImgWidth = this.ZoneWidth
  if (!this.ShowAD.ImgHeight) this.ShowAD.ImgHeight = this.ZoneHeight
  if (this.ShowAD.ADDIR=="") this.ShowAD.ADDIR = "AD"

  document.write("<div id='" + this.DivName + "L" + "' style='position:absolute; visibility:visible; z-index:1; width:" + this.ZoneWidth + "px; height:" + this.ZoneHeight + "px; left:" + this.Left + ";top:" + this.Top + "'>" + AD_Content(this.ShowAD) + "<div style='cursor:pointer;float:left;text-align:left;width:" + this.ZoneWidth + "px;height:20px;left:" + this.Left + ";' onclick='javascript:document.getElementById(\"" + this.DivName + "L" + "\").style.display = \"none\";document.getElementById(\"" + this.DivName + "R" + "\").style.display = \"none\";'>¡Á¹Ø±Õ</div></div>");
  document.write("<div id='" + this.DivName + "R" + "' style='position:absolute; visibility:visible; z-index:1; width:" + this.ZoneWidth + "px; height:" + this.ZoneHeight + "px; right:" + this.Right + ";top:" + this.Top + "'>" + AD_Content(this.ShowAD) + "<div style='cursor:pointer;float:left;text-align:right;width:" + this.ZoneWidth + "px;height:20px;right:" + this.Right + ";' onclick='javascript:document.getElementById(\"" + this.DivName + "L" + "\").style.display = \"none\";document.getElementById(\"" + this.DivName + "R" + "\").style.display = \"none\";'>¡Á¹Ø±Õ</div></div>");

  if (this.ShowAD.CountView) {
    document.write ("<script src='" + this.ShowAD.InstallDir + this.ShowAD.ADDIR + "/ADCount.asp?Action=View&ADID=" + this.ShowAD.ADID + "'></script>")
  }
  this.Div1 = document.getElementById(this.DivName + "L");
  this.Div2 = document.getElementById(this.DivName + "R");
  setInterval(this.ID + ".Move()", 10);
}

function CoupletZoneAD_Move() {
  this.Div1.style.top=document.documentElement.scrollTop+10;
  this.Div1.style.left=this.Left;
  this.Div2.style.top=document.documentElement.scrollTop+10;
  this.Div2.style.right=this.Right;
  //this.Div1.style.display = '';
  //this.Div2.style.display = '';
}

function AD_Content(o) {
  var str = "";
  if (o.ADType == 1 || o.ADType == 2) {
  imgurl = o.ImgUrl .toLowerCase()
    if (o.InstallDir.indexOf("http://") != - 1) imgurl = o.InstallDir.substr(0, o.InstallDir.length - 1) + imgurl;
    if (imgurl.indexOf(".swf") !=  - 1) {
      str = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0'";
      str += " name='AD_" + o.ADID + "' id='AD_" + o.ADID + "'";
      str += " width='" + o.ImgWidth + "'";
      str += " height='" + o.ImgHeight + "'";
      if (o.style) str += " style='" + o.style + "'";
      if (o.extfunc) str += " " + o.extfunc + " ";
      str += ">";
      str += "<param name='movie' value='" + imgurl + "'>";
      if (o.FlashWmode == 1) str += "<param name='wmode' value='Transparent'>";
      if (o.play) str += "<param name='play' value='" + o.play + "'>";
      if (typeof(o.loop) != "undefined") str += "<param name='loop' value='" + o.loop + "'>";
      str += "<param name='quality' value='autohigh'>";
      str += "<embed ";
      str += " name='AD_" + o.ADID + "' id='AD_" + o.ADID + "'";
      str += " width='" + o.ImgWidth + "'";
      str += " height='" + o.ImgHeight + "'";
      if (o.style) str += " style='" + o.style + "'";
      if (o.extfunc) str += " " + o.extfunc + " ";
      str += " src='" + imgurl + "'";
      if (o.FlashWmode == 1) str += " wmode='Transparent'";
      if (o.play) str += " play='" + o.play + "'";
      if (typeof(o.loop) != "undefined") str += " loop='" + o.loop + "'";
      str += " quality='autohigh'"
      str += " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>";
      str += "</object>";
    } else if (imgurl.indexOf(".gif") !=  - 1 || imgurl.indexOf(".jpg") !=  - 1 || imgurl.indexOf(".jpeg") !=  - 1 || imgurl.indexOf(".bmp") !=  - 1 || imgurl.indexOf(".png") !=  - 1) {
      if (o.LinkUrl) {
        if (o.CountClick) o.LinkUrl = o.InstallDir + o.ADDIR + "/ADCount.asp?Action=Click&ADID=" + o.ADID
        str += "<a href='" + o.LinkUrl + "' target='" + ((o.LinkTarget == 0) ? "_self" : "_blank") + "' title='" + o.LinkAlt + "'>";
      }
      str += "<img ";
      str += " name='AD_" + o.ADID + "' id='AD_" + o.ADID + "'";
      if (o.style) str += " style='" + o.style + "'";
      if (o.extfunc) str += " " + o.extfunc + " ";
      str += " src='" + imgurl + "'";
      if (o.ImgWidth) str += " width='" + o.ImgWidth + "'";
      if (o.ImgHeight) str += " height='" + o.ImgHeight + "'";
      str += " border='0'>";
      if (o.LinkUrl) str += "</a>";
    }
  } else if (o.ADType == 3 || o.ADType == 4) {
    str = o.ADIntro
  } else if (o.ADType == 5) {
    str = "<iframe id='" + "AD_" + o.ADID + "' marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no width=100% height=100% src='" + o.ADIntro + "'>wait</iframe>";
  }
  return str;
}

