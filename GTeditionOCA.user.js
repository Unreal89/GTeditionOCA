// ==UserScript==
// @name         GT Edition script for OCA
// @namespace    http://tampermonkey.net/
// @version      0.9822
// @description  OCA custom script adding features https://hpe.sharepoint.com/teams/TecHub_EMEA/Team/wiki/OCA%20GT%20Edition.aspx
// @author       Global TecHub
// will be applied only on below sites
// @match        https://ngc-pro-oca-internal.houston.hp.com/oca/OCAInternalLogin
// @match        https://ngc-pro-ocac2-internal.houston.hp.com/ocacluster2/OCAInternalLogin
// @match        https://ngc-pro-oca-internal.bbn.hpecorp.net:1181/oca/OCAInternalLogin
// @match        https://h22246.www2.hpe.com/oca/OCAInternalLogin
// @match        https://ngc-pro-oca-internal.sgp.hpecorp.net:1181/oca/OCAInternalLogin
// @match        https://ngc-pro-oca-internal.in.hpecorp.net/oca/OCAInternalLogin
// this is the variable to store css
// @resource     customCSS https://github.hpe.com/GlobalTecHub/GTeditionOCA/raw/master/black.css?v=0.9823
// @require https://github.hpe.com/GlobalTecHub/GTeditionOCA/raw/master/powerusers.js?v=0.9823
// @require https://github.hpe.com/GlobalTecHub/GTeditionOCA/raw/master/notsopowerusers.js?v=0.9823
// @require https://github.hpe.com/GlobalTecHub/GTeditionOCA/raw/master/pimpedexport.js?v=0.9823
// @require https://github.hpe.com/GlobalTecHub/GTeditionOCA/raw/master/exceljs.js?v=0.9823
// @grant       GM_getResourceText
// ==/UserScript==

(function() {
    $( "body" ).keypress(function( event ) {
        if (event.which == 69) {
            $("#edit_mode.option_label").click()
            setTimeout(function () {$("#extended_overview_menu").data().menu_widget_data.edit_mode=true}, 500)
        }
    })
    checkmauiagree=function () {
        if (!maui.blockUI.isWaiting()){// wait until loading dissapeares
            clearInterval(checkMAagree); //stop checking
            //this is the general message that contains blah blah and links to code
            displayHTMLInModalDialogDefault('Do you want your OCA to be modified?',
                                            '<div class="bluecolor">You have successfully installed OCA GT Edition script. You can start using it by clicking "I Agree" button here.</div>' +
                                             '<div class="bluecolor">By clicking on it you agree that Tampermonkey will alter appearance of OCA and add custom scripts to ease your work.</div>' +
                                             '<div class="bluecolor">This script does not collect, send nor analyze your interaction with the OCA tool.</div>' +
                                             '<br><div class="bluecolor">Should you want to review the source code of injected scripts, visit following links:</div>' +
                                             '<div>Main script: <a target="_blank" href="https://github.hpe.com/GlobalTecHub/GTeditionOCA/blob/master/pimpedexport.js">https://github.hpe.com/GlobalTecHub/PimpMyOCA/blob/master/pimpedexport.js</a></div>'+
                                             '<div>Export script: <a target="_blank" href="https://github.hpe.com/GlobalTecHub/GTeditionOCA/blob/master/pimpedexport.js">https://github.hpe.com/GlobalTecHub/PimpMyOCA/blob/master/pimpedexport.js</a></div>'+
                                             '<div>custom CSS styles: <a target="_blank" href="https://github.hpe.com/GlobalTecHub/GTeditionOCA/blob/master/black.css">https://github.hpe.com/GlobalTecHub/PimpMyOCA/blob/master/black.css</a></div>'+
                                             '<div>ExcelJS plugin: <a target="_blank" href="https://github.com/guyonroche/exceljs">https://github.com/guyonroche/exceljs</a></div>'+
                                             '<div>Download plugin: <a target="_blank" href="https://github.com/rndme/download">https://github.com/rndme/download</a></div>'+
                                             '<button onclick="enablePimping()">I agree</button>',800,800);
        }
    };
    if (!localStorage.getItem("letsPimp")) {// letsPimp is the variable which saves user consent
        localStorage.setItem("letsPimp",false);//if there is none, let's create it and set it to false
    }
    letsPimp=localStorage.getItem("letsPimp");//get it normal variable

    enablePimping = function(){//the fun starts
    localStorage.setItem("letsPimp",true);//enable pimping
    $(".ui-button-text").click();//hide the current message
    pimping();//inject everything
        //display message that informs user that pimping finished
    displayHTMLInModalDialogDefault('Thank you for using GT ediation for OCA!','<div class="bluecolor">If you want to disable the script, click on Tampermonkey icon next to adress bar of your browser and toggle Pimp My OCA. You will need to refresh the page so save your work first.</div><div class="bluecolor">Should you want to withdraw previous consent about using this script, just clear your browser cookies.</div>',800,800);
    };
    pimping = function() {
        pimpedColor=localStorage.getItem('pimpedColor')
        if (pimpedColor=='light')
        {
        
        } else {
        $('body').addClass('darkGUI');//default color theme
        }
        //$.fx.off=true; // turns off those ugly transition effects. Quick bug fix
        css = GM_getResourceText("customCSS"); //css stores injected styles
        head = document.body; //where to inject styles
        //preparing style object
        style = document.createElement('style');
        style.id = 'DarkGUI';
        style.type = 'text/css';
        if (style.styleSheet){style.styleSheet.cssText = css;} else {style.appendChild(document.createTextNode(css));} //pasting css into object
        head.appendChild(style); //injecting object

        //script names
        var addGUI = document.createElement('script');
        var removeGUI = document.createElement('script');
        var openByUCID = document.createElement('script');
        var pimped_summary = document.createElement('script');
        var pimped_summary2 = document.createElement('script');
        var pimped_export = document.createElement('script');

        //I Injected excel.js library and download.js library plus our pimped excel lib
        //my_awesome_script = document.createElement('script');
        //my_awesome_script.setAttribute('src', 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/1.1.3/exceljs.js');
        //document.head.appendChild(my_awesome_script);
        my_awesome_script2 = document.createElement('script');
        my_awesome_script2.setAttribute('src', 'https://cdnjs.cloudflare.com/ajax/libs/downloadjs/1.4.7/download.js');
        document.head.appendChild(my_awesome_script2);

        //this is used to enable dark theme
        addGUI.innerHTML = "function DarkON(){var list = document.getElementById('DarkGUI'); list.innerHTML=css;$('#DarkOFFButton').show();$('#DarkONButton').hide();}";
        //this is used to disable dark theme
        removeGUI.innerHTML = "function DarkOFF(){var list = document.getElementById('DarkGUI'); list.innerHTML='';$('#DarkONButton').show();$('#DarkOFFButton').hide();};function toggleGUI(){if($('.darkGUI').length>0){$('body').removeClass('darkGUI');localStorage.setItem('pimpedColor','light')}else{$('body').addClass('darkGUI');localStorage.setItem('pimpedColor','dark')}}";
        //open by ucid funtion
        openByUCID.innerHTML = "function openByUCIDf(){var myucid = prompt('Enter UCID of configuration you want to open. Current quote WILL NOT BE SAVED.', '');if (myucid != null) {show_recent_config(myucid)}}";
        //pimped summary uses BoM view data. it recursively creates something like a BoM but without PNs. It skips the lines in which SKU has space in it.
        pimped_summary.innerHTML = "function pimpedSummary(){aaa=getServerData({'method':'get_tab_content','tabId': 'bom','widget_id':'extended_overview','selectedKey': 'root','menu_prefs' :localStorage.getItem('menu_prefs')}).data.subconfigs[0].bom.subnodes;tralala = function (arr,q,d) {	var temp='';	var d;	for (var i=0;i<arr.length;i++) {		if (!arr[i].attributes.product_number.includes(' ')){			temp=temp+Array(d).join('&emsp;')+(arr[i].attributes.quantity/q)+ 'x '+arr[i].description+'<br>'+tralala(arr[i].subnodes,arr[i].attributes.quantity,d+1)		};	};	return temp;};displayHTMLInModalDialog('Summary 1',$('#eo_nav_li_0')[0].outerText+'<br>UCID:'+$('.config-ucid')[0].outerText+'<br>'+tralala(aaa,1,1),800,1000);}";
        //even more pimped summary uses layout view data
        pimped_summary2.innerHTML = "function pimpedSummary2(){bbb=$('#theData').data().mainPageData.pageData.widgets.extended_overview.root.children;tralala2 = function (arr,q,d) {var temp='';var d;for (var i=0;i<arr.length;i++) {if (true){temp=temp+'<b>'+Array(d).join('&emsp;')+arr[i].subconfig_name+(arr[i].summary ? (', '+arr[i].summary):'') +'</b><br>'+Array(d).join('&emsp;')+(arr[i].qty)+ 'x '+arr[i].description+'<br>'+tralala2(arr[i].children,arr[i].qty,d+1)}};return temp};displayHTMLInModalDialog('Alternative Summary',tralala2(bbb,1,1),800,1000)}";
        //running pimped excel function, pimpedexceljs() is in my_awesome_script3
        pimped_export.innerHTML = "function pimpedExport() {setTimeout(function () {pimpedexcel(), 100});}";

        //injecting scripts
        document.body.appendChild(addGUI);
        document.body.appendChild(removeGUI);
        document.body.appendChild(openByUCID);
        document.body.appendChild(pimped_summary);
        document.body.appendChild(pimped_summary2);
        document.body.appendChild(pimped_export);
        //this set 100ms interval to check if ui is blocked
        checkMA=setInterval(function(){checkmaui();},100);
        checkmaui=function () {
            if (!maui.blockUI.isWaiting())
                //if ui is not blocked, stop checking and continue
            {clearInterval(checkMA);
             //place to insert buttons
             var userright = document.getElementById('dataWrap');
             //button to turn on dark mode
             buttonON=document.createElement('button');
             buttonON.type='button';
             buttonON.style.width='28px';
             buttonON.style.height='28px';
             buttonON.id='DarkONButton';
             buttonON.style.position='absolute';
             buttonON.style.zIndex='9999';
             buttonON.addEventListener("click",function (){DarkON();});
             buttonON.innerHTML="<div style='font-size: 12px;margin-left: -6px'>ON</div>";
             buttonON.title="Switch to dark theme";
             //button to turn off dark mode
             buttonOFF=document.createElement('button');
             buttonOFF.type='button';
             buttonOFF.id='DarkOFFButton';
             buttonOFF.style.width='28px';
             buttonOFF.style.height='28px';
             buttonOFF.style.position='absolute';
             buttonOFF.style.zIndex='9999';
             buttonOFF.addEventListener("click",function (){DarkOFF();});
             buttonOFF.innerHTML="<div style='font-size: 12px;margin-left: -6px'>OFF</div>";
             buttonOFF.title="Switch to light theme";
             //button to open quote by ucid
             buttonOPEN=document.createElement('button');
             buttonOPEN.addEventListener("click",function (){openByUCIDf();});
             buttonOPEN.type='button';
             buttonOPEN.style.position='absolute';
             buttonOPEN.style.top='28px';
             buttonOPEN.style.width='28px';
             buttonOPEN.style.height='28px';
             buttonOPEN.style.zIndex='9999';
             buttonOPEN.title="Open directly quote using UCID";
             buttonOPEN.innerHTML="<div style='font-size: 35px;margin-left: -9px;margin-top: -14px'>&#9758</div>";
             //button to create summary
             buttonSummary=document.createElement('button');
             buttonSummary.style.position='absolute';
             buttonSummary.style.top='56px';
             buttonSummary.style.width='28px';
             buttonSummary.style.height='28px';
             buttonSummary.style.zIndex='9999';
             buttonSummary.addEventListener("click",function (){pimpedSummary2();pimpedSummary();});
             buttonSummary.title="Show Pimped Summary";
             buttonSummary.innerHTML="<div style='font-size: 35px;margin-left: -9px;margin-top: -14px'>&#9776</div>";
             //button to export
             buttonExport=document.createElement('button');
             buttonExport.type='button';
             buttonExport.style.position='absolute';
             buttonExport.style.left='28px';
             buttonExport.style.width='28px';
             buttonExport.style.height='28px';
             buttonExport.style.zIndex='9999';
             //buttonExport.addEventListener("click",function (){pimpedExport();});
             buttonExport.innerHTML="E";
             buttonExport.title="Export Excel, xml and oca file.";
             buttonExport.id="ButtonExport"
             //button to lighten
             buttonLighten=document.createElement('button');
             buttonLighten.type='button';
             buttonLighten.style.position='absolute';
             buttonLighten.style.top='28px';
             buttonLighten.style.left='28px';
             buttonLighten.style.width='28px';
             buttonLighten.style.height='28px';
             buttonLighten.style.zIndex='9999';
             buttonLighten.addEventListener("click",function (){toggleGUI();});
             buttonLighten.innerHTML="&#9733";
             buttonLighten.title="GUI modification without color scheme";
             //inject buttons
             userright.insertBefore(buttonON,userright.childNodes[0]);
             userright.insertBefore(buttonOFF,userright.childNodes[0]);
             userright.insertBefore(buttonOPEN,userright.childNodes[0]);
             userright.insertBefore(buttonSummary,userright.childNodes[0]);
             if ($.inArray($("#theData").data().homePageData.pageData.user.email,powerusers)>-1)
             {userright.insertBefore(buttonExport,userright.childNodes[0]);}
             if ($.inArray($("#theData").data().homePageData.pageData.user.email,notsopowerusers)>-1)
             {userright.insertBefore(buttonExport,userright.childNodes[0]);}
             userright.insertBefore(buttonLighten,userright.childNodes[0]);
             $("#DarkONButton").hide(); //initially the on button is hidden
             $("#ButtonExport").click(function(e) {
                 if (e.shiftKey) {
                     pimpedexcel(true);
                 }
                 else {
                     pimpedexcel()
                 }
             });
            }
        };
    };
    if (letsPimp=="false") {
        checkMAagree=setInterval(function(){checkmauiagree();},100);
    } else {
        pimping();
    }
})();
