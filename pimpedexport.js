pimpedexcel = function (sdd) {
    if (unsaveDirtyFlag()) { //if quote is not saved
        displayHTMLInModalDialog('Save the quote first!', 'Quote was not yet saved on server. Please save it.', 500, 500)
    } else {
        $('#theData').save_config('load_dialog', {type: 'save_to_local', isRoot: true}); //this saves oca file
        if ($("#theData").data().homePageData.pageData.user.email=="david.firbach@hpe.com") {
            displayHTMLInModalDialogDefault('Export in Progress', '<div class="tcontent"><div class="tcontent__container"><p class="tcontent__container__text">David je:</p><ul class="tcontent__container__list"><li class="tcontent__container__list__item">krásný!</li><li class="tcontent__container__list__item">mužný!</li><li class="tcontent__container__list__item">vtipný!</li></ul></div></div>', 500, 800);
        }
        else {
            displayHTMLInModalDialogDefault('Export in Progress', '<div class="tcontent"><div class="tcontent__container"><p class="tcontent__container__text">HPE values:</p><ul class="tcontent__container__list"><li class="tcontent__container__list__item">PARTNER!</li><li class="tcontent__container__list__item">INNOVATE!</li><li class="tcontent__container__list__item">ACT!</li></ul></div></div>', 500, 800);
        }
        // show cool image
        $('.ui-draggable').css('z-index', 10000); //bring it forward


        // initiate variable for workbook
        workbook = new ExcelJS.Workbook();
        // creater is the email of OCA user
        //workbook.creator = $('#theData').data().homePageData.pageData.user.email; //this would contain email of user

        // as above just for last modified by
        //workbook.lastModifiedBy = $('#theData').data().homePageData.pageData.user.email; //this would contain email of user
        // add dates to workbook
        if ($.inArray($("#theData").data().homePageData.pageData.user.email,notsopowerusers)>0) {
            workbook.lastModifiedBy=$("#theData").data().homePageData.pageData.user.email
            workbook.creator = $("#theData").data().homePageData.pageData.user.email
        } else {
            workbook.lastModifiedBy = "WW TecHub"
            workbook.creator = "WW TecHub"
        }
        workbook.created = new Date();
        workbook.modified = new Date();
        workbook.lastPrinted = new Date();
        UCID = '' // ucid variable
        oppty = '' // SFDC opportuntity ID variable
        customer = '' // customer name
        getServerData({ //call to initiate Bom tab in OCA
            'method': 'get_tab_content',
            'tabId': 'summary',
            'widget_id': 'extended_overview',
            'selectedKey': 'root',
            'menu_prefs': localStorage.getItem('menu_prefs')
        }, function (data) {
            incoterms = getServerData({ //get pricing info
            'method': 'getCountryGroupData',
            "widget_id": "extended_overview:summary:properties"
            }).data.priceParam.values[0].id
            quoteDetails=data.data.widgets.properties.props;
            getServerData({ // get info bom info from server
                "method": "get_tab_content",
                "tabId": "bom",
                "widget_id": "extended_overview",
                "selectedKey": "root",
                "menu_prefs": "compact_mode=false|show_extended_products=true|show_filters=false|hide_unavailable_items=true|filter_by_date=false|include_mylib=true|include_uds=false|show_extra_columns=false|show_dates=false|show_gauges=true|show_section_messages=true|show_choice_messages=true|show_item_messages=true"
            }, function () {
                getServerData({ // after that get bom with pricing, which is needed.
                    "method": "get_bom",
                    "showPrice": true,
                    "productType": true,
                    "productLine": true,
                    "configName": false,
                    "instanceName": false,
                    "solutionId": false,
                    "showLeadTime": false,
                    "supportFor": true,
                    "showPLC": false,
                    "showLineId": true,
                    "selectAll": true,
                    "widget_id": "extended_overview:bom",
                    "addPrice": true
                }, function (result) {
                    bombig = result.subconfigs[0]; //save server response to bombig variable
                    for (var i = 0; i < quoteDetails.length; i++) { // loop through the quote details to get ucid, oppid and customer name
                        if (quoteDetails[i].id === "ucid") {
                            UCID = quoteDetails[i].value
                        } else if (quoteDetails[i].id === "opportunity_id") {
                            oppty = quoteDetails[i].value
                        } else if (quoteDetails[i].id === "customer_ci_name") {
                            customer = quoteDetails[i].value
                        }
                    }
                    cliccheck() //continue with clock check
                })
            })


        }
                     )




    }

    //procedure that initiates click tab in excel
    cliccheck = function () {
        //preparing the bom
        bom = bombig.bom.subnodes //this contains the subconfigs
        currency = bombig.currency // this stores currency code i.e. EUR, USD...
        bomexport = []
        populateRows(bom, 1, 1) // this thing puts row and columns of future table in variable bomexport
        PLs = [] // product lines appearing in quote array
        for (var i = 0; i < bomexport.length; i++) {
            PLs.push(bomexport[i].pl)
        }// collect all PLs into PL variable
        PLs = PLs.sort().filter((x, i, a) => !i || x != a[i - 1]
                               )// sort and unique
        rows_to_add = Math.max(0, PLs.length - 8)
        // this tells how many rows will beed to be added so that PL table fits. If PLs is smaller than 8, we don't add any rows
        headerRow = 21 + rows_to_add
        // this contains row of header of main table
        bomLength = bomexport.length
        // number of lines in BoM
        lastRow = headerRow + bomLength
        // last line of BoM
        lastPLrow = 11 + PLs.length
        // last line of PLs

        worksheet = workbook.addWorksheet("Budgetary_Quote"); // initiate worksheet for  quote
        clicksheet = workbook.addWorksheet('CLIC') // initiate worksheet for click
        clicksheet.columns = [ // initiate table for CLIC
            {header: 'Item/subitem', key: 'CLitem', width: 10},
            {header: 'Severity', key: 'CLseverity', width: 10},
            {header: 'Rule Number', key: 'CLdesc1', width: 10},
            {header: 'Text', key: 'CLtext', width: 80},
        ];
        clicdata = getServerData({ // call server for click check and store it in clicdata
            "method": "doClicCheck",
            "widget_id": "extended_overview:bom",
            "fanNumber": "N/A"
        }).advicetext
        sheetName = "'Budgetary_Quote'"
        isUplus = false // variable of checking if quote is unbuildable
        unbuildableE=0
        errorE=0
        warningE=0

        for (i = 0; i < clicdata.length; i++) { // iterating through click report
            clicdata[i].assetItemNo=eval(clicdata[i].assetItemNo)
            if (clicdata[i].assetItemNo > bomexport.length) { //workaround for bug 6/8/2018
                clicdata[i].assetItemNo=1
            }
            clicdata[i].adviceText=clicdata[i].adviceText.replace(/<br\/>/g, '\n');
            clicksheet.getCell(i + 2, 1).value = {
                text: clicdata[i].heartItemNo,
                hyperlink: "#" + sheetName + "!A" + (clicdata[i].assetItemNo + headerRow)
            } // populating first cell with heart item and hyperlink
            clicksheet.getCell(i + 2, 1).alignment = {vertical: 'top'}
            // alignment to top
            clicksheet.getCell(i + 2, 1).font = {color: {argb: 'FF0000FF'}, underline: true}
            if (clicdata[i].ruleSeverity == "U+") {
                if (clicdata[i].realSeverity == "F+") {
                    var f = "Override";
                } else {
                    var f = "Unbuildable";
                    unbuildableE++
                }
            } else {
                if (clicdata[i].ruleSeverity == "W+" || clicdata[i].ruleSeverity == "W*" || clicdata[i].ruleSeverity == "F+") {
                    var f = "Warning";
                    warningE++
                } else {
                    var f = "Error";
                    errorE++
                }
            }
            clicdata[i].ruleSeverityText=f
            clicksheet.getCell(i + 2, 2).value = clicdata[i].ruleSeverityText;
            if (clicdata[i].ruleSeverity == "U+") {
                isUplus = true
            }
            // adding rule severity and if U+, unbuildable flag is set to true
            clicksheet.getCell(i + 2, 2).alignment = {vertical: 'top'}
            clicksheet.getCell(i + 2, 3).value = clicdata[i].ruleId;
            // rule ID in third column
            clicksheet.getCell(i + 2, 3).alignment = {vertical: 'top'}
            clicksheet.getCell(i + 2, 4).value = clicdata[i].adviceText
            // text of error in fourth column
            clicksheet.getCell(i + 2, 4).alignment = {vertical: 'top', wrapText: true}
            if (!bomexport[clicdata[i].assetItemNo-1].errors) {
                bomexport[clicdata[i].assetItemNo-1].errors = []
            }
            bomexport[clicdata[i].assetItemNo-1].errors.push(clicdata[i])
        }
        clicdata= _.sortBy(clicdata,'assetItemNo')
        continueQuote()
        //call continue after click check is done
    }

    continueQuote = function () {


        //defining styles
        greenFill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00B388'}, bgColor: {argb: 'FF00B388'}};
        grayFill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFD9D9D9'}, bgColor: {argb: 'FFD9D9D9'}};
        whiteFill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFFFFF'}, bgColor: {argb: 'FFFFFFFF'}};
        thinBorder = {
            top: {style: 'thin'},
            left: {style: 'thin'},
            bottom: {style: 'thin'},
            right: {style: 'thin'}
        }
        thinBorderVertical = {
            left: {style: 'thin'},
            right: {style: 'thin'},
        }
        thinBorderVerticalBottom = {
            left: {style: 'thin'},
            right: {style: 'thin'},
            bottom: {style: 'thin'}
        }
        thinBorderHorizontal = {
            top: {style: 'thin'},
            bottom: {style: 'thin'},
        }
        thinBorderTotal = {
            top: {style: 'double'},
            left: {style: 'thin'},
            bottom: {style: 'thin'},
            right: {style: 'thin'}
        }
        whiteFont = {
            name: 'Arial',
            family: 4,
            size: 12,
            underline: false,
            bold: true,
            color: {argb: 'FFFFFFFF'}
        }
        blackFont = {
            name: 'Arial',
            family: 4,
            size: 12,
            underline: false,
            bold: true,
            color: {argb: 'FF000000'}
        }
        redFont = {
            name: 'Arial',
            family: 4,
            size: 12,
            underline: false,
            bold: true,
            color: {argb: 'FFFF0000'}
        }
        blackFontRegular = {
            name: 'Arial',
            family: 4,
            size: 12,
            underline: false,
            bold: false,
            color: {argb: 'FF000000'}
        }
        smallTitleFont = {
            name: 'Arial',
            family: 4,
            size: 14,
            underline: false,
            bold: true,
            color: {argb: 'FF000000'}
        }
        bigTitleFont = {
            name: 'Arial',
            family: 4,
            size: 24,
            underline: false,
            bold: true,
            color: {argb: 'FF000000'}
        }
        //defining columns and widths

        worksheet.columns = [
            {header: 'Item', key: 'item', width: 7},
            {header: 'Product number', key: 'pn', width: 24},
            {header: 'Description', key: 'desc1', width: 40},
            {header: '', key: 'desc2', width: 40}, //description is split into two so we can have nice columns
            {header: 'Quantity', key: 'qty', width: 10},
            {header: 'Unit List price', key: 'lp', width: 16},
            {header: 'Total List price', key: 'lp_total', width: 16},
            {header: 'Total Disc %', key: 'disc_total', width: 11},
            {header: 'Unit List Price', key: 'np', width: 16},
            {header: 'Total Net Price', key: 'np_total', width: 16},
            {header: 'Class', key: 'cl', width: 8},
            {header: 'PL', key: 'pl', width: 8},
            {header: 'Support for', key: 'supp_for', width: 10},
            {header: 'Target Net Unit Price', key: 'target_np', width: 20},
            {header: 'Target Disc %', key: 'target_disc', width: 20}
        ];
        titleRow = ["HPE Budgetary quote"] //Text in row 1
        worksheet.spliceRows(1, 0, titleRow) // inserting row 1
        //worksheet.addRow(titleRow)
        worksheet.mergeCells('A1:O1'); // all cells merged
        worksheet.getRow(1).getCell(1).font = bigTitleFont; //setting font
        worksheet.getRow(1).getCell(1).alignment = {horizontal: 'center', vertical: 'middle'} //setting alignment
        worksheet.getRow(1).height = 70; //setting row height

        warningRow = ["Internal use only. Only a document \"Legal Quote\" that contains the HPE's terms and conditions engages HPE."]//second row
        worksheet.spliceRows(2, 0, warningRow) // adding second row
        //worksheet.addRow(titleRow)
        worksheet.mergeCells('A2:O2'); // merging cells
        worksheet.getRow(2).getCell(1).font = { //setting font
            name: 'Arial',
            family: 4,
            size: 12,
            bold: true,
            color: {argb: "FFFF0000"}
        };
        worksheet.getRow(2).getCell(1).alignment = {horizontal: 'center'} //centering row 2
        worksheet.getRow(2).height = 16; //second row height
        worksheet.getRow(3).height = 32; // third row height
        if ($.inArray($("#theData").data().homePageData.pageData.user.email,notsopowerusers)>-1) {
            helpRow = [""]
        } else {
            helpRow = ["Should you need assistance with this quote, contact TecHub via Salesforce.com"]
        }// text in row 3
        worksheet.spliceRows(3, 0, helpRow); //inserting the row
        worksheet.getRow(3).alignment = {horizontal: 'center', vertical: 'top'}; //styles
        worksheet.mergeCells('A3:O3'); //merging
        worksheet.getRow(3).getCell(1).value = "Should you need assistance with this quote,contact TecHub via Salesforce.com."; //text in row 3
        worksheet.getRow(3).getCell(1).font = blackFontRegular //styles
        fillerRow = [""] // blank row

        dateRow = ["Date", , new Date(), "Customer"] //date row

        UCIDRow = ["UCID", , UCID, customer] // ucid row
        oppIDRow = ["Opportunity ID", , oppty] // oppty row
        currencyRow = ["Currency", , currency] // currency row
        incotermsRow = ["Incoterms", , incoterms] // incoterm row
        worksheet.spliceRows(4, 0, dateRow, UCIDRow, oppIDRow, currencyRow, incotermsRow, fillerRow) // add rows
        // logo in base64 follows
        var myBase64Image = "data:image/gif;base64,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"
        var imageId2 = workbook.addImage({
            base64: myBase64Image,
            extension: 'gif',
        }); //pasting image into workbook logic
        worksheet.addImage(imageId2, {
            tl: {col: 1, row: 0},
            br: {col: 2.9999, row: 1},
            //tl: { col: 1, row: 0 },
            //br: { col: 2.9, row: 2 },
            editAs: 'oneCell'
        }); //pasting image into worksheet

        for (var i = 4; i < 9; i++) { // adding styles to header
            worksheet.getRow(i).font = blackFont
            worksheet.getRow(i).getCell(1).font = whiteFont
            worksheet.getRow(i).getCell(1).fill = greenFill
            worksheet.getRow(i).getCell(1).border = thinBorder
            worksheet.getRow(i).getCell(1).alignment = {horizontal: 'left'}
            worksheet.getRow(i).getCell(3).alignment = {horizontal: 'center'}
            worksheet.getRow(i).getCell(3).border = thinBorder
            worksheet.mergeCells('A' + i + ':B' + i);
        }
        worksheet.mergeCells('D5:O8'); //merging
        worksheet.mergeCells('D4:O4');

        worksheet.getRow(4).getCell(4).border = thinBorder //headers styles
        worksheet.getRow(4).getCell(4).fill = greenFill
        worksheet.getRow(4).getCell(4).font = whiteFont
        worksheet.getRow(4).getCell(4).alignment = {horizontal: 'center'}
        worksheet.getRow(5).getCell(4).alignment = {horizontal: 'center', vertical: 'middle', wrapText: true}
        worksheet.getRow(5).getCell(4).font = {
            name: 'Arial',
            family: 4,
            size: 24,
            underline: false,
            bold: true,
            color: {argb: 'FF000000'}
        }
        worksheet.getRow(5).getCell(4).border = thinBorder // border for eader

        subtotalTitleRow = ['Subtotals by Product Class', , , , , , , , 'Subtotals by Product Line'] //titles for subtotals and PLs tables
        worksheet.spliceRows(10, 0, subtotalTitleRow) //add title row
        worksheet.getRow(10).font = smallTitleFont //formating
        for (c = 1; c < 16; c++) {
            worksheet.getRow(10).getCell(c).fill = whiteFill; //whitening header
        }
        headersSubtotalsRow = ["Class", , "List Price", "Net Price", "Disc. %", "Target Price", "Target Disc %", , "PL", "List Price", "Net Price", , "Disc %", "Target Price", "Target Disc %"]
        worksheet.spliceRows(11, 0, headersSubtotalsRow) //headers of subtotals and PLs
        for (var c = 1; c < 16; c++) { //and their formating
            if (c != 8) {
                worksheet.getRow(11).getCell(c).fill = greenFill
                worksheet.getRow(11).getCell(c).font = whiteFont
                worksheet.getRow(11).getCell(c).border = thinBorderVertical
                worksheet.getRow(11).getCell(c).alignment = {horizontal: 'center'}
            }
        }
        worksheet.mergeCells('A11:B11'); //merging of description
        worksheet.mergeCells('K11:L11'); //merging of  PL and PC in PLs


        HWRow = ['Hardware'] //subtotals rows
        SWRow = ['Software']
        SupportRow = ['Support']
        ServicesRow = ['Services']
        TotalsRow = ['Totals']

        worksheet.spliceRows(12, 0, HWRow, SWRow, SupportRow, ServicesRow, TotalsRow) //psting rows

        worksheet.mergeCells('A12:B12'); //necessary merhing
        worksheet.mergeCells('A13:B13');
        worksheet.mergeCells('A14:B14');
        worksheet.mergeCells('A15:B15');
        worksheet.mergeCells('A16:B16');

        //formating
        for (var r = 12; r < 17; r++) {
            worksheet.getRow(r).font = blackFontRegular //font
            worksheet.getRow(r).getCell(1).border = thinBorder //border
            worksheet.getRow(r).getCell(3).border = thinBorder
            worksheet.getRow(r).getCell(4).border = thinBorder
            worksheet.getRow(r).getCell(5).border = thinBorder
            worksheet.getRow(r).getCell(6).border = thinBorder
            worksheet.getRow(r).getCell(7).border = thinBorder
            worksheet.getRow(r).getCell(3).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@'; //numformat
            worksheet.getRow(r).getCell(4).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getRow(r).getCell(5).numFmt = '0.00%';
            worksheet.getRow(r).getCell(5).value = { //formula
                formula: 'IF(C' + r + '=0, 0, 1-D' + r + '/C' + r + ')',
                result: undefined
            }
            worksheet.getRow(r).getCell(6).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getRow(r).getCell(7).numFmt = '0.00%';
            worksheet.getRow(r).getCell(6).fill = grayFill // fills
            worksheet.getRow(r).getCell(7).fill = grayFill
            worksheet.getRow(r).getCell(8).fill = whiteFill
            worksheet.getRow(r).getCell(1).font = blackFont // fonts
            worksheet.getRow(r).getCell(4).font = blackFont
            worksheet.getRow(r).getCell(5).font = blackFont
            worksheet.getRow(r).getCell(6).font = redFont
            worksheet.getRow(r).getCell(7).font = redFont
            worksheet.getRow(r).getCell(7).numFmt = '0.00%'
        }
        worksheet.getRow(16).getCell(1).border = thinBorderTotal //borders
        worksheet.getRow(16).getCell(3).border = thinBorderTotal
        worksheet.getRow(16).getCell(4).border = thinBorderTotal
        worksheet.getRow(16).getCell(5).border = thinBorderTotal
        worksheet.getRow(16).getCell(6).border = thinBorderTotal
        worksheet.getRow(16).getCell(7).border = thinBorderTotal
        if (unbuildableE+errorE+warningE == 0) {
            ErrorRowText = ["The Quote contains no errors"]
        } else {
            ErrorRowText = ["The Quote contains: " + unbuildableE + " Unbuildable Error"+(unbuildableE>1?'s':'')+", " + errorE + " Error"+(errorE>1?'s':'')+" and " + warningE  + " Warning"+(warningE>1?'s':'')+". See the other tab for CLIC report."]
        }
        //ErrorRow = worksheet.getRow(headerRow-1);
        //worksheet.spliceRows(headerRow, 0, ErrorRowText) // adding second row
        //headerRow++


        TotalRow = ["", "TOTAL NET excl. VAT :", , 6666] //total row
        worksheet.spliceRows(17, 0, fillerRow, TotalRow, fillerRow, fillerRow)
        errorsummary={}
        errorsummary.richText=[]
        if (unbuildableE+errorE+warningE == 0) {
            errorsummary.richText.push(
                {'font': {'size': 12,'color': {'argb': 'FF000000'},'name': 'Arial','family': 2,'scheme': 'minor'},'text': 'This quote contains no errors'})
        } else {
            errorsummary.richText.push(
                {'font': {'size': 20,'color': {'argb': 'FF000000'},'name': 'Arial','family': 2,'scheme': 'minor','bold':'true'},'text': '!!!!! This quote contains: '})
            if (unbuildableE>0) {
                errorsummary.richText.push(
                    {'font': {'size': 20,'color': {'argb': 'FFFF0000'},'name': 'Arial','family': 2,'scheme': 'minor','bold':'true'},'text': unbuildableE + ' Unbuildable errors, '})
            }
            if (errorE>0) {
                errorsummary.richText.push(
                    {'font': {'size': 20,'color': {'argb': 'FF7F007F'},'name': 'Arial','family': 2,'scheme': 'minor','bold':'true'},'text': errorE + ' Errors, '})
            }
            if (warningE>0) {
                errorsummary.richText.push(
                    {'font': {'size': 20,'color': {'argb': 'FF2BA6CB'},'name': 'Arial','family': 2,'scheme': 'minor','bold':'true'},'text': warningE + ' Warnings, '})
            }
            errorsummary.richText.push(
                {'font': {'size': 20,'color': {'argb': 'FF000000'},'name': 'Arial','family': 2,'scheme': 'minor','bold':'true'},'text': 'which you can review CLIC report in the other tab. !!!!!'})
        }
        worksheet.getCell('A20').value = errorsummary;
        worksheet.mergeCells('B18:C19'); //merging
        worksheet.mergeCells('D18:F19');
        worksheet.getRow(18).getCell(2).border = thinBorderHorizontal //formating
        worksheet.getRow(18).getCell(4).border = thinBorderHorizontal
        worksheet.getRow(18).getCell(4).alignment = {horizontal: 'right'}
        worksheet.getRow(18).getCell(4).numFmt = '###\ ###\ ###.00" ' + currency + '";-###\ ###\ ###.00" ' + currency + '";;@'
        worksheet.getRow(18).getCell(2).font = bigTitleFont;
        worksheet.getRow(18).getCell(4).font = bigTitleFont;


        //this is the hidden area on the right from former excel. It helps to calculate discounts correctly
        worksheet.getCell('AA10').value = 'Hidden Formulas'
        worksheet.getCell('AA11').value = 'BS'
        worksheet.getCell('AA12').value = 'ES'
        worksheet.getCell('AA13').value = 'HS'
        worksheet.getCell('AA14').value = 'HW'
        worksheet.getCell('AA15').value = 'IN'
        worksheet.getCell('AA16').value = 'SS'
        worksheet.getCell('AA17').value = 'SV'
        worksheet.getCell('AA18').value = 'SW'
        worksheet.getCell('AB10').value = { // there is lot of formulas
            formula: 'IF(OR(C16=0, F16=0),IF(G16<>0,G16,0),(1-F16/C16))',
            result: undefined
        }
        worksheet.getCell('AB11').value = {
            formula: 'IF(AB10=0,IF(OR(C15=0, F15=0),IF(G15<>0,G15,0),(1-F15/C15)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB12').value = {
            formula: 'IF(AB10=0,IF(OR(C14=0, F14=0),IF(G14<>0,G14,0),(1-F14/C14)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB13').value = {
            formula: 'IF(AB10=0,IF(OR(C15=0, F15=0),IF(G15<>0,G15,0),(1-F15/C15)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB14').value = {
            formula: 'IF(AB10=0,IF(OR(C12=0, F12=0),IF(G12<>0,G12,0),(1-F12/C12)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB15').value = {
            formula: 'IF(AB10=0,IF(OR(C15=0, F15=0),IF(G15<>0,G15,0),(1-F15/C15)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB16').value = {
            formula: 'IF(AB10=0,IF(OR(C14=0, F14=0),IF(G14<>0,G14,0),(1-F14/C14)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB17').value = {
            formula: 'IF(AB10=0,IF(OR(C15=0, F15=0),IF(G15<>0,G15,0),(1-F15/C15)),AB10)',
            result: undefined
        }
        worksheet.getCell('AB18').value = {
            formula: 'IF(AB10=0,IF(OR(C13=0, F13=0),IF(G13<>0,G13,0),(1-F13/C13)),AB10)',
            result: undefined
        }
        worksheet.getCell('AC11').value = 'SERVICE' //texts
        worksheet.getCell('AC12').value = 'SUPPORT'
        worksheet.getCell('AC13').value = 'SERVICE'
        worksheet.getCell('AC14').value = 'HARDWARE'
        worksheet.getCell('AC15').value = 'SERVICE'
        worksheet.getCell('AC16').value = 'SUPPORT'
        worksheet.getCell('AC17').value = 'SERVICE'
        worksheet.getCell('AC18').value = 'SOFTWARE'

        for (i = 1; i <= rows_to_add; i++) {
            worksheet.spliceRows(20, 0, fillerRow) //adding filler rows if there is too much PLS
        }

        for (i = 0; i < PLs.length; i++) { // formating PLs summary and adding formlas
            worksheet.getRow(12 + i).getCell(9).value = PLs[i]
            worksheet.getRow(12 + i).getCell(9).alignment = {horizontal: 'center'}
            worksheet.getRow(12 + i).getCell(9).font = blackFont
            worksheet.getRow(12 + i).getCell(9).border = thinBorder
            worksheet.getRow(12 + i).getCell(10).value = {
                formula: 'SUMIF($L$' + (headerRow + 1) + ':$L$' + lastRow + ',I' + (12 + i) + ',$G$' + (headerRow + 1) + ':$G$' +
                lastRow + ')', result: undefined
            }
            worksheet.getRow(12 + i).getCell(10).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@'
            worksheet.mergeCells('K' + (12 + i) + ':L' + (12 + i));
            worksheet.getRow(12 + i).getCell(10).border = thinBorder
            worksheet.getRow(12 + i).getCell(10).font = blackFontRegular
            worksheet.getRow(12 + i).getCell(11).value = {
                formula: 'SUMIF($L$' + (headerRow + 1) + ':$L$' + lastRow + ',I' + (12 + i) + ',$J$' + (headerRow + 1) + ':$J$' +
                lastRow + ')', result: undefined
            }
            worksheet.getRow(12 + i).getCell(11).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@'
            worksheet.getRow(12 + i).getCell(11).border = thinBorder
            worksheet.getRow(12 + i).getCell(11).font = blackFont
            worksheet.getRow(12 + i).getCell(13).value = {
                formula: 'IF(J' + (12 + i) + '=0,0,1-K' + (12 + i) + '/J' + (12 + i) + ')',
                result: undefined
            }
            worksheet.getRow(12 + i).getCell(13).numFmt = '0.00%'
            worksheet.getRow(12 + i).getCell(13).border = thinBorder
            worksheet.getRow(12 + i).getCell(13).font = blackFont
            worksheet.getRow(12 + i).getCell(14).border = thinBorder
            worksheet.getRow(12 + i).getCell(14).font = redFont
            worksheet.getRow(12 + i).getCell(14).fill = grayFill
            worksheet.getRow(12 + i).getCell(14).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@'
            worksheet.getRow(12 + i).getCell(15).numFmt = '0.00%'
            worksheet.getRow(12 + i).getCell(15).border = thinBorder
            worksheet.getRow(12 + i).getCell(15).font = redFont
            worksheet.getRow(12 + i).getCell(15).fill = grayFill
            worksheet.getRow(12 + i).getCell(26).value = {
                formula: 'IF(OR(J' + (12 + i) + '=0, N' + (12 + i) + '=0),IF(O' + (12 + i) + '<>0,O' + (12 + i) + ',0),(1-N' + (12 + i) + '/J' + (12 + i) + '))',
                result: undefined
            }
        }
        //pls formatting and formalas ends

        for (i = 1; i < 16; i++) {
            cell = worksheet.getRow(headerRow).getCell(i);
            cell.font = whiteFont
            cell.fill = greenFill
            cell.border = thinBorder
        } //formatting header row of main table per column
        firstRow = worksheet.getRow(headerRow);
        firstRow.alignment = {vertical: 'top', horizontal: 'center', wrapText: true};
        firstRow.height = 36; //formating header row

        for (var r = 0; r < bomexport.length; r++) { //creating values for main table
            var row = []
            row[1] = bomexport[r].item
            row[2] = bomexport[r].pn
            row[3] = bomexport[r].descr
            row[5] = bomexport[r].qty
            row[6] = bomexport[r].lp
            row[11] = bomexport[r].cl
            row[12] = bomexport[r].pl
            row[13] = bomexport[r].supp_for
            aa = worksheet.addRow(row);
        }
        worksheet.getCell('G' + (headerRow + 1)).value = {
            formula: 'E' + (headerRow + 1) + '*F' + (headerRow + 1),
            result: undefined
        }; //Total List price formula
        worksheet.getCell('B' + (r + headerRow + 1)).alignment = {indent: bomexport[0].depth}
        //setting indents so the nesting appears

        for (var r = 1; r < bomexport.length; r++) { //formulas for main table are shared
            worksheet.getCell('G' + (r + headerRow + 1)).value = {
                sharedFormula: 'G' + (headerRow + 1),
                result: undefined
            }; //Total List price
            worksheet.getCell('B' + (r + headerRow + 1)).alignment = {indent: bomexport[r].depth}
            //outlining gives the + buttons to hide and show stuff. I think it's ugly, but it might come handy some time
            //worksheet.getRow(r + headerRow + 1).outlineLevel=bomexport[r].depth
            /*worksheet.properties.outlineProperties = {
             summaryBelow: false,
             summaryRight: false,
             };
             */
        }


        for (var r = (headerRow + 1); r < (lastRow + 1); r++) { //formating and main formulas

            worksheet.getCell('F' + r).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getCell('G' + r).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getCell('I' + r).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getCell('J' + r).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getCell('N' + r).numFmt = '###\ ###\ ###.00"";-###\ ###\ ###.00;;@';
            worksheet.getCell('H' + r).numFmt = '0.00%;-0.00%;;@';
            worksheet.getCell('O' + r).numFmt = '0.00%';
            worksheet.getCell('E' + r).alignment = {horizontal: 'center'};
            worksheet.getCell('K' + r).alignment = {horizontal: 'center'};
            worksheet.getCell('L' + r).alignment = {horizontal: 'center'};
            worksheet.getCell('N' + r).fill = grayFill;
            worksheet.getCell('O' + r).fill = grayFill;
            worksheet.mergeCells('C' + r + ':D' + r); //should be removed to allow sorting
            for (var c = 1; c < 16; c++) {
                worksheet.getRow(r).getCell(c).border = thinBorderVertical
                worksheet.getRow(r).getCell(c).font = blackFontRegular
            }
            worksheet.getCell('N' + r).font = redFont;
            worksheet.getCell('O' + r).font = redFont;
            worksheet.getCell('G' + r).font = blackFont;
            worksheet.getCell('H' + r).font = blackFont;
            worksheet.getCell('H' + r).value = { // it looks nasty but I used our excel and added just r and last PLrow when needed etc.
                formula: 'IF(E' + r + '*F' + r + '=0,0,IF(AA' + r + '<>0,AA' + r + ',IF(LOOKUP(L' + r + ',$I$12:$I$' + lastPLrow + ',$Z$12:$Z$' + lastPLrow + ')<>0,LOOKUP(L' + r + ',$I$12:$I$' + lastPLrow + ',$Z$12:$Z$' + lastPLrow + '),IF(E' + r + '*F' + r + '=0,0,IF(LOOKUP(K' + r + ',$AA$11:$AA$18,$AB$11:$AB$18)<>0,LOOKUP(K' + r + ',$AA$11:$AA$18,$AB$11:$AB$18),Z' + r + '+AC' + r + ')))))',
                result: undefined
            }
            worksheet.getCell('I' + r).value = {
                formula: 'IF(E' + r + '*F' + r + '=0,0,F' + r + '*(1-H' + r + '))',
                result: undefined
            };
            worksheet.getCell('J' + r).value = {formula: 'E' + r + '*I' + r, result: undefined};
            worksheet.getCell('J' + r).value = {formula: 'E' + r + '*I' + r, result: undefined};
            worksheet.getCell('AA' + r).value = {
                formula: 'IF(OR(F' + r + '=0,N' + r + '=0),IF(O' + r + '<>0,O' + r + ',0),(1-N' + r + '/F' + r + '))',
                result: undefined
            };
        }

        worksheet.getCell('D18').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"HW",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SW",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"ES",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"IN",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"HS",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SV",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"BS",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SS",J' + (headerRow + 1) + ':J' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('C12').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"HW",G' + (headerRow + 1) + ':G' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('C13').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SW",G' + (headerRow + 1) + ':G' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('C14').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"ES",G' + (headerRow + 1) + ':G' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SS",G' + (headerRow + 1) + ':G' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('C15').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"IN",G' + (headerRow + 1) + ':G' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SV",G' + (headerRow + 1) + ':G' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"HS",G' + (headerRow + 1) + ':G' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"BS",G' + (headerRow + 1) + ':G' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('C16').value = {formula: 'SUM(C12:C15)', result: undefined}

        worksheet.getCell('D12').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"HW",J' + (headerRow + 1) + ':J' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('D13').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SW",J' + (headerRow + 1) + ':J' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('D14').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"ES",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SS",J' + (headerRow + 1) + ':J' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('D15').value = {
            formula: 'SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"IN",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"SV",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"HS",J' + (headerRow + 1) + ':J' + lastRow + ') + SUMIF(K' + (headerRow + 1) + ':K' + lastRow + ',"BS",J' + (headerRow + 1) + ':J' + lastRow + ')',
            result: undefined
        }
        worksheet.getCell('D16').value = {formula: 'SUM(D12:D15)', result: undefined}

        worksheet.autoFilter = { // autofilter on header row, this helps filtering i.e. 0D1
            from: 'A' + headerRow,
            to: 'O' + headerRow
        };
        worksheet.mergeCells('C' + headerRow + ':D' + headerRow); // merge description


        for (i = 9; i < 16; i++) { // border on last PL row
            worksheet.getRow(lastPLrow + 1).getCell(i).border = {
                top: {style: 'thin'},
            }
        }


        for (i = 1; i < 16; i++) { // border on last row
            worksheet.getRow(lastRow).getCell(i).border = thinBorderVerticalBottom
        }

        worksheet.pageSetup.paperSize = 9 // paper = A4
        worksheet.views = [{zoomScale: 70, showGridLines: false}]; // cells that have no borders defined don't display border and setting zoom to 70% so that whole document in width is nicely seen
        worksheet.pageSetup.margins = { // margins in inches
            left: 0.7, right: 0.7,
            top: 0.75, bottom: 0.75,
            header: 0.3, footer: 0.3
        };
        for (i = 16; i < 36; i++) { // hidding columns that should be hidden
            worksheet.getColumn(i).hidden = true
        }

        if (unbuildableE+errorE > 0) {
            worksheet.getRow(headerRow-1).getCell(1).font = redFont
        }

        worksheet.pageSetup.fitToPage = true; // fit to page for print
        worksheet.pageSetup.fitToHeight = 9999; // maximum pages in height
        worksheet.pageSetup.fitToWidth = 1; // always 1 peage in width
        worksheet.pageSetup.printArea = 'A1:O' + (lastRow); // set print area

        if (isUplus) { // if unbildable, pop up this message
            displayHTMLInModalDialog('Quote contains Unbuildable Error', 'Your quote was exported however it contains unbuildable error(s). You should either correct the errors or let the requestor know about them.', 500, 500)
        }
        var quoteinfo=getServerData({method: "get_config_info",widget_id:'toolbar'})
        var filename=(quoteinfo.configName + " " + quoteinfo.ucid).replace(/\.(?=.*?\.)/g, "_").replace(/ /g,"_");
        var a = getServerData({method: 'export_configuration', 'config_name': '', 'exportType': 'ECLIPSE_XML'}) //downloads xml file
        a.filename = filename + ".xml"
        download(atob(a.file_content), a.filename) //and this downloads it
        if (sdd)
        {
            var b = getServerData({ // this exports sdd file which we decided not to continue with
                "method": "export_to_SBWorWastonConfiguration",
                "config_name": "",
                "single_bunde": false,
                "exportType": "WQSDD"
            })
            b.filename = filename + ".sdd"
            download(atob(b.file_content), b.filename)
        }
        var c = a.filename.replace('xml', 'xlsx') // xml will have the same name as xlsx
        var buff = workbook.xlsx.writeBuffer().then(function (data) { //load the buffer with file
            var blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
            download(blob, c); // download the xlsx
        });
        step3 = function () { //step 3 confirms download button
            if (!maui.blockUI.isWaiting()) { // waiting for unblock
                $('#download_button').click(); // click on download button
                setTimeout(function (){$('#ok_button').click()},4000); // and on ok button
                clearInterval(myVar1) // stops interval calling step 3
            }
            ;
        };
        step2 = function () { //step 3 clicks on save to local button
            if (!maui.blockUI.isWaiting()) { // waiting for unblock
                $('#save2local_btn').click(); // clicking on save button
                clearInterval(myVar); //stops interval checking step 2
                myVar1 = setInterval(function () { //starts interval checking step 3
                    step3()
                }, 100);
            }
        };
        myVar = setInterval(function () { //launches step 2
            step2()
        }, 100)
    }
}


populateRows = function (arr, q, d) { //this helps populate rows doing recursive dig. arr is array, q us quantity of parent product (not used) and d is depth in which we are compared inside the quote nesting.
    var d;
    for (var i = 0; i < arr.length; i++) {
        var row = {}
        row.item = arr[i].attributes.sequencial_number
        row.pn = arr[i].attributes.product_number
        row.descr = arr[i].description
        row.qty = arr[i].attributes.quantity
        row.lp = arr[i].attributes.unit_price_value
        row.cl = arr[i].attributes.product_type
        row.pl = arr[i].attributes.productLine
        row.supp_for = arr[i].attributes.support_for
        row.depth = d - 1
        bomexport.push(row)
        populateRows(arr[i].subnodes, arr[i].attributes.quantity, d + 1)
        ;
    }
};
