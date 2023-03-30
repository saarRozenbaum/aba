<%ShowNoProcedure = "yes"%>
<!--#include file="../Inc/Header.asp"-->

<%  
    openDB

    Dim iMaxRowNumber
        iMaxRowNumber = 200
	
    If Len(Trim(Request.Form("selCustomerFrom"))) <> Int(0) Then
        arrTemp = Split(Request.Form("selCustomerFrom"), "_")

        Dim iFromSRCompany
            iFromSRCompany = arrTemp(0)
    End If
    If Len(Trim(iFromSRCompany)) = Int(0) Then
        iFromSRCompany = 0
    End If

	Dim strInnerLineCounter
	    strInnerLineCounter = 10

    Dim iShippingCost
        iShippingCost = 0

    Dim iTheDay
        iTheDay		= Day(date)
    Dim iTheMonth
        iTheMonth	= Month(date)
    Dim iTheYear
        iTheYear	= Year(date)
		
	Dim strAlignAndDir
    	strAlignAndDir = " align='left' dir=ltr "
	Dim strAlignAndDirBottom
	    strAlignAndDirBottom = " align='right' dir=ltr "
	Dim strDirTable
	    strDirTable = " dir=ltr "
	Dim strRowNumTitle
	    strRowNumTitle = "&nbsp;"

	' פונקציה פשוטה שמקדמת משתנה	
	Function IncTabIndex(ByRef iVal)
		iVal=iVal+1
		IncTabIndex=iVal
	End Function
	
	iTabIndex = 0

	Function FormatDateToString(Day, Month, Year)
	    Select Case Month
	    case 1:
		    strMonth = "January"
	    case 2:
		    strMonth = "February"
	    case 3:
		    strMonth = "March"
	    case 4:
		    strMonth = "April"
	    case 5:
		    strMonth = "May"
	    case 6:
		    strMonth = "June"
	    case 7:
		    strMonth = "July"
	    case 8:
		    strMonth = "August"
	    case 9:
		    strMonth = "September"
	    case 10:
		    strMonth = "October"
	    case 11:
		    strMonth = "November"
	    case 12:
		    strMonth = "December"
	    case else
		    strMonth = ""
	    End select

	    FormatDateToString = strMonth & " " & Day & "th, " & Year
	End Function
	
	Function SetInvoiceTitle(strInvoiceType, strInvoiceNo)
        Select Case strInvoiceType
	    case "INVOICE":
		    SetInvoiceTitle = "I&nbsp;&nbsp;N&nbsp;&nbsp;V&nbsp;&nbsp;O&nbsp;&nbsp;I&nbsp;&nbsp;C&nbsp;&nbsp;E"
	    case "MEMO":
		    SetInvoiceTitle = "&nbsp;&nbsp;M&nbsp;&nbsp;E&nbsp;&nbsp;M&nbsp;&nbsp;O"
	    case "PRO-FORMA":
		    SetInvoiceTitle = "&nbsp;&nbsp;P&nbsp;&nbsp;R&nbsp;&nbsp;O&nbsp;&nbsp;-&nbsp;&nbsp;F&nbsp;&nbsp;O&nbsp;&nbsp;R&nbsp;&nbsp;M&nbsp;&nbsp;A"
	    case "RETURN CONSIGNMENT":
		    SetInvoiceTitle = "&nbsp;R&nbsp;E&nbsp;T&nbsp;U&nbsp;R&nbsp;N&nbsp;&nbsp;&nbsp;&nbsp;C&nbsp;O&nbsp;N&nbsp;S&nbsp;I&nbsp;G&nbsp;N&nbsp;M&nbsp;E&nbsp;N&nbsp;T"
	    case "PURCHASE":
	        SetInvoiceTitle = "&nbsp;P&nbsp;U&nbsp;R&nbsp;C&nbsp;H&nbsp;A&nbsp;S&nbsp;E"
	    case "RETURN INV":
	        SetInvoiceTitle = "&nbsp;R&nbsp;E&nbsp;T&nbsp;U&nbsp;R&nbsp;N&nbsp;&nbsp;I&nbsp;N&nbsp;V"
	    case else
		    SetInvoiceTitle = "I&nbsp;&nbsp;N&nbsp;&nbsp;V&nbsp;&nbsp;O&nbsp;&nbsp;I&nbsp;&nbsp;C&nbsp;&nbsp;E"
	    End select
	    
	    If Len(Trim(strInvoiceNo)) > Int(0) Then
	        SetInvoiceTitle = SetInvoiceTitle & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;NO.&nbsp;&nbsp;" & strInvoiceNo
	    End If
	End Function
	

    Function getExportFromDescription(iCompID)
        If iCompID = "0_0" Then
            getExportFromDescription = ""
            Exit Function
        End If
        Dim arr
            arr = split(iCompID, "_")
        
        iCompID = arr(0)


        Dim strSql
		    strSql = "Select * From " & strCustomerDB_Alias & "CustomerSRCompany Where ID = " & iCompID

        doQuery(strSql)
        	
        If rs.RecordCount <> 0 Then
    		
	        Do While Not rs.EOF
		        strButtomCompanyName = "<b>" & rs("CompanyName") & "</b><br>"
    			
		        If Len(Trim(rs("CompanyName"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & "<font size=3><b>" & Replace(rs("CompanyName"), " ", "&nbsp;") & "</b></font><br>"
		        End If
		        If Len(Trim(rs("LineLogo1"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & rs("LineLogo1") & "<br>"
		        End If
		        If Len(Trim(rs("LineLogo2"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & rs("LineLogo2") & "&nbsp;"
		        End If
		        If Len(Trim(rs("LineLogo3"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & rs("LineLogo3") & "&nbsp;"
		        End If
		        If Len(Trim(rs("LineLogo4"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & rs("LineLogo4") & "<br>"
		        End If
		        If Len(Trim(rs("LineLogo5"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & "TEL:&nbsp;" & rs("LineLogo5") & "<br>"
		        End If
		        If Len(Trim(rs("LineLogo6"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & "FAX:&nbsp;" & rs("LineLogo6") & "<br>"
		        End If
		        If Len(Trim(rs("LineLogo7"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & "email:&nbsp;" & rs("LineLogo7") & "<br>"
		        End If
		        If Len(Trim(rs("LineLogo8"))) > Int(0) Then
			        strSRCompanyTitle = strSRCompanyTitle & "website:&nbsp;" & rs("LineLogo8") & "<br>"
		        End If

		        rs.MoveNext
	        Loop
        End If
        getExportFromDescription = strSRCompanyTitle
    End Function
    
    Function FormatDestCustomerDescription(strCusDescription, iShowName, iShowAddres)
        FormatDestCustomerDescription = ""
        
        If Len(Trim(strCusDescription)) = Int(0) Then
            Exit Function
        End If

        Dim arrLine
            arrLine = Split(strCusDescription, "<br>")
        
        If Int(iShowName) = Int(1) Then
            FormatDestCustomerDescription = FormatDestCustomerDescription & arrLine(0) & "<br>"
        End If
        
        If Int(iShowAddres) = Int(1) Then
        
            For IndxID = 1 To UBound(arrLine)
                FormatDestCustomerDescription = FormatDestCustomerDescription & arrLine(IndxID) & "<br>"
			Next
        End If
    End Function
%>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">

    // Function that makes and Calls the Ajax
	function createRequestObject() {
		var reqObj = null;
		if (window.XMLHttpRequest){
			try {
				reqObj = new XMLHttpRequest();
			} catch(e) {
			//some kind of a weird mistake...you choose what you want to do
				reqObj = false;
			}
		} else if (window.ActiveXObject){
			try{
				reqObj = new ActiveXObject("Msxml2.HTMLHTTP");
			} catch(e) {
				try{
					reqObj=new ActiveXObject("Microsoft.XMLHTTP");
				} catch(e) {
					//don't know what to do...you choose
					reqObj = false;
				}
			}
		} 
		return reqObj;
	}	

    /////////////////////////////////////////////////////////////////////////
    function getCustomerAddress(hunSelect, hunText){
        var iSelectValue = hunSelect.options[hunSelect.selectedIndex].value;
        var iSelectText = hunSelect.options[hunSelect.selectedIndex].text;
        
        
        if(parseInt(iSelectValue, 10) != parseInt(0, 10) && parseInt(hunText.value.length, 10) == parseInt(4, 10)){
            var strSRSerialNum = iSelectText + hunText.value;
            var strUrl = "../CRM/getCustomerAddress.asp?IDToken=<%=IDToken%>&SrSerialNumber=" + strSRSerialNum;

            var strAddressDesc = getAjaxCustomerAddressDescription(strUrl);
        }
    }

    function getAjaxCustomerAddressDescription(strUrl){
		var http_request = createRequestObject();
	
		http_request.onreadystatechange = function() { getActionCustomerAddressDescription(http_request); };
        http_request.open("POST", strUrl, true);
        http_request.send(null);
    }
    
    function getActionCustomerAddressDescription(http_request, divName) {
        if (http_request.readyState == 4) {
            if (http_request.status == 200) {
				var responceString = http_request.responseText;

                setCustomerAddressDescription(responceString);
            } else {
                return 'There was a problem with the request.';
            }
        }
    }		
    
    function setCustomerAddressDescription(strText){

        
        var iCusID = strText.toString().substring(0, strText.toString().indexOf("___", 0));
        strText = strText.toString().substring(strText.toString().indexOf("___", 0)+4, strText.length);

        document.theForm.hidCustomerDesc.value = strText;
        document.theForm.hidDestCustomerID.value = iCusID;
        
        strText = strText.toString().replace("<br>", '\n');
        strText = strText.toString().replace("<br>", '\n');
        strText = strText.toString().replace("<br>", '\n');
        strText = strText.toString().replace("<br>", '\n');
        
        //alert(iCusID)
        document.theForm.taCustomerDesc.value = strText;
    }
    
    function replace_CR_BR(srcHun, destHun){
        strText = srcHun.value;
        
        strText = strText.toString().replace('\n', "<br>");
        strText = strText.toString().replace('\n', "<br>");
        strText = strText.toString().replace('\n', "<br>");
        strText = strText.toString().replace('\n', "<br>");
        strText = strText.toString().replace('\n', "<br>");
        strText = strText.toString().replace('\n', "<br>");
        strText = strText.toString().replace('\n', "<br>");
        
        destHun.value = strText;
    }
    /////////////////////////////////////////////////////////////////////////
    function getBankSelectOptions(hudSelect){
        
        var iCusID = hudSelect.options[hudSelect.selectedIndex].value;
        //alert(iCusID);
        var strUrl = "../CRM/getSelectOption.asp?IDToken=<%=IDToken%>&SelType=1&iRecID=" + iCusID;
        
        //window.open(strUrl);
        makeRequestGetBankOptions(strUrl, "divBankDescription", "Post");
    }
    
	function makeRequestGetBankOptions(url, divName, strMethod){
		var http_request = createRequestObject();
	
		http_request.onreadystatechange = function() { goActionGetBankOptions(http_request, divName); };
        http_request.open(strMethod, url, true);
        http_request.send(null);
	}
	
	function goActionGetBankOptions(http_request, divName) {
        if (http_request.readyState == 4) {
            if (http_request.status == 200) {
				var responceString = http_request.responseText;

                setDivActionGetBankOptions(divName, responceString)
            } else {
                return 'There was a problem with the request.';
            }
        }
    }		

    function setDivActionGetBankOptions(div, text){
        var strDivText = "Select shipper bank:<SELECT name=selBankID onchange=getBankDescription(this);><OPTION value=-1></OPTION>" + text + "</SELECT>";
    
        eval("document.all." + div + ".innerHTML='" + strDivText + "';");
    }
    
    
    
    
    function getBankDescription(hudSelect){
        var iBankID = hudSelect.options[hudSelect.selectedIndex].value;

        var strUrl = "../CRM/getBankAddress.asp?IDToken=<%=IDToken%>&iBankID=" + iBankID;
        
        //window.open(strUrl);
        getAjaxBankAddressDescription(strUrl);
    }
    
    
    
    
    function getAjaxBankAddressDescription(strUrl){
		var http_request = createRequestObject();
	
		http_request.onreadystatechange = function() { getActionBankAddressDescription(http_request); };
        http_request.open("POST", strUrl, true);
        http_request.send(null);
    }
    
    function getActionBankAddressDescription(http_request, divName) {
        if (http_request.readyState == 4) {
            if (http_request.status == 200) {
				var responceString = http_request.responseText;

                setBankAddressDescription(responceString);
            } else {
                return 'There was a problem with the request.';
            }
        }
    }		
    
    function setBankAddressDescription(strText){

/*        
        document.theForm.hidCustomerBankTitle.value = strText;
        
        strText = strText.toString().replace("<br>", '\n');
        strText = strText.toString().replace("<br>", '\n');
        strText = strText.toString().replace("<br>", '\n');
        strText = strText.toString().replace("<br>", '\n');

        document.theForm.txtCustomerBankTitle.value = strText;
*/        
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ///////////////////////////////////////////////////////////////////////////
    

    function getSentBySelectOptions(hudSelect){
        
        var iCusID = hudSelect.options[hudSelect.selectedIndex].value;
        //alert(iCusID);
        var strUrl = "../CRM/getSelectOption.asp?IDToken=<%=IDToken%>&SelType=2&iRecID=" + iCusID;
        
        //window.open(strUrl);
        makeRequestGetSentByOptions(strUrl, "divSentByDescription", "Post");
    }
    
	function makeRequestGetSentByOptions(url, divName, strMethod){
		var http_request = createRequestObject();
	
		http_request.onreadystatechange = function() { goActionGetSentByOptions(http_request, divName); };
        http_request.open(strMethod, url, true);
        http_request.send(null);
	}
	
	function goActionGetSentByOptions(http_request, divName) {
        if (http_request.readyState == 4) {
            if (http_request.status == 200) {
				var responceString = http_request.responseText;

                setDivActionGetSentByOptions(divName, responceString)
            } else {
                return 'There was a problem with the request.';
            }
        }
    }		

    function setDivActionGetSentByOptions(div, text){
    
        var strDivText = "Sent by:<SELECT name=selSentByID><OPTION value=-1></OPTION>" + text + "</SELECT>";

        eval("document.all." + div + ".innerHTML='" + strDivText + "';");
    }
    
    /////////////////////////////////////////////////////////////////////////////
    function addRemRows(){
        
        if(!isNumberValid(document.theForm._txtRowCount)){
            alert("Please insert number only.");
            return;
        }
        
        var iNewRowToSee = document.theForm._txtRowCount.value;
        if(parseInt(iNewRowToSee) > parseInt(<%=iMaxRowNumber%>)){
            alert("Please insert number between 0 and <%=iMaxRowNumber%>.");
            return;
        }
        document.theForm._txtInnerLineCounter.value = iNewRowToSee;
        
        for(iIdx_1=0; iIdx_1 < iNewRowToSee; iIdx_1++){
            eval("document.all.divInvoiceRow_" + iIdx_1 + ".style.display='block'");
        }

        for(iIdx_2=iNewRowToSee; iIdx_2 < <%=iMaxRowNumber%>; iIdx_2++){
            eval("document.all.divInvoiceRow_" + iIdx_2 + ".style.display='none'");
        }
    }
    
    
    function isNumberValid( str ){
		var sText = str.value;
		
		var ValidChars = "0123456789<div";
		var IsNumber=true;
		var Char;
 
		for (iCounter = 0; iCounter < sText.length && IsNumber == true; iCounter++){ 
			Char = sText.charAt(iCounter);
			if (ValidChars.indexOf(Char) == -1){
				IsNumber = false;
			}
		}

		return IsNumber;
	}

		
	function calcTotalPrice(iCounter){
		var docCarat = eval("document.theForm.txtCarat" + iCounter);
		var docPrice = eval("document.theForm.txtPrice" + iCounter);			
		var docSalesSum = eval("document.theForm.txtSalesSum" + iCounter);

        myFormatNumber(docCarat);
        myFormatNumber(docPrice);
			
			if ( docCarat.value != '' && docPrice.value != '' ){
				docSalesSum.value = parseFloat(docCarat.value) * parseFloat(docPrice.value);
				myFormatNumber(docSalesSum);
				doSumFields();
		}
	}

	function calcCaratPrice(iCounter){
		var docCarat = eval("document.theForm.txtCarat" + iCounter);
		var docPrice = eval("document.theForm.txtPrice" + iCounter);			
		var docSalesSum = eval("document.theForm.txtSalesSum" + iCounter);

        if(docPrice.value.length == 0 && docCarat.value.length != 0 && docSalesSum.value.length != 0){
            docPrice.value = docSalesSum.value / docCarat.value;
            myFormatNumber(docCarat);
            myFormatNumber(docPrice);
            myFormatNumber(docSalesSum);
            doSumFields();
        }
	}


	function doSumFields(){
		
		var iRecCount = <%=strInnerLineCounter%> + 1;

		var iSumQTY = 0;
		var iSumCarat = parseFloat(0);
		var iSumPrice = parseFloat(0);

		for(i=0; i<iRecCount; i++){
			if (eval("document.theForm.txtAmount1_" + i + ".value")){
				var dblTotalQTY1 = eval("document.theForm.txtAmount1_" + i + ".value");
				dblTotalQTY1 = dblTotalQTY1.replace(",","");
				iSumQTY = iSumQTY + parseFloat(dblTotalQTY1);
			}
			
			if (eval("document.theForm.txtCarat1_" + i + ".value")){
			    			
				var dblTotalCarat1 = eval("document.theForm.txtCarat1_" + i + ".value");
				dblTotalCarat1 = dblTotalCarat1.replace(",","");
				iSumCarat = Math.round((parseFloat(iSumCarat) + parseFloat(dblTotalCarat1)) * parseFloat(100)) / parseFloat(100);
			}

			if (eval("document.theForm.txtSalesSum1_" + i + ".value")){				
				var dblTotalPrice1 = eval("document.theForm.txtSalesSum1_" + i + ".value");
				dblTotalPrice1 = dblTotalPrice1.replace(",","");
				iSumPrice = parseFloat(iSumPrice) + parseFloat(dblTotalPrice1);
			}			
		}


        if (iSumQTY != 0){
		    document.theForm.textTotalQTYSum.value = iSumQTY;
		}else{
		    document.theForm.textTotalQTYSum.value = '';
        }
        
		document.theForm.textTotalCaratSum.value = iSumCarat;
		myFormatNumber(document.theForm.textTotalCaratSum);		
		
		document.theForm.textTotalPriceSum.value = iSumPrice;
		myFormatNumber(document.theForm.textTotalPriceSum);
		
	    if(document.theForm.textShippingCost.value.length != 0){
	        document.theForm.textTotalPriceSumAfterShippingCost.value = parseFloat(iSumPrice) + parseFloat(document.theForm.textShippingCost.value);
	        myFormatNumber(document.theForm.textShippingCost);
	        myFormatNumber(document.theForm.textTotalPriceSumAfterShippingCost);
	    }else{
	        document.theForm.textTotalPriceSumAfterShippingCost.value = parseFloat(iSumPrice);
	        myFormatNumber(document.theForm.textTotalPriceSumAfterShippingCost);
	    }
	}
	
	
	function myFormatNumber(txtHundle){
		var iFldValue = txtHundle.value;
		var iDotIndex = iFldValue.indexOf('.');
		
		if(iDotIndex != -1){
			var strPre = iFldValue.substring(0, iDotIndex);
			var strSuf = iFldValue.substring(iDotIndex + 1);
			if(strSuf.length == 0){
				txtHundle.value = strPre + ".00";
			}
			if(strSuf.length == 1){
				txtHundle.value = strPre + "." + strSuf + "0";
			}
			if(strSuf.length > 2){
				txtHundle.value = strPre + "." + strSuf.substring(0, 2);
			}
		}else{
			if(txtHundle.value.length > 0)
				txtHundle.value = txtHundle.value + ".00";
		}
	}
	
	
	function fillDivLikeTextArea(hunTextArea, strTextArea, strTitle)
	{
		var divHtml = strTextArea;
		
		divHtml = "<font size=2 face=arial color=black>" + strTitle + divHtml + "</font>";
		divHtml = divHtml.replace(new RegExp( "\\n", "g" ), "<br>");
		
		hunTextArea.innerHTML = divHtml;
	}
	
	function ShowPrintDocumentTotal(){
	    
	
		document.all.divTotalMode.style.display='none';

	    ShowPrintDocument();
	}
	
	function prepareCustomerDetail(){
	    var strData = "";
	    var bAddTo = false;

	    if(document.all.cbCustomerName.checked){
	        strData = strData + document.all.txtCustomerName.value;
	        if(document.all.txtCustomerName.value.length > 0){
    	        strData = strData + "<br>";
    	        bAddTo = true;
    	    }
	    }

	    if(document.all.cbCustomerAddress.checked){
	        strData = strData + document.all.txtCustomerAddress1.value;
	        if(document.all.txtCustomerAddress1.value.length > 0){
    	        strData = strData + "<br>";
    	        bAddTo = true;
    	    }
	        strData = strData + document.all.txtCustomerAddress2.value;
	        if(document.all.txtCustomerAddress3.value.length > 0){
    	        strData = strData + "<br>";
    	        bAddTo = true;
    	    }
	        strData = strData + document.all.txtCustomerAddress3.value;
	        if(document.all.txtCustomerAddress3.value.length > 0){
    	        strData = strData + "<br>";
    	        bAddTo = true;
    	    }
	    }

        if(bAddTo){
            strData = "To:<br>" + strData;
        }
        
        strData = "<font face='Arial, Helvetica, sans-serif' size='2' color='black'>" + strData + "</font>";
        
	    document.all.divShowCustomer_2.innerHTML = strData;
	}
	
	function ShowPrintDocument(){
	    
	    window.open ("../CRM/blank.html", 'PrintInvoice', 'toolbar=no, location=no, status=no, menubar=no, scrollbars=no,resizable=no, width=800, height=610');
	    
	    setTimeout("document.theForm.submit();", 1000);
	    return;
	    alert()
	    
	    <%If ShowLogo <> "false" Then%>
	    if(document.all.cbShowLogo.checked){
	        document.all.divShowLogo_1.style.display='none';
	        document.all.divShowLogo_2.style.display='block';
        }else{
	        document.all.divShowLogo_1.style.display='none';
	        document.all.divShowLogo_2.style.display='none';
        }
        <%End If%>
        prepareCustomerDetail();
        document.all.divShowCustomer_1.style.display='none';
        document.all.divShowCustomer_2.style.display='block';
        
        
	    var iNewRowToSee = document.theForm._txtRowCount.value;

		var iRecCount = iNewRowToSee;
		var iCounterCatalogsTypedByUsers = 0;
		
		for(i=0; i<iRecCount; i++){		

			eval("document.theForm.txtCatalog1_" + i + ".readOnly=true;");
			eval("document.theForm.txtAmount1_" + i + ".readOnly=true;");
			eval("document.theForm.txtCarat1_" + i + ".readOnly=true;");
			eval("document.theForm.txtDescription1_" + i + ".readOnly=true;");			
			eval("document.theForm.txtSalesSum1_" + i + ".readOnly=true;");
			eval("document.theForm.txtPrice1_" + i + ".readOnly=true;");

			eval("document.theForm.txtCatalog1_" + i + ".className='noBorder';");
			eval("document.theForm.txtAmount1_" + i + ".className='noBorder';");
			eval("document.theForm.txtCarat1_" + i + ".className='noBorder';");
			eval("document.theForm.txtDescription1_" + i + ".className='noBorder';");
			eval("document.theForm.txtSalesSum1_" + i + ".className='noBorder';");
			eval("document.theForm.txtPrice1_" + i + ".className='noBorder';");

			var x = eval("document.theForm.txtCatalog1_" + i + ".value;");
			if (x.length > 0)
				iCounterCatalogsTypedByUsers++;
		}
		
		if (iCounterCatalogsTypedByUsers == 0){
		    document.all.divCatalogueTitleLable.style.display='none';
       		document.all.divCatalogueTitleLable1.style.display='block';
		}
		
		var strInvoiceNO = document.all.txtInvoiceNo.value;
		
	    document.all.divInvoiceTitle_1.style.display='none';
   		document.all.divInvoiceTitle_2.style.display='block';

		if(strInvoiceNO.length == 0){
		    document.all.divInvoiceTitle_2.innerHTML = "";
		}else{
		    document.all.divInvoiceTitle_2.innerHTML = "<font size='3' face='arial' color='black'><b>&nbsp;&nbsp;NO. &nbsp;&nbsp;" + strInvoiceNO + "</b></font>";
		}
        <%If Len(Request.Form("txtDestCustomerID")) > Int(0) Then %>		
        <%Else%>
		    eval("document.theForm.txtSRToCompanyTitle.readOnly=true;");
		    
		    eval("document.theForm.txtCustomerSerialNumber1.readOnly=true;");
		    eval("document.theForm.txtCustomerSerialNumber2.readOnly=true;");
    		
		    eval("document.theForm.txtSRToCompanyTitle.style.overflow='hidden';");
		    eval("document.theForm.txtSRToCompanyTitle.style.bordercolor='#FFFFFF';");
		    eval("document.theForm.txtSRToCompanyTitle.style.border='none';");
		    //border-style: none; border-color: #FFFFFF; overflow: hidden;
		    eval("document.theForm.txtCustomerSerialNumber1.className='noBorder';");
		    eval("document.theForm.txtCustomerSerialNumber2.className='noBorder';");
		    if (document.theForm.txtCustomerSerialNumber1.value == '' && document.theForm.txtCustomerSerialNumber2.value == '')
				document.all.divCustomerSerialNumber.style.display='none';
				
		    eval("document.all.txtSRToCompanyTitle.style.display='none';");
		    eval("document.all.divSRToCompanyTitle.style.display='block';");
		    eval("document.all.divSRToCompanyTitleTemp.style.display='none';");
		    
		<%End If%>

		<%If Len(strBank) > Int(0) Then %>
		    eval("document.theForm.txtCustomerBankTitle.style.overflow='hidden';");
		    eval("document.theForm.txtCustomerBankTitle.style.bordercolor='#FFFFFF';");
		    eval("document.theForm.txtCustomerBankTitle.style.border='none';");
			eval("document.theForm.txtCustomerBankTitle.readOnly=true;");
			eval("document.all.txtCustomerBankTitle.style.display='none';");
			eval("document.all.divCustomerBankTitle.style.display='block';");
		<%End If%>

        document.all.divNumbersOfRows.style.display='none';

        if(document.theForm.txtLegalRemark.value.length > 0){
            document.all.divLegalRemarkLable.innerHTML = "<font face=arial  size=2>Remarks: " + document.theForm.txtLegalRemark.value + "</font><br><br>";
            eval("document.all.divLegalRemarkLable.style.display='block';")
        }else{
            document.all.divLegalRemarkLable.innerHTML = "";
        }
        eval("document.all.divLegalRemark.style.display='none';")



		    eval("document.theForm.textShippingCost.readOnly=true;");
		    eval("document.theForm.textShippingCost.className='noBorder';");
		
        if(document.theForm.textShippingCost.value.length == 0 || parseInt(document.theForm.textShippingCost.value, 10)==parseInt(0, 10)){
            eval("document.all.divShippingLable.style.display='none';")
            eval("document.all.divShippingValue.style.display='none';")
            eval("document.all.divSecounderyTotalLable.style.display='none';")
            eval("document.all.divSecounderyTotalValue.style.display='none';")

            eval("document.all.divShippingLable1.style.display='block';")
            eval("document.all.divShippingValue1.style.display='block';")
            eval("document.all.divSecounderyTotalLable1.style.display='block';")
            eval("document.all.divSecounderyTotalValue1.style.display='block';")
        }

		document.theForm.butPrt.disabled = true;
		document.all.divPrintBtn.style.display = 'none';
		window.print();			
	}
	
	function ShowReshomonDocument(){
        var strUrl = "";

        var iSelectedIndex;
        var iSelectedValue;
        var arrSplit;

	    strUrl = './PrtStockShippingDetailReshomonManual.asp?IDToken=<%=IDToken%>';
        iSelectedIndex = document.theForm.selCompanyTo.selectedIndex;
        iSelectedValue = document.theForm.selCompanyTo.options[iSelectedIndex].value;
        
        //arrSplit = iSelectedValue.split("_");
	    
	    strUrl = strUrl + '&selDest=' + iSelectedValue;//arrSplit[0];
	    strUrl = strUrl + '&txtDestCustomerID=' + document.theForm.hidDestCustomerID.value;
	    
	    strUrl = strUrl + '&txtShippingCost=';// + document.theForm.textShippingCost.value;

        iSelectedIndex = document.theForm.selCustomerFrom.selectedIndex;
        iSelectedValue = document.theForm.selCustomerFrom.options[iSelectedIndex].value;
        arrSplit = iSelectedValue.split("_");
	    strUrl = strUrl + '&txtFromCustomerID=' + arrSplit[1];
	    strUrl = strUrl + '&selFrom=' + arrSplit[0];


	    strUrl = strUrl + '&selDay=' + document.theForm.selDay.options[document.theForm.selDay.selectedIndex].value;
	    strUrl = strUrl + '&selMonth=' + document.theForm.selMonth.options[document.theForm.selMonth.selectedIndex].value;
	    strUrl = strUrl + '&selYear=' + document.theForm.selYear.options[document.theForm.selYear.selectedIndex].value;

	    strUrl = strUrl + '&BayerID=' + document.theForm.txtInvoiceNo.value;

	    strUrl = strUrl + '&selBankID=' + document.theForm.selBankID.options[document.theForm.selBankID.selectedIndex].value;
	    if(document.theForm.selCustomerFrom.selectedIndex == 0)
	        strUrl = strUrl + '&ShowFrom=false';
	    else
	        strUrl = strUrl + '&ShowFrom=true';


        if(document.theForm.taCustomerDesc.value.length == 0)
	        strUrl = strUrl + '&ShowCustomer=false';
	    else
	        strUrl = strUrl + '&ShowCustomer=true';

	    strUrl = strUrl + '&SentByID=' + document.theForm.selSentByID.options[document.theForm.selSentByID.selectedIndex].value;
	    

	    strUrl = strUrl + '&TotalBalanceCarat=' + document.theForm.textTotalCaratSum.value;
		<%If Len(iShippingCost) > Int(0) Then%>
		
		    //strUrl = strUrl + '&iLogicSalePricetSum=' + document.theForm.textTotalPriceSumAfterShippingCost.value;
		<%Else%>
		    //strUrl = strUrl + '&iLogicSalePricetSum=' + document.theForm.textTotalPriceSum.value;
		<%End If%>

		strUrl = strUrl + '&iLogicSalePricetSum=' + document.theForm.textTotalPriceSumAfterShippingCost.value;

	    window.open(strUrl, '', 'toolbar=no, location=no, status=no, menubar=no, scrollbars=yes,resizable=no, width=730, height=700');
	}
	
	function ShowShipperDocument(){
	    //strUrl = strUrl + '&selDest=<%=strToCompanyID%>';
	    //strUrl = strUrl + '&txtDestCustomerID=<%=strDestCustomerID%>';
	    //strUrl = strUrl + '&txtShippingCost=<%=iShippingCost%>';

	    //strUrl = strUrl + '&txtFromCustomerID=<%=strCustomerID%>';
	    //strUrl = strUrl + '&selFrom=<%=strFromCompany%>';	    

	    //strUrl = strUrl + '&selDay=<%=iTheDay%>';
	    //strUrl = strUrl + '&selMonth=<%=iTheMonth%>';
	    //strUrl = strUrl + '&selYear=<%=iTheYear%>';

        //strUrl = strUrl + '&selBankID=<%=strBank%>';
        
	    //strUrl = strUrl + '&ShowFrom=<%=strFrom%>';
	    //strUrl = strUrl + '&ShowCustomer=<%=strShowCustomerName%>';

        var strUrl = "";

        var iSelectedIndex;
        var iSelectedValue;
        var arrSplit;

        iSelectedIndex = document.theForm.selCompanyTo.selectedIndex;
        iSelectedValue = document.theForm.selCompanyTo.options[iSelectedIndex].value;
        
        //arrSplit = iSelectedValue.split("_");
	    
	    //strUrl = strUrl + '&selDest=' + iSelectedValue;//arrSplit[0];
	    //strUrl = strUrl + '&txtDestCustomerID=' + document.theForm.hidDestCustomerID.value;



	    strUrl = './PrtStockShippingDetailShipperLetterManual.asp?IDToken=<%=IDToken%>';
        iSelectedIndex = document.theForm.selCompanyTo.selectedIndex;
        iSelectedValue = document.theForm.selCompanyTo.options[iSelectedIndex].value;
        //arrSplit = iSelectedValue.split("_");
	    strUrl = strUrl + '&selDest=' + iSelectedValue;//arrSplit[0];
	    strUrl = strUrl + '&txtDestCustomerID=' + document.theForm.hidDestCustomerID.value;//arrSplit[1];
	    strUrl = strUrl + '&hidCustomerDesc=' + document.theForm.hidCustomerDesc.value;//arrSplit[1];
	    
	    
	    
	    
	    strUrl = strUrl + '&txtShippingCost=';// + document.theForm.textShippingCost.value;


        iSelectedIndex = document.theForm.selCustomerFrom.selectedIndex;
        iSelectedValue = document.theForm.selCustomerFrom.options[iSelectedIndex].value;
        arrSplit = iSelectedValue.split("_");
	    strUrl = strUrl + '&txtFromCustomerID=' + arrSplit[1];
	    strUrl = strUrl + '&selFrom=' + arrSplit[0];


	    strUrl = strUrl + '&selDay=' + document.theForm.selDay.options[document.theForm.selDay.selectedIndex].value;
	    strUrl = strUrl + '&selMonth=' + document.theForm.selMonth.options[document.theForm.selMonth.selectedIndex].value;
	    strUrl = strUrl + '&selYear=' + document.theForm.selYear.options[document.theForm.selYear.selectedIndex].value;

   	    strUrl = strUrl + '&selBankID=' + document.theForm.selBankID.options[document.theForm.selBankID.selectedIndex].value;
   	    strUrl = strUrl + '&hidCustomerBankTitle=' + document.theForm.hidCustomerBankTitle.value;


        strUrl = strUrl + '&selCusBankDesc=' +  document.theForm.txtCustomerBankTitle.value;
        
		<%If Len(strBank) > Int(0) Then %>
		    strUrl = strUrl + '&selCusBankDesc=' + document.all.divCustomerBankTitle.innerHTML;
		<%Else%>
		    strUrl = strUrl + '&selCusBankDesc=';
		<%End If%>

		<%If Len(Request.Form("txtDestCustomerID")) = Int(0) Then %>
		    //strUrl = strUrl + '&selSRToCompanyDesc=' + document.all.divSRToCompanyTitle.innerHTML;
		<%Else%>
		    strUrl = strUrl + '&selSRToCompanyDesc=';
		<%End If%>
		
		
	    strUrl = strUrl + '&txtInvoiceRemark=' + document.all.txtLegalRemark.value;


        if(document.theForm.selCustomerFrom.selectedIndex == 0)
            strUrl = strUrl + '&ShowFrom=false';
        else
            strUrl = strUrl + '&ShowFrom=true';

        if(document.theForm.taCustomerDesc.value.length == 0)
            strUrl = strUrl + '&ShowCustomer=false';
        else
            strUrl = strUrl + '&ShowCustomer=true';

	    //strUrl = strUrl + '&BayerID=<%=strInvoiceNO%>';
	    strUrl = strUrl + '&TotalBalanceCarat=' + document.theForm.textTotalCaratSum.value;
		<%If Len(iShippingCost) > Int(0) Then%>
		    strUrl = strUrl + '&iLogicSalePricetSum=' + document.theForm.textTotalPriceSumAfterShippingCost.value;
		<%Else%>
		    strUrl = strUrl + '&iLogicSalePricetSum=' + document.theForm.textTotalPriceSum.value;
		<%End If%>	    
	    window.open(strUrl, '', 'toolbar=no, location=no, status=no, menubar=no, scrollbars=yes,resizable=no, width=740, height=700');
	}

	
	function ShowPrePrintDocument(){
	    document.all.divTotalMode.style.display='block';
	    
        <%If ShowLogo <> "false" Then%>
     	    document.all.divShowLogo_1.style.display='block';
	        document.all.divShowLogo_2.style.display='none';
	    <%End If%>
        document.all.divShowCustomer_1.style.display='block';
        document.all.divShowCustomer_2.style.display='none';

        var iNewRowToSee = document.theForm._txtRowCount.value;
        var iRecCount = iNewRowToSee;

		for(i=0; i<iRecCount; i++){		

			eval("document.theForm.txtCatalog1_" + i + ".readOnly=false;");
			eval("document.theForm.txtAmount1_" + i + ".readOnly=false;");
			eval("document.theForm.txtCarat1_" + i + ".readOnly=false;");
			eval("document.theForm.txtDescription1_" + i + ".readOnly=false;");
			eval("document.theForm.txtSalesSum1_" + i + ".readOnly=false;");
			eval("document.theForm.txtPrice1_" + i + ".readOnly=false;");

			eval("document.theForm.txtCatalog1_" + i + ".className='noError';");
			eval("document.theForm.txtAmount1_" + i + ".className='noError';");
			eval("document.theForm.txtCarat1_" + i + ".className='noError';");
			eval("document.theForm.txtDescription1_" + i + ".className='noError';");
			eval("document.theForm.txtSalesSum1_" + i + ".className='noError';");
			eval("document.theForm.txtPrice1_" + i + ".className='noError';");
		}

   		document.all.divInvoiceTitle_1.style.display='block';
	    document.all.divInvoiceTitle_2.style.display='none';
		
		document.all.divCatalogueTitleLable.style.display='block';
		document.all.divCatalogueTitleLable1.style.display='none';

		document.all.divNumbersOfRows.style.display='block';


        <%If Len(Request.Form("txtDestCustomerID")) > Int(0) Then %>		
        <%Else%>
		    eval("document.theForm.txtSRToCompanyTitle.readOnly=false;");


		    eval("document.theForm.txtCustomerSerialNumber1.readOnly=false;");
		    eval("document.theForm.txtCustomerSerialNumber2.readOnly=false;");


		    eval("document.theForm.txtSRToCompanyTitle.style.overflow='auto';");
		    eval("document.theForm.txtSRToCompanyTitle.style.bordercolor='grey';");
		    eval("document.theForm.txtSRToCompanyTitle.style.border='1px solid grey';");

		    eval("document.theForm.txtCustomerSerialNumber1.className='noError';");
		    eval("document.theForm.txtCustomerSerialNumber2.className='noError';");
		    eval("document.all.divCustomerSerialNumber.style.display='block';");
		    
		    eval("document.all.divSRToCompanyTitle.style.display='none';");		    
		    eval("document.all.divSRToCompanyTitleTemp.style.display='block';");		    
		    eval("document.all.txtSRToCompanyTitle.style.display='block';");
		<%End If%>

		<%If Len(strBank) > Int(0) Then %>
		    eval("document.theForm.txtCustomerBankTitle.readOnly=false;");

		    eval("document.theForm.txtCustomerBankTitle.style.overflow='auto';");
		    eval("document.theForm.txtCustomerBankTitle.style.bordercolor='grey';");
		    eval("document.theForm.txtCustomerBankTitle.style.border='1px solid grey';");

			eval("document.all.divCustomerBankTitle.style.display='none';");
			eval("document.all.txtCustomerBankTitle.style.display='block';");
		<%End If%>

        document.all.divLegalRemarkLable.innerHTML = "<font face=arial  size=2>Remarks:</font>";
        document.all.divLegalRemarkLable.style.display='block';

        eval("document.all.divLegalRemark.style.display='block';")

		<%If Len(iShippingCost) > Int(0) Then %>
		    eval("document.theForm.textShippingCost.readOnly=false;");
		    eval("document.theForm.textShippingCost.className='noError';");
		<%End If%>

            eval("document.all.divShippingLable.style.display='block';")
            eval("document.all.divShippingValue.style.display='block';")
            eval("document.all.divSecounderyTotalLable.style.display='block';")
            eval("document.all.divSecounderyTotalValue.style.display='block';")

            eval("document.all.divShippingLable1.style.display='none';")
            eval("document.all.divShippingValue1.style.display='none';")
            eval("document.all.divSecounderyTotalLable1.style.display='none';")
            eval("document.all.divSecounderyTotalValue1.style.display='none';")

		document.theForm.butPrt.disabled = false;
		document.all.divPrintBtn.style.display = 'block';
	}	
</SCRIPT>





<!--FORM action="./PrtStockShippingDocumentCheckManual.asp?IDToken=<%=IDToken%>&ParentRecType=68" method="Post" name=theBackForm>
    <INPUT type=hidden name="_logo" value="<%=Request.Form("logo")%>">
    <INPUT type=hidden name="_customer" value="<%=Request.Form("customer")%>">
    <INPUT type=hidden name="_address" value="<%=Request.Form("address")%>">
    <INPUT type=hidden name="_from" value="<%=Request.Form("from")%>">


    <INPUT type=hidden name="_selDay" value="<%=Request.Form("selDay")%>">
    <INPUT type=hidden name="_selMonth" value="<%=Request.Form("selMonth")%>">
    <INPUT type=hidden name="_selYear" value="<%=Request.Form("selYear")%>">

    <INPUT type=hidden name="_selFrom" value="<%=Request.Form("selFrom")%>">
    <INPUT type=hidden name="_txtFromStockCusID" value="<%=Request.Form("txtFromStockCusID")%>">
    <INPUT type=hidden name="_txtFromCustomerID" value="<%=Request.Form("txtFromCustomerID")%>">

    <INPUT type=hidden name="_selDest" value="<%=Request.Form("selDest")%>">
    <INPUT type=hidden name="_txtDestStockCusID" value="<%=Request.Form("txtDestStockCusID")%>">
    <INPUT type=hidden name="_txtDestCustomerID" value="<%=Request.Form("txtDestCustomerID")%>">

    <INPUT type=hidden name="_selSentBy" value="<%=Request.Form("selSentBy")%>">
    <INPUT type=hidden name="_selInvoiceType" value="<%=Request.Form("selInvoiceType")%>">

    <INPUT type=hidden name="_txtInvoiceNO" value="<%=Request.Form("txtInvoiceNO")%>">
    <INPUT type=hidden name="_txtShippingCost" value="<%=Request.Form("txtShippingCost")%>">

    <INPUT type=hidden name="_selBankID" value="<%=Request.Form("selBankID")%>">
    <INPUT type=hidden size=3 name="_txtInnerLineCounter" value="<%=Request.Form("txtInnerLineCounter")%>">
</form-->

<FORM action="./doPrtStockShippingDetailManual_New.asp?IDToken=<%=IDToken%>&RecID=<%=strRecID%>" method="Post" name=theForm target="PrintInvoice">
	<input type="hidden" name="iApprover">
	
    <INPUT type=hidden size=3 name="_txtInnerLineCounter" value="<%=Request.Form("txtInnerLineCounter")%>">
	
	
	
<%

	'Added By Ziv -- 11.08.2009 -- Start
	Function getCustomerDetail(iCustomerID, strWhatToOutput)
		If Len(Trim(iCustomerID)) > Int(0) And Len(Trim(strWhatToOutput)) > Int(0) Then
			If LCase(strWhatToOutput) = LCase("Address") Then
				'strSql = "Select Address1, Address2, Address3, Address4 From " & strCustomerDB_Alias & "CustomerCustomer Where ID = " & iCustomerID
				'doQuery(strSql)
				strAddress1Tmp = getDBField("Select Address1 From " & strCustomerDB_Alias & "CustomerCustomer Where ID = " & iCustomerID, "Address1")
				strAddress1Tmp = "<input id='cbCustomerName' type='checkbox' checked><input id='txtCustomerName' style='border:none;background-color:white;' type='text' size=70 value='" & strAddress1Tmp & "'>"
				strAddress2Tmp = "<input id='cbCustomerAddress' type='checkbox' checked><input id='txtCustomerAddress1' style='border:none;background-color:white;' type='text' size=70 value='" & getDBField("Select Address2 From " & strCustomerDB_Alias & "CustomerCustomer Where ID = " & iCustomerID, "Address2") & "'>"
				strAddress3Tmp = "<input id='txtCustomerAddress2' style='border:none;background-color:white;' type='text' size=70 value='" & getDBField("Select Address3 From " & strCustomerDB_Alias & "CustomerCustomer Where ID = " & iCustomerID, "Address3") & "'>"
				strAddress4Tmp = "<input id='txtCustomerAddress3' style='border:none;background-color:white;' type='text' size=70 value='" & getDBField("Select Address4 From " & strCustomerDB_Alias & "CustomerCustomer Where ID = " & iCustomerID, "Address4") & "'>"
			
				'If rs.RecordCount <> 0 And Not rs.EOF Then
					'getCustomerDetail = rs("Address1") & "<BR>" & rs("Address2") & "<BR>" & rs("Address3") & "<BR>" & rs("Address4")
					getCustomerDetail = strAddress1Tmp & "<BR>" & strAddress2Tmp & "<BR>" & strAddress3Tmp & "<BR>" & strAddress4Tmp
				'End If
			Else
				If LCase(strWhatToOutput) = LCase("CustomerName") Then				
					'strSql = " Select CompanyName From " & strCustomerDB_Alias & "CustomerSRCompany Where ID In (Select Top 1 SRCompanyID From " & strCustomerDB_Alias & "CustomerCustomerSRCompany Where CustomerID = " & iCustomerID & ")"
					'doQuery(strSql)
					strCompanyNameTmp = getDBField("Select CompanyName From " & strCustomerDB_Alias & "CustomerSRCompany Where ID In (Select Top 1 SRCompanyID From " & strCustomerDB_Alias & "CustomerCustomerSRCompany Where CustomerID = " & iCustomerID & ")", "CompanyName")
			
					'If rs.RecordCount <> 0 And Not rs.EOF Then
						'getCustomerDetail = rs("CompanyName")
						getCustomerDetail = strCompanyNameTmp
					'End If
				Else
					getCustomerDetail = ""
				End If
			End If
		Else
			getCustomerDetail = ""
		End If
	End Function

	Dim iQuantitySum
		iQuantitySum = 0
	Dim iCaratSum
		iCaratSum = 0	
	Dim iCaratPriceSum
		iCaratPriceSum = 0
			
	Dim iTotalSum
		iTotalSum = 0
		
	Dim iTotal		
		iTotal = 0

		strTitle = "Invoice:"
		strText = "<TABLE WIDTH=95% BORDER=0 CELLSPACING=0 CELLPADDING=0 bordercolor=red>"

		strText = strText & "<TR>"
		strText = strText & "<TD valign=top  colspan=2 nowrap>"
		If Len(Trim(Request.Form("txtLegalRemark"))) > Int(0) Then
		strText = strText & "<font face=arial  size=2><b>Remarks:</b></font><br>"
		strText = strText & "<font face=arial  size=2>" & Request.Form("txtLegalRemark") & "</font><br><br>"
		End If
		strText = strText & "</TD>"

        If (iFromSRCompany = 2 Or iFromSRCompany = 5 Or iFromSRCompany = 34) And Request.Form("selShowCondition") = 1 Then
		    strText = strText & "<TR>"
		    strText = strText & "<TD width=18 valign=top><font face=arial color=black size=2>1)</TD>"
		    strText = strText & "<TD valign=top><font face=arial color=black size=2>Country of origin - Israel.<br></TD>"
		    strText = strText & "</TR>"
		    strText = strText & "<TR>"
		    strText = strText & "<TD width=18 valign=top><font face=arial color=black size=2>2)</TD><TD valign=top><font face=arial color=black size=2>The diamonds herein invoiced have been purchased from legitimate sources not involved in funding "
		    strText = strText & "conflicts and in compliance with United Nations resolutions. <BR>We hereby guarantee that these diamonds "
		    strText = strText & "are conflict free, based on personal knowledge and/or written guarantee by the supplier of "
		    strText = strText & "these diamonds.</TD>"
		    strText = strText & "</TR>"
		    strText = strText & "<TR>"
		    strText = strText & "<TD valign=top><font face=arial color=black size=2>3)</TD><TD valign=top><font face=arial color=black size=2>The goods remain the legal property of the seller until the goods have been paid in full.</TD>"  
		    strText = strText & "</TR>"

            'Response.Write "<p dir=ltr>" & Request.Form("selBoxID") & "</p>"
            arrBoxAddress = Split(Request.Form("selBoxID"), "_")
            
            strText4 = ""
            If UBound(arrBoxAddress) >= 2 Then

                strText4 = strText4 & "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>"
                strText4 = strText4 & "<TR>"
                strText4 = strText4 & "    <TD colspan=4><font face=arial color=black size=2><u>Payment Instructions</u></font></TD>"
                strText4 = strText4 & "</TR>"
                strText4 = strText4 & "<TR>"
                strText4 = strText4 & "    <TD>"
                strText4 = strText4 & "        <font face=arial color=black size=2>Due date:</font>&nbsp;"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "    <TD>"
                strText4 = strText4 & "       <font face=arial color=black size=2>" & Request.Form("selDueDateDay") & "/" & Request.Form("selDueDateMonth") & "/" & Request.Form("selDueDateYear") & "</font>"
                strText4 = strText4 & "       &nbsp;&nbsp;&nbsp;&nbsp;"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "    <TD>"
                strText4 = strText4 & "        <font face=arial color=black size=2>amount:</font>&nbsp;"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "    <TD>"
                strText4 = strText4 & "        <font face=arial color=black size=2>" & MyFormatNumber(Request.Form("textTotalPriceSumAfterShippingCost_4Rem")) & "</font>"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "   <TR>"

                strText4 = strText4 & "   </TR>"
                strText4 = strText4 & "    <TD>"
                strText4 = strText4 & "        <font face=arial color=black size=2>Bank account:</font>&nbsp;"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "    <TD colspan=3>"
                strText4 = strText4 & "     <font face=arial color=black size=2>" & arrBoxAddress(1) & "</font>"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "</TR>"

                strText4 = strText4 & "</TR>"
                strText4 = strText4 & "    <TD>&nbsp;</TD>"
                strText4 = strText4 & "    <TD colspan=3>"
                strText4 = strText4 & "     <font face=arial color=black size=2>" & arrBoxAddress(2) & "</font>"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "</TR>"
If UBound(arrBoxAddress) > 2 Then
                strText4 = strText4 & "</TR>"
                strText4 = strText4 & "    <TD>&nbsp;</TD>"
                strText4 = strText4 & "    <TD colspan=3>"
                strText4 = strText4 & "     <font face=arial color=black size=2>" & arrBoxAddress(3) & "</font>"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "</TR>"
End If
If UBound(arrBoxAddress) > 3 Then
                strText4 = strText4 & "</TR>"
                strText4 = strText4 & "    <TD>&nbsp;</TD>"
                strText4 = strText4 & "    <TD colspan=3>"
                strText4 = strText4 & "     <font face=arial color=black size=2>" & arrBoxAddress(4) & "</font>"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "</TR>"
End If
If UBound(arrBoxAddress) > 4 Then
                strText4 = strText4 & "</TR>"
                strText4 = strText4 & "    <TD>&nbsp;</TD>"
                strText4 = strText4 & "    <TD colspan=3>"
                strText4 = strText4 & "     <font face=arial color=black size=2>" & arrBoxAddress(5) & "</font>"
                strText4 = strText4 & "    </TD>"
                strText4 = strText4 & "</TR>"
End If
                strText4 = strText4 & "</TABLE>"                
                strText4 = Replace(strText4, "<geres>", "'")
                
		        strText = strText & "<TR>"
		        strText = strText & "<TD valign=top><font face=arial color=black size=2>4)</TD>"
                strText = strText & "<TD valign=top>"
                strText = strText & strText4
                strText = strText & "</TD>"
                
            End If
            If iFromSRCompany = 34 Then
                'strText = strText4
            End If
        End If
		strText = strText & "</TR>"
		strText = strText & "</TABLE>"
		strBottom = strButtomCompanyName & "&nbsp;&nbsp;<br>"
		iTableWidth = "100%"
	

%>


<TABLE WIDTH="600"  BORDER=0 CELLSPACING=0 CELLPADDING=0 <%=strDirTable%> align=center> <%'הטבלה התוחמת הכי חיצונית %>
<TR>
	<TD>
        <TABLE WIDTH="100%"  BORDER=0 CELLSPACING=4 CELLPADDING=0 <%=strDirTable%> align=center>
        <TR>
       	    <%If Int(Request.Form("selShowLogo")) = Int(1) And Request.Form("selCustomerFrom") <> "0_0" Then%>
	        <TD width="50%" align=left>
        	    <img src="../Images/sr_logo_300X160.jpg">
            </TD>
       	    <%End If%>
	        <TD <%If Int(Request.Form("selShowLogo")) = Int(1) Then%>width="50%"<%Else%>width="100%" colspan=2<%End If%> align=center>
                <font face=arial  size=2>
                <%strCompSrcDesc = getExportFromDescription(Request.Form("selCustomerFrom"))%>
                <%=strCompSrcDesc%>
                </font>
            </TD>
        </TR>
        <TR>
	        <TD colspan=2 width="100%" align=center>
	            <br>
				<font face=arial size=3>
				<b>
                <%=SetInvoiceTitle(Request.Form("selInvoiceType"), Request.Form("txtInvoiceNo"))%>
                </b>
                </font>
                <br><br>
            </TD>
        </TR>
        <TR>
	        <TD width="50%">
                <%STRFormatDestCustomerDescription = FormatDestCustomerDescription(Request.Form("hidCustomerDesc"), Request.Form("selShowCustomerName"), Request.Form("selShowCustomerAddress"))%>


	            <%If Len(STRFormatDestCustomerDescription) > Int(0) Then %>
                <font face=arial  size=2>To:<br>
                    <%=Replace(STRFormatDestCustomerDescription, "_AMP_", "&")%>
                </font>
                <%End If%>
            </TD>
   	        <TD width="50%" align="right">
                <%
                    If Len(Trim(Request.Form("selCompanyTo"))) > Int(0) Then
                    
                        If Int(Request.Form("selCompanyTo")) <> Int(-1) And Len(Trim(Request.Form("txtFourDigits"))) = Int(4) Then
                        
                            strDestCustomerSerialNumber = getDBField("Select Description From CustomerSRCompany Where ID = " & Request.Form("selCompanyTo"), "Description") & "-" & Request.Form("txtFourDigits")
                        End If
                    End If
                %>
                
                <font face=arial  size=2>
                    <%=FormatDateToString(Request.Form("selDay"), Request.Form("selMonth"), Request.Form("selYear"))%>
                    <br>
                    <%=strDestCustomerSerialNumber%>
                </font>

            </TD>
        </TR>
        <TR>
	        <TD colspan=2 width="100%" align=left>
	            <font face="Arial, Helvetica, sans-serif" size="2" color="black">
				    <%If Int(Request.Form("selBankID")) <> Int(-1) Then %>
				        <%strBankDescription =getDBField("Select DescriptionEng From CustomerBanks Where ID = " & Request.Form("selBankID"), "DescriptionEng")%>
			            <br><font face="Arial, Helvetica, sans-serif" size="2" color="black">For one parcel of polished diamonds despatched to you on behalf by <%=UCase(strBankDescription)%> Diamond Exchange Branch, Israel through:</font><br><br>
				    <%End If %>
                    
                    <%=Request.Form("hidCustomerBankTitle")%>
	            </font>
            </TD>
        </TR>
        </TABLE>
    </TD>
</TR>
<TR>
	<TD>
		<br>
		
		<TABLE WIDTH=<%=iTableWidth%>  BORDER=0 CELLSPACING=0 CELLPADDING=0 <%=strDirTable%> align=center>		
		<TR>
			<TD>
				<TABLE WIDTH="644" border=0 bordercolor=black cellpadding="0" cellspacing="0"  valign="top">
				<TR>
				    <TD><img src="../Images/Blankt.gif" height="10" width="40" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="91" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="78" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="100" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="146" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="90" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="97" border="0"></TD>
				</TR>
				<TR>
					<TD bgcolor="white" align="center" <%If strDocType <> "total" then%>class="topbuttomleftblack"<%Else%>class="lefttopborderblack"<%End If%>><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b><%=strRowNumTitle%></b></font></TD>
					<TD bgcolor="white" <%If Len(Trim(Request.Form("textTotalQTYSum"))) = Int(0) Then%>colspan=2<%End If%> align="center" height=50 <%If strDocType <> "total" then%>class="topbottomBlack"<%Else%>class="topBorderBlack"<%End If%>>
					    &nbsp;
					    <font id="CatalogTitle" face="Arial, Helvetica, sans-serif" size="2" color="black">
					    <%If strDocType <> "total" And Int(Request.Form("selShowRowCatalog")) = Int(1) then%><b><div id='divCatalogueTitleLable' style='display:block;'>Catalogue</div></b><%End If%></font>
					    <div id='divCatalogueTitleLable1' style='display:none;'>&nbsp;</div>
                    </TD>
                    <%If Len(Trim(Request.Form("textTotalQTYSum"))) > Int(0) Then%>
					<TD bgcolor="white" align="center" <%If strDocType <> "total" then%>class="topbuttomleftblack"<%Else%>class="lefttopborderblack"<%End If%>><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>QTY</b></font></TD>
					<%End If%>
					<TD bgcolor="white" align="center" <%If strDocType <> "total" then%>class="topbuttomleftblack"<%Else%>class="lefttopborderblack"<%End If%>><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>Carats</b></font></TD>
					<TD bgcolor="white" align="center" class="topbuttomleftblack"><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>Description</b></font></TD>
                    <TD bgcolor="white" align="center" class="topbuttomleftblack"><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>P/C</b></font></TD>
					<TD bgcolor="white" align="center" <%If strDocType <> "total" then%>class="topbuttomleftrightblack"<%Else%>class="lefttoprightborderblack"<%End If%>><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>Total US$</b></font></TD>
				</TR>
				</TABLE>
				<div id='divTotalMode' style='display:block;'>
				<TABLE WIDTH="642" border=0 bordercolor=black cellpadding="0" cellspacing="0"  valign="top">
				<TR>
				    <TD class="leftborderblack"><img src="../Images/Blankt.gif" height="10" width="40" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="91" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="78" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="100" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="146" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="90" border="0"></TD>
				    <TD class="rightborderblack"><img src="../Images/Blankt.gif" height="10" width="97" border="0"></TD>
				</TR>

				<TR>
					<TD bgcolor="white" colspan=4 align="center" class="leftborderblack">&nbsp;</TD>
					<TD bgcolor="white" align="center"><font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;Polished Diamonds&nbsp;&nbsp;&nbsp;</b></font></TD> 
					<TD colspan=2 bgcolor="white" align="center" class="rightBorderBlack">&nbsp;</TD>
				</TR>
<%If Int(Request.Form("showTotal")) = Int(0) Then%>
		        <TR>
		            <TD colspan=100 class="leftrightborderblack">
<%				iRecCounter = 1
				iLogicSalePricetSum = 0
				
				Dim iCounterTR
					iCounterTR = 1				
				
				'Do While Not rs.EOF
				For iCounterInner = 0 To iMaxRowNumber
%>
				<%If Int(iCounterInner) >= Int(Request.Form("_txtRowCount")) Then%>
				    <div id='divInvoiceRow_<%=iCounterInner%>' style='display:none;'>
				<%Else%>
				    <div id='divInvoiceRow_<%=iCounterInner%>' style='display:block;'>
				<%End If%>
				<TABLE border=0 cellpadding=0 cellspacing=0 width="642">
				<TR>
					<TD width="40" bgcolor="white" align="left">
						<%If Int(strInnerLineCounter) > Int(0) Then%>
							<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;<b><%=iCounterInner + 1%>.&nbsp;</b></font>
						<%Else%>
							&nbsp;
						<%End If%>
					</TD>
					<TD width="91" bgcolor="<%=TD_BG_Color%>" align=center>
					    <%If Int(Request.Form("selShowRowCatalog")) = Int(1) Then%>
						<INPUT class="noBorder" readonly type="text" name="txtCatalog<%=iCounterTR%>_<%=iCounterInner%>" size="12" tabIndex=<%=IncTabIndex(iTabIndex)%> maxlength=10 value="<%=Request.Form("txtCatalog" & iCounterTR & "_" & iCounterInner)%>">  <!--OnKeyUP="setCatalogLST(this, '<%=iCounterTR%>_<%=iCounterInner%>', <%=strInnerLineCounter%>);" OnFocus="setCatalogLST(this, '<%=iCounterTR%>_<%=iCounterInner%>', <%=strInnerLineCounter%>);" onBlur="doDeleteID('<%=iCounterTR%>_<%=iCounterInner%>')">-->
						<%End If%>
						<INPUT type="hidden" name="txtRecID<%=iCounterTR%>_<%=iCounterInner%>" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=Request.Form("txtRecID" & iCounterTR & "_" & iCounterInner)%>">
					</TD>
					<TD width="78"	bgcolor="<%=TD_BG_Color%>" dir="rtl" align="center">
						<INPUT class="noBorder" readonly type="text" name="txtAmount<%=iCounterTR%>_<%=iCounterInner%>" size="7" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=Request.Form("txtAmount" & iCounterTR & "_" & iCounterInner)%>"  onBlur="JavaScript: calcTotalPrice('<%=iCounterTR%>_<%=iCounterInner%>')"> <!--readonly class=noBorder>-->
					</TD>
					<TD width="100"	bgcolor="<%=TD_BG_Color%>" dir=rtl align="center">					    
						<font face="Arial, Helvetica, sans-serif" size="2" color="white">
						<INPUT class="noBorder" readonly type="text" name="txtCarat<%=iCounterTR%>_<%=iCounterInner%>" size="11" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=Request.Form("txtCarat" & iCounterTR & "_" & iCounterInner)%>"  onBlur="JavaScript: calcTotalPrice('<%=iCounterTR%>_<%=iCounterInner%>')" >
						</font>									
					</TD>
					<TD width="146" bgcolor="white" align="center">
					    <INPUT class="noBorder" readonly type="text" name="txtDescription<%=iCounterTR%>_<%=iCounterInner%>" size="20" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=Request.Form("txtDescription" & iCounterTR & "_" & iCounterInner)%>">
					</TD>					
					<TD width="90"	bgcolor="<%=TD_BG_Color%>" dir="rtl" align="center">
						<INPUT class="noBorder" readonly type="text" name="txtPrice<%=iCounterTR%>_<%=iCounterInner%>" size="10" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=MyFormatNumber(Request.Form("txtPrice" & iCounterTR & "_" & iCounterInner))%>" onBlur="JavaScript: calcTotalPrice('<%=iCounterTR%>_<%=iCounterInner%>')">
					</TD>
					<TD width="97" bgcolor="<%=TD_BG_Color%>" dir="rtl" align="center">
						<INPUT class="noBorder" readonly tabIndex=<%=IncTabIndex(iTabIndex)%> type="text" name="txtSalesSum<%=iCounterTR%>_<%=iCounterInner%>" value="<%=MyFormatNumber(Request.Form("txtSalesSum" & iCounterTR & "_" & iCounterInner))%>" size="11" onBlur="JavaScript: calcCaratPrice('<%=iCounterTR%>_<%=iCounterInner%>');doSumFields();">
					</TD>
				</TR>
				</TABLE>
				</div>
<%					
					iLogicSalePricetSum = CDbl(iLogicSalePricetSum) + CDbl(CDbl(iTotal))
					iTotalSum = CDbl(iTotalSum) + CDbl(iTotal)


					iTotal = 0
					iRecCounter = iRecCounter + 1
				Next
%>
				    </TD>			
				</TR>
<%End If%>
                </TABLE>
                </div>
				<TABLE WIDTH="642" border=0 bordercolor=black cellpadding="0" cellspacing="0"  valign="top">
				<TR>
				    <TD class="leftborderblack"><img src="../Images/Blankt.gif" height="10" width="40" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="91" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="78" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="100" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="146" border="0"></TD>
				    <TD><img src="../Images/Blankt.gif" height="10" width="90" border="0"></TD>
				    <TD class="rightborderblack"><img src="../Images/Blankt.gif" height="10" width="97" border="0"></TD>
				</TR>

				<TR>
					<TD bgcolor="white" align="center" nowrap class="topleftblack">
					    <font face="Arial, Helvetica, sans-serif" size="2" color="black"><b>Total:</b></font>
					</TD>
					<TD bgcolor="white" align="center" class="topBorderBlack">
						&nbsp;
						<!--<font face="Arial, Helvetica, sans-serif" size="2" color="black">Total:</font>-->
					</TD>
					<TD bgcolor="white" align="center" dir="rtl" align="center" class="topBorderBlack">
						<INPUT class="noBorder" readonly type="text" name="textTotalQTYSum" dir="rtl" value="<%=(Request.Form("textTotalQTYSum"))%>" size="7" readonly class=noBorder>
					</TD>
					<TD bgcolor="white" align="center" class="topBorderBlack" nowrap>						
						<font face="Arial, Helvetica, sans-serif" size="2" color="black">
                        <INPUT class="noBorder" readonly type="text" name="textTotalCaratSum" dir="rtl" value="<%=MyFormatNumber(Request.Form("textTotalCaratSum"))%>" size="11" readonly class=noBorder>
						</font>									
					</TD>
					
					<TD bgcolor="white" align="right" class="topBorderBlack">
						<font face="Arial, Helvetica, sans-serif" size="2" color="black">&nbsp;</font>
						<img src="../Images/Blankt.gif" height=1 width="5" border=0>
					</TD>
					<TD bgcolor="white" align=right class="topBorderBlack">
						&nbsp;
					</TD>
					<TD nowrap bgcolor="white"  dir="rtl" <%If strDocType <> "total" then%>align=center class="toprightblack"<%Else%>align=center class="topbuttomleftrightblack"<%End If%>>
						<INPUT class="noBorder" readonly type="text" name="textTotalPriceSum" value="<%=MyFormatNumber(Request.Form("textTotalPriceSum"))%>" size="11" readonly class=noBorder>
					</TD>						
				</TR>
<%
    ShowShippingTotal = 1
    If Len(Trim(Request.Form("textShippingCost"))) = Int(0) Then
        ShowShippingTotal = 0
    Else
        If Int(Request.Form("textShippingCost")) = Int(0) Then
            ShowShippingTotal = 0
        End If
    End If
%>
<%If Int(ShowShippingTotal) = Int(1) Then%>
				<TR>
					<TD <%If strDocType <> "total" then%>colspan=5 <%Else%>colspan=4 <%End If%> bgcolor="white" <%If strDocType <> "total" then%>class="leftborderblack"<%Else%>class="topleftBlack"<%End If%>>&nbsp;</TD>
					<TD align="right" class="<%If strDocType <> "total" then%><%Else%>topBorderBlack<%End If%>">
                        <font face="Arial, Helvetica, sans-serif" size="2"><div id='divShippingLable' style='display:block;'><b><%=Request.Form("textShippingTitle1")%>:</b></div><div id='divShippingLable1' style='display:none;'>&nbsp;</div></font>
                    </TD>
					<TD valign="middle" <%If strDocType <> "total" then%>align="center" class="rightborderblack"<%Else%>align="center" class="toprightBlack"<%End If%>>
                        <div id='divShippingValue' style='display:block;'>
                        <INPUT class="noBorder" readonly type="text" name="textShippingCost" dir="rtl" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=MyFormatNumber(Request.Form("textShippingCost"))%>" size="11" onBlur=""><img src="../Images/Blankt.gif" height="1" width="3" border="0">
                        </div>
                        <div id='divShippingValue1' style='display:none;'>&nbsp;</div>
					</TD>
				</TR>


                <%If Len(Trim(Request.Form("textShippingTitle2"))) > Int(0) Or Len(Trim(Request.Form("textShippingCost2"))) > Int(0) Then%>
				<TR>
					<TD <%If strDocType <> "total" then%>colspan=5 <%Else%>colspan=4 <%End If%> bgcolor="white" <%If strDocType <> "total" then%>class="leftborderblack"<%Else%>class="topleftBlack"<%End If%>>&nbsp;</TD>
					<TD align="right" class="<%If strDocType <> "total" then%><%Else%>topBorderBlack<%End If%>">
                        <font face="Arial, Helvetica, sans-serif" size="2"><div id='div1' style='display:block;'><b><%=Request.Form("textShippingTitle2")%>:</b></div><div id='div2' style='display:none;'>&nbsp;</div></font>
                    </TD>
					<TD valign="middle" <%If strDocType <> "total" then%>align="center" class="rightborderblack"<%Else%>align="center" class="toprightBlack"<%End If%>>
                        <div id='div3' style='display:block;'>
                        <INPUT class="noBorder" readonly type="text" name="textShippingCost2" dir="rtl" tabIndex=<%=IncTabIndex(iTabIndex)%> value="<%=MyFormatNumber(Request.Form("textShippingCost2"))%>" size="11" onBlur=""><img src="../Images/Blankt.gif" height="1" width="3" border="0">
                        </div>
                        <div id='div4' style='display:none;'>&nbsp;</div>
					</TD>
				</TR>
                <%End If%>
				<TR>											
					<TD colspan=5 bgcolor="white" <%If strDocType <> "total" then%>class="buttomleftBlack"<%Else%>class="topleftBlack"<%End If%>>&nbsp;</TD>
					<TD align="right" <%If strDocType <> "total" then%>class="lightButtomTable"<%Else%>class="topBorderBlack"<%End If%>>
						<font face="Arial, Helvetica, sans-serif" size="2" color="black"><div id='divSecounderyTotalLable' style='display:block;'><b>Total:</b></div></font><div id='divSecounderyTotalLable1' style='display:none;'>&nbsp;</div>
					</TD>
					<TD valign="middle" align="center" <%If strDocType <> "total" then%>class="buttomrightBlack"<%Else%>class="toprightBlack"<%End If%>>
                        <div id='divSecounderyTotalValue' style='display:block;'><INPUT class="noBorder" readonly type="text" name="textTotalPriceSumAfterShippingCost" dir="rtl" value="<%=MyFormatNumber(Request.Form("textTotalPriceSumAfterShippingCost"))%>" size="11" readonly class=noBorder><img src="../Images/Blankt.gif" height="1" width="3" border="0"></div><div id='divSecounderyTotalValue1' style='display:none;'>&nbsp;</div>
					</TD>
				</TR>
<%Else%>
				<TR>
					<TD class=topBorderBlack colspan=7>&nbsp;</TD>
				</TR>
<%End If%>					
				</TABLE>
				</TD>
			</TR>						
			</TABLE>
<br>
<br>



<div <%=strAlignAndDir%> >
	<font face="Arial, Helvetica, sans-serif" size="2" color="black"><%=strText%></font><br><br>
</div>


<%If Int(iFromSRCompany) = Int(5) Then%>
<TABLE <%=strAlignAndDirBottom%>>
<TR>
	<TD ALIGN=CENTER>
		<font face="Arial, Helvetica, sans-serif" size="2" color="black">
			<b>TA ROZ&nbsp;&nbsp;&nbsp;LTD</b>
		</font>
	</TD>
</TR>
<TR>
	<TD ALIGN=CENTER>
		<font face="Arial, Helvetica, sans-serif" size="2" color="black">
		    <b>514702737</b>
		</font>
	</TD>
</TR>
</TABLE>
<%Else%>
<TABLE <%=strAlignAndDirBottom%>>
<TR>
	<TD ALIGN=CENTER>
	    <font face=arial  size=2>
        <%
            arrCompSrcDesc = Split(strCompSrcDesc, "<br>")
            
            If Len(Trim(Ubound(arrCompSrcDesc))) > Int(0) Then
                If Int(Ubound(arrCompSrcDesc)) >= Int(1) Then
                    Response.Write arrCompSrcDesc(0)
                End If
            End If
        %>
        </font>
	</TD>
</TR>
</TABLE>
<%End If%>

<!--div id=divPrintBtn style="display: block;">
<br>
<TABLE align=left border=0>
<TR>
	<TD ALIGN=right>
		<INPUT type="button" value="Back" name="butBack" onClick="document.theBackForm.submit();" tabIndex=<%=IncTabIndex(iTabIndex)%>  style="width: 70px;">
		<INPUT type="button" value="Print Invoice" name="butPrt" onClick="ShowPrintDocument();" tabIndex=<%=IncTabIndex(iTabIndex)%>  style="width: 90px;">
		<INPUT type="button" value="Print Total" name="butPrt" onClick="ShowPrintDocumentTotal();" tabIndex=<%=IncTabIndex(iTabIndex)%>  style="width: 80px;">
		<INPUT type="button" value="Print Reshomon" name="butReshomon" onClick="ShowReshomonDocument();" tabIndex=<%=IncTabIndex(iTabIndex)%>  style="width: 105px;">
		<INPUT type="button" value="Print Shipper's Letter" name="butShipper" onClick="ShowShipperDocument();" tabIndex=<%=IncTabIndex(iTabIndex)%> style="width: 150px;">
	</TD>
</TR>	
</TABLE>	
</div-->

</FORM>


	</TD>
</TR>
</TABLE>


<div id='divDoPrint' style='display:block;' align=center>
    <input id="Button1" type="button" value="Print" onclick="doPrint();">
</div>

<%
	closeDB
%>

<script language="javaScript">
    function doPrint(){
        document.all.divDoPrint.style.display='none';
        window.print();
        setTimeout("window.close();", 500);
    }
</script>
<!--#include file="../Inc/Footer.asp"-->
