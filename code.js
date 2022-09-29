function onClickFulfillInShopify() {
    var ui = SpreadsheetApp.getUi();

    var response = ui.alert('Mark Fulfilled In Shopify', 'Should we mark the Packed, Delivered, Shipped, or Picked Up items fulfilled in Shopify?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {

        var activeSheet = SpreadsheetApp.getActiveSheet();
        var activeSheetName = activeSheet.getName();
        var lastRowOfDataInActiveSheet = activeSheet.getLastRow();
        var lastColumnOfDataInActiveSheet = activeSheet.getLastColumn();
        var firstColumnName = "Order #";

        // get header row index and number
        var headerRowNumber = getHeaderRowNumber(activeSheet, firstColumnName);
        var firstRowOfData = headerRowNumber + 1;
        var headerRowArray = getHeaderRowArrayFromHeaderRowNumber(activeSheet, headerRowNumber);

        // Get column numbers
        
            // Fulfillment Location column number
            var fulfillmentLocationColumnName = "Fulfillment Location"
            var fulfillmentLocationColumnNumber = getColumnNumberBasedOnHeaderRowArrayAndColumnName(headerRowArray, fulfillmentLocationColumnName);
            var fulfillmentLocationColumnIndex = fulfillmentLocationColumnNumber - 1;
    
            // Fulfilled in Shopify column number
            var fulfilledInShopifyColumnName = "Fulfilled in Shopify"
            var fulfilledInShopifyColumnNumber = getColumnNumberBasedOnHeaderRowArrayAndColumnName(headerRowArray, fulfilledInShopifyColumnName);
            var fulfilledInShopifyColumnIndex = fulfilledInShopifyColumnNumber - 1;

        var nextRowToPasteToInDeliveryCompletedSheet = lastRowOfDataInActiveSheet + 1;
        var fulfillmentLocationId;
        var fulfillOrderInShopifyResponse;
        var fulfilledList = [];
        var notFulfilledList = [];
        var missingOrderIdList = [];
        var missingLocationList = [];
        var transferList = [];
        var alreadyFulfilledList = [];

        // Loop through each row and mark fulfilled in shopify if appropriate
        for (var i=7; i<=lastRowOfDataInActiveSheet; i++) {
            Logger.log("Up to row: " + i);

            // If Delivered then mark delivered in Shopify and add flag
            var rangeForRowOfData = activeSheet.getRange(i, 1, 1, lastColumnOfDataInActiveSheet);
            var valuesForRowOfData = rangeForRowOfData.getValues();
            var fulfillmentStatusRange = activeSheet.getRange(i, fulfilledInShopifyColumnNumber, 1, 1);
            var fulfillmentStatusValue = activeSheet.getRange(i, fulfilledInShopifyColumnNumber, 1, 1).getValue();
            var orderName = valuesForRowOfData[0][0];
            var fulfillmentLocation = valuesForRowOfData[0][fulfillmentLocationColumnIndex];
            var orderId = valuesForRowOfData[0][44];
            
            // Depending on sheet name, set data
            // If sheet name is Delivery Plan, then Status is in column index 6, otherwise column index 3
            if ((activeSheetName == "Delivery Plan") || (activeSheetName == "Delivery Plan - Hunt")) {
                var status = valuesForRowOfData[0][6];
            } else {
                var status = valuesForRowOfData[0][3];
            };

            // Mark fulfilled if Delivered, Packed, or Picked Up
            if (
              ((status == "5. Delivered") && (activeSheetName == "Delivery Plan")) || 
              ((status == "5. Delivered") && (activeSheetName == "Delivery Plan - Hunt")) || 
              ((status == "2. Packed") && (activeSheetName == "Delivery Plan")) ||
              ((status == "2. Packed") && (activeSheetName == "Delivery Plan - Hunt")) ||
              ((status == "2. Packed") && (activeSheetName == "Pickup Plan - Gowanus")) || 
              ((status == "2. Packed") && (activeSheetName == "Pickup Plan - Greenpoint")) || 
              ((status == "2. Packed") && (activeSheetName == "Pickup Plan - Huntington")) || 
              (status == "3. Picked Up") || (status == "4. Picked Up")
            ) {

                if (fulfillmentLocation == 'Transfer') {
                    transferList.push(orderName);
                    fulfillmentStatusRange.setValue('transfer');
                } else if ((orderId == '') || (orderId == null)) {
                    missingOrderIdList.push(orderName);
                } else if ((fulfillmentLocation == '') || (fulfillmentLocation == null)) {
                    missingLocationList.push(orderName);
                } else if (fulfillmentStatusValue == "success") {
                    alreadyFulfilledList.push(orderName);
                } else {
                    
                    if (fulfillmentLocation == "Gowanus") { 
                        fulfillmentLocationId = INSERT_ID;
                    } else if (fulfillmentLocation == "Greenpoint") { 
                        fulfillmentLocationId = INSERT_ID;
                    } else if (fulfillmentLocation == "Huntington") { 
                        fulfillmentLocationId = INSERT_ID;
                    } else if (fulfillmentLocation == "225_Third") { 
                        fulfillmentLocationId = INSERT_ID;
                    };

                    fulfillOrderInShopifyResponse = fulfillOrderInShopifyWithLocation(orderId, fulfillmentLocationId);
                    fulfillmentStatusRange.setValue(fulfillOrderInShopifyResponse);
                    
                    if (fulfillOrderInShopifyResponse == "success") {
                        fulfilledList.push(orderName);
                    } else {
                        notFulfilledList.push(orderName);
                    };
                };   
            } else {
                notFulfilledList.push(orderName);
            }; // Closes if status
       }; // Closes for loop
       
      ui.alert("Fulfilled " + fulfilledList.length + " in Shopify: " + fulfilledList + 
      "\r\n\r\nDid NOT fulfill " + alreadyFulfilledList.length + " orders in Shopify because they were already fulfilled: " + alreadyFulfilledList +
      "\r\n\r\nDid NOT fulfill " + notFulfilledList.length + " orders in Shopify for some other reason (maybe wrong status): " + notFulfilledList +
      "\r\n\r\nDid NOT fulfill " + missingOrderIdList.length + " orders in Shopify because they are missing Order ID: " + missingOrderIdList +
      "\r\n\r\nDid NOT fulfill " + missingLocationList.length + " orders in Shopify because they are missing Fulfillment Location: " + missingLocationList +
      "\r\n\r\nDid NOT fulfill " + transferList.length + " orders in Shopify because they're transfers not orders: " + transferList);
   };
};

function simpleWorkingExample() {
    var url = "https://threesbrewing.myshopify.com/admin/api/2020-04/orders.json";
    var response = UrlFetchApp.fetch(url, {
            "method": "get",
            'contentType': 'application/json',
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode("INSERT_KEY")
                }
            });
    Logger.log(response.getContentText());
    var responseContentText = response.getContentText();
    var responseContentHash = JSON.parse(responseContentText);
};

function getFulfillmentForOrder() {
    var orderId = "2270441013310";
    var url = "https://threesbrewing.myshopify.com/admin/api/2020-04/orders/" + orderId + "/fulfillment_orders.json"
    Logger.log("url: " + url);
    var response = UrlFetchApp.fetch(url, {
            "method": "get",
            'contentType': 'application/json',
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode("INSERT_KEY")
                }
            }
    );

    var responseContentText = response.getContentText();
    Logger.log("responseContentText: " + responseContentText);
    
    var responseContentHash = JSON.parse(responseContentText);
    Logger.log("responseContentHash: " + responseContentHash);  
};

function getLocationsList() {
// https://shopify.dev/docs/admin-api/rest/reference/inventory/location#index-2020-04

    var url = "https://threesbrewing.myshopify.com/admin/api/2020-04/locations.json"
    Logger.log("url: " + url);
    var response = UrlFetchApp.fetch(url, {
            "method": "get",
            'contentType': 'application/json',
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode("INSERT_KEY")
                }
            }
    );

    var responseContentText = response.getContentText();
    Logger.log("responseContentText: " + responseContentText);
    
    var responseContentHash = JSON.parse(responseContentText);
    Logger.log("responseContentHash: " + responseContentHash);
    Logger.log("locations: " + responseContentHash["locations"]);
    Logger.log("locations: " + responseContentHash["locations"].length);

};
