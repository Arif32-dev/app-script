function atEdit(e) {
    let row_data = row_values(e);
    if (!row_data) return;

    let response = get_bol_data(row_data);

    Logger.log(response);

    if (!response) return;
    let response_data = JSON.parse(response);

    switch (response_data.type) {
        case 'success':
            if (row_data.url != undefined) {
                set_value_on_success(response_data.pdf_link, response_data.shipping_link, 'bolFetching', e);
            } else {
                set_value_on_success(response_data.response.bol_link, response_data.response.shipping_link, response_data.bol_ID, e);
            }
            break;
        case 'error':
            show_warning(e, "BOL could'nt be generated. Check if the order and BOL pdf exists in website");
            break;
        default:
            return false;
    }

}

function onOpen(e) {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .createMenu('Shipping Address')
        .addItem('Enter Info', 'initDataCollection')
        .addToUi();
}

function initDataCollection() {
    let promptResponse = openPropmt("Buyer's Name");
    let resStatus = promptResponse.resStatus;
    let resText = promptResponse.resText;
    if (resStatus == 'CLOSE' ||
        resStatus == 'CANCEL' ||
        resStatus != 'OK' ||
        resText == '') return false;

    let shippingData = {};

    let followupRequiredFields = {
        'companyName': 'Company Name',
        'streetAddress': 'Street Address',
        'city': "City",
        'state': "State",
        'zipCode': 'Zip Code',
        'country': "Country",
        'phoneNumber': 'Phone Number',
    }
    if (resText) {
        shippingData.buyerName = resText;

        for (const property in followupRequiredFields) {
            let promptResponse = openPropmt(followupRequiredFields[property]);
            let resStatus = promptResponse.resStatus;
            let resText = promptResponse.resText;

            if (resStatus == 'CLOSE' ||
                resStatus == 'CANCEL' ||
                resStatus != 'OK' ||
                resText == '') break;

            if (resText) {
                shippingData[property] = resText;
            } else {
                break;
            }
        }
    }

    if (shippingData.buyerName &&
        shippingData.companyName &&
        shippingData.streetAddress &&
        shippingData.city &&
        shippingData.state &&
        shippingData.zipCode &&
        shippingData.country &&
        shippingData.phoneNumber) {

        setShippingData(shippingData);

    } else {
        setShippingData(false);
    }

}

function setShippingData(data) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let activeRange = sheet.getActiveRange();
    let activeRow = activeRange.getRow();
    let range = SpreadsheetApp.getActiveSpreadsheet().getRange(`I${activeRow}`);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");

    if (data) {
        let shippingData = `${data.buyerName}\n${data.companyName}\n${data.streetAddress}\n${data.city}, ${data.state}, ${data.zipCode}  ${data.country}\nT: ${data.phoneNumber}`;
        range.setValue(shippingData);

        let clearRange = SpreadsheetApp.getActiveSpreadsheet().getRange(`L${activeRow}`);
        clearRange.clear();

    } else {
        range.setValue('Info is not completed yet. Please try again')
    }
}

function openPropmt(requiredText) {
    let ui = SpreadsheetApp.getUi();
    let propmt = ui.prompt(`${requiredText}?`, ui.ButtonSet.OK_CANCEL);
    let resStatus = propmt.getSelectedButton();
    let resText = propmt.getResponseText();
    return {
        resStatus,
        resText
    };
}

function split_shipping_data(shipping_data, e) {

    let { spreadsheet, current_row } = get_event_data(e);

    let values = spreadsheet.getRange(`E${current_row}:K${current_row}`).getValues();
    let source = values[0][5] ? values[0][5] : null;
    let carrier = values[0][6] ? values[0][6] : null;

    if (source == 'USCD' && carrier == 'R&L') {

        if (!shipping_data) return show_warning(e, "Shipping data is empty");
        let shipping_array = shipping_data.split(/\n/);

        if (shipping_array.length <= 4) return show_warning(e, 'Field missing');

        let location = shipping_array[3] ? shipping_array[3].split(',') : null;
        if (!location) return false;
        let shipping_obj = {
            'name': shipping_array[0] ? shipping_array[0].trim() : null,
            'company_name': shipping_array[1] ? shipping_array[1].trim() : null,
            'address': shipping_array[2] ? shipping_array[2].trim() : null,
            'city': location[0] ? location[0].trim() : null,
            'state': location[1] ? location[1].trim() : null,
            'zip': location[2] ? location[2].split(" ") ? location[2].trim().split(" ")[0].trim() : null : null,
            'country': "USA",
            'phone': shipping_array[4] ? shipping_array[4].trim().replace("T:", "").trim() : null,
        }

        let warning_msg = 'Field missing';

        if (!shipping_obj.name ||
            !shipping_obj.company_name ||
            !shipping_obj.address ||
            !shipping_obj.city ||
            !shipping_obj.state ||
            !shipping_obj.zip ||
            !shipping_obj.country ||
            !shipping_obj.phone) return show_warning(e, warning_msg);

        return shipping_obj;

    }
}

function row_values(e) {
    let { spreadsheet, current_row, current_sheet, current_column, active_sheet } = get_event_data(e);

    if (current_sheet == 'Orders') {

        if (current_column > 11) return false;

        let hasValue = active_sheet.getRange(`L${current_row}`).getValues()[0][0];

        // if already value exist, do nothing
        if (hasValue) return;

        if (current_column == 10 || current_column == 11) {

            let values = spreadsheet.getRange(`E${current_row}:K${current_row}`).getValues();

            let order_date = values[0][0] ? values[0][0] : null;
            let order_number = values[0][3] ? values[0][3] : null;
            let source = values[0][5] ? values[0][5] : null;
            let carrier = values[0][6] ? values[0][6] : null;

            if (source == 'Hoodsly' || source == 'WWH') {
                let getUrl = hoodslyOrWWH(order_number, source);
                return {
                    url: getUrl
                }
            }

            let shipping_data = split_shipping_data(values[0][4] ? values[0][4] : null, e) ? split_shipping_data(values[0][4] ? values[0][4] : null, e) : null;

            if (source == 'USCD' && carrier == 'R&L') {

                if (!order_date) return show_warning(e, "Order date is empty");
                if (!order_number) return show_warning(e, "Order number is empty");
                if (!shipping_data) return false;

                return {
                    order_date,
                    order_number,
                    shipping_data,
                    source,
                    carrier
                }
            }
        }

    } else {
        return false;
    }
}

function hoodslyOrWWH(order_number, source) {
    let url = null;
    if (source == 'Hoodsly') {
        url = "https://hoodsly.com/wp-json/generate_bol/v1/bol-fetching?orderID=" + order_number + "";
        return url;
    } else if (source == 'WWH') {
        url = "https://wholesalewoodhoods.com/wp-json/generate_bol/v1/bol-fetching?orderID=" + order_number + "";
        return url;
    }
    return url;
}

function get_event_data(e) {
    return {
        spreadsheet: e.source,
        active_sheet: e.source.getActiveSheet(),
        range_obj: e.range,
        current_row: e.range.getRow(),
        current_column: e.range.getColumn(),
        current_sheet: e.source.getActiveSheet().getName()
    };
}

function get_bol_data(data) {
    if (data.url != undefined) {

        if (data.url == "") return false;

        let response = UrlFetchApp.fetch(data.url);
        return response;
    }

    if (!data.order_date ||
        !data.order_number ||
        !data.shipping_data ||
        !data.source ||
        !data.carrier) return false;

    let url = "https://hoodsly.com/wp-json/spreadsheet/v1/bol-generation";
    let options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(data)
    };
    let response = UrlFetchApp.fetch(url, options);
    return response;
};

function show_warning(e, msg) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(msg);
}

function clear_warning(e) {
    let { active_sheet, current_row } = get_event_data(e)
    let range = active_sheet.getRange(`L${current_row}`);
    range.clear();
}

function set_value_on_success(pdf_link, shipping_link, bol_ID, e) {

    let { active_sheet, current_row } = get_event_data(e)
    let range = active_sheet.getRange(`L${current_row}:M${current_row}`);

    let pdfLink = '';
    let shippingLink = '';

    if (bol_ID == 'bolFetching') {
        pdfLink = `=HYPERLINK("${pdf_link}", "View BOL")`;
        shippingLink = `=HYPERLINK("${shipping_link}", "Shipping Label")`;
    } else {
        pdfLink = `=HYPERLINK("${pdf_link}", "WC${bol_ID}")`;
        shippingLink = `=HYPERLINK("${shipping_link}", "Shipping Label")`;
    }

    range.setValues([
        [pdfLink, shippingLink]
    ]);
    range.setFontColor("cornflower blue");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
}