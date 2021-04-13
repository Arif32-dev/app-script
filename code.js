function atEdit(e) {
    let row_data = row_values(e);
    if (!row_data) return;

    let response = get_bol_data(row_data);
    if (!response) return;
    let response_data = JSON.parse(response);

    switch (response_data.type) {
        case 'success':
            set_value_on_success(response_data.response, response_data.bol_ID, e);
            break;
        case 'error':
            show_warning(e, response_data.response);
            break;
        default:
            return false;
    }

    Logger.log(response_data);

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

    if (!shipping_obj.name) return show_warning(e, warning_msg);

    if (!shipping_obj.company_name) return show_warning(e, warning_msg);

    if (!shipping_obj.address) return show_warning(e, warning_msg);

    if (!shipping_obj.city) return show_warning(e, warning_msg);

    if (!shipping_obj.state) return show_warning(e, warning_msg);

    if (!shipping_obj.zip) return show_warning(e, warning_msg);

    if (!shipping_obj.country) return show_warning(e, warning_msg);

    if (!shipping_obj.phone) return show_warning(e, warning_msg);


    return shipping_obj;
}

function row_values(e) {
    let { spreadsheet, current_row, current_sheet, current_column, active_sheet } = get_event_data(e);

    if (current_sheet == 'Orders') {

        if (current_column > 11) return false;

        let is_pdf_generated = active_sheet.getRange(`L${current_row}`).getValues()[0][0];

        if (is_pdf_generated.match(/WC\d+/g)) return false;

        clear_warning(e);

        let values = spreadsheet.getRange(`E${current_row}:K${current_row}`).getValues();

        let order_date = values[0][0] ? values[0][0] : null;
        let order_number = values[0][3] ? values[0][3] : null;
        let shipping_data = split_shipping_data(values[0][4] ? values[0][4] : null, e) ? split_shipping_data(values[0][4] ? values[0][4] : null, e) : null;
        let source = values[0][5] ? values[0][5] : null;
        let carrier = values[0][6] ? values[0][6] : null;

        if (!order_date) return show_warning(e, "Order date is empty");
        if (!order_number) return show_warning(e, "Order number is empty");
        if (!shipping_data) return false;
        if (source != 'USCD') return show_warning(e, "Source need's to be USCD");
        if (carrier != 'R&L') return show_warning(e, "Carrier need's to be R&L");

        return {
            order_date,
            order_number,
            shipping_data,
            source,
            carrier
        }
    } else {
        return false;
    }
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
    if (!data.order_date ||
        !data.order_number ||
        !data.shipping_data ||
        !data.source ||
        !data.carrier) return false;

    let url = "https://hoodsly.com/wp-json/spreadsheet/v1/bol-generation";
    var options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(data)
    };
    let response = UrlFetchApp.fetch(url, options);
    return response;
};

function show_warning(e, msg) {
    let { active_sheet, current_row } = get_event_data(e)
    let range = active_sheet.getRange(`L${current_row}`);
    range.setValue(msg)
    range.setFontColor("red");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    return false;
}

function clear_warning(e) {
    let { active_sheet, current_row } = get_event_data(e)
    let range = active_sheet.getRange(`L${current_row}`);
    range.clear();
}

function set_value_on_success(pdf_link, bol_ID, e) {
    let { active_sheet, current_row } = get_event_data(e)
    let range = active_sheet.getRange(`L${current_row}`);
    let value = `=HYPERLINK("${pdf_link}", "WC${bol_ID}")`;
    range.setValue(value);
    range.setFontColor("cornflower blue");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
}
