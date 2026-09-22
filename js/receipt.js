document.addEventListener("DOMContentLoaded", function () {

    // =========================================================
    // FORM
    // =========================================================

    const receiptForm = document.getElementById("ReceiptForm");

    if (!receiptForm) {
        console.error("ReceiptForm not found.");
        return;
    }


    // =========================================================
    // HELPER FUNCTIONS
    // =========================================================

    function getValue(id) {
        return document.getElementById(id).value.trim();
    }

    function getNumber(id) {
        return Number(document.getElementById(id).value) || 0;
    }


    // =========================================================
    // SERVICE & FINANCIAL INPUTS
    // =========================================================

    const quantity1 = document.getElementById("SI-Qty");
    const unitPrice1 = document.getElementById("SI-UnitPrice");

    const serviceDesc2 = document.getElementById("SI-Desc2");
    const quantity2 = document.getElementById("SI-Qty2");
    const unitPrice2 = document.getElementById("SI-UnitPrice2");

    const discount = document.getElementById("FI-Discount");
    const adjustment = document.getElementById("FI-AddAdj");
    const adjustmentDescription = document.getElementById("FI-AddAdjDesc");


    // =========================================================
    // CONDITIONAL REQUIRED FIELDS
    // =========================================================

    function updateServiceDesc2Required() {
        const hasDescription = serviceDesc2.value.trim() !== "";
        quantity2.required = hasDescription;
        unitPrice2.required = hasDescription;
    }

    serviceDesc2.addEventListener("input", updateServiceDesc2Required);

    function updateAdjustmentRequired() {
        const hasAdjustment = adjustment.value.trim() !== "";
        adjustmentDescription.required = hasAdjustment;
    }

    adjustment.addEventListener("input", updateAdjustmentRequired);


    // =========================================================
    // RECEIPT CALCULATIONS
    // =========================================================

    function calculateReceipt() {

        const qty1 = Number(quantity1.value) || 0;
        const price1 = Number(unitPrice1.value) || 0;
        const qty2 = Number(quantity2.value) || 0;
        const price2 = Number(unitPrice2.value) || 0;
        const discountPercent = Number(discount.value) || 0;
        const adjustmentAmount = Number(adjustment.value) || 0;
        const lineTotal1 = qty1 * price1;
        const lineTotal2 = qty2 * price2;
        const subtotal = lineTotal1 + lineTotal2;
        const discountAmount = subtotal * (discountPercent / 100);
        const afterDiscount = subtotal - discountAmount;
        const total = afterDiscount + adjustmentAmount;


        // Display results
        document.getElementById("SubtotalDisplay").textContent =
            subtotal.toFixed(2);

        document.getElementById("FI-DiscountAmount").textContent =
            discountAmount.toFixed(2);

        document.getElementById("AfterDiscountDisplay").textContent =
            afterDiscount.toFixed(2);

        document.getElementById("TotalDisplay").textContent =
            total.toFixed(2);
    }


    [
        quantity1,
        unitPrice1,
        quantity2,
        unitPrice2,
        discount,
        adjustment
    ].forEach(function (input) {
        input.addEventListener("input", calculateReceipt);
    });


    // Calculate once when the page loads
    calculateReceipt();


    // =========================================================
    // RECEIPT NUMBER
    // =========================================================

    const initials = document.getElementById("IN-Initials");
    const year = document.getElementById("IN-Year");
    const seqNum = document.getElementById("IN-SeqNum");

    const receiptNumberPreview =
        document.getElementById("ReceiptNumberPreview");


    function buildReceiptNumber() {
        return (
            initials.value.trim().toUpperCase() +
            "-R" +
            year.value.trim() +
            seqNum.value.trim()
        );
    }


    function updateReceiptNumber() {
        receiptNumberPreview.textContent =
            buildReceiptNumber();
    }


    initials.addEventListener("input", function () {
        initials.value = initials.value.toUpperCase();
        updateReceiptNumber();
    });

    year.addEventListener("input", updateReceiptNumber);
    seqNum.addEventListener("input", updateReceiptNumber);


    // =========================================================
    // SUBMIT FORM
    // =========================================================

    receiptForm.addEventListener(
        "submit",
        async function (event) {

            event.preventDefault();

            try {

                // =====================================================
                // 1. GET FORM VALUES
                // =====================================================

                const formData = {

                    // Company Info
                    CIrep: getValue("CI-Rep"),
                    CIcompany: getValue("CI-Name"),
                    CIstreet: getValue("CI-Street"),
                    CIcity: getValue("CI-City"),
                    CIstate: getValue("CI-State"),
                    CIzip: getValue("CI-Zip"),
                    CIemail: getValue("CI-Email"),
                    CIphone: getValue("CI-Phone"),

                    // Pick Up Info
                    PUcompany: getValue("PU-Company"),
                    PUstreet: getValue("PU-Street"),
                    PUcity: getValue("PU-City"),
                    PUstate: getValue("PU-State"),
                    PUzip: getValue("PU-Zip"),
                    PUdate: getValue("PU-Date"),
                    PUtime: getValue("PU-Time"),

                    // Delivery Info
                    Delcompany: getValue("Del-Company"),
                    Delstreet: getValue("Del-Street"),
                    Delcity: getValue("Del-City"),
                    Delstate: getValue("Del-State"),
                    Delzip: getValue("Del-Zip"),
                    Deldate: getValue("Del-Date"),
                    Deltime: getValue("Del-Time"),
                    DelPoc: getValue("Del-POC"),

                    // Service Info
                    SIservices: getValue("SI-Services"),
                    SIdesc: getValue("SI-Desc"),
                    SIqty: getNumber("SI-Qty"),
                    SIunitPrice: getNumber("SI-UnitPrice"),

                    SIdesc2: getValue("SI-Desc2"),
                    SIqty2: getNumber("SI-Qty2"),
                    SIunitPrice2: getNumber("SI-UnitPrice2"),

                    // Financial Info
                    FImethod: getValue("FI-Method"),

                    // Excel percentage value
                    FIdiscount: getNumber("FI-Discount") / 100,

                    FIadjustments: getNumber("FI-AddAdj"),

                    // Receipt Number
                    INfinal: buildReceiptNumber(),

                    // Additional Info
                    Notes:
                        getValue("FI-AddAdjDesc") +
                        "\n" +
                        getValue("Notes"),

                    SubDate: getValue("SubDate")
                };


                // =====================================================
                // 2. LOAD EXCEL TEMPLATE
                // =====================================================

                const arrayBuffer =
                    await downloadDropboxTemplate(
                        "ReceiptTemplate.xlsx"
                    );
                const zip =
                    await JSZip.loadAsync(arrayBuffer);


                // =====================================================
                // 3. LOAD WORKBOOK XML
                // =====================================================

                let workbookXML =
                    await zip
                        .file("xl/workbook.xml")
                        .async("string");


                // =====================================================
                // 4. LOAD WORKSHEET XML
                // =====================================================

                const worksheetFile =
                    zip.file("xl/worksheets/sheet1.xml");

                if (!worksheetFile) {
                    throw new Error(
                        "xl/worksheets/sheet1.xml was not found."
                    );
                }

                let worksheetXML =
                    await worksheetFile.async("string");


                // =====================================================
                // 5. LOAD SHARED STRINGS
                // =====================================================

                const sharedStringsFile =
                    zip.file("xl/sharedStrings.xml");

                if (!sharedStringsFile) {
                    throw new Error(
                        "xl/sharedStrings.xml was not found."
                    );
                }

                let sharedStringsXML =
                    await sharedStringsFile.async("string");


                // =====================================================
                // 6. GET SHARED STRING COUNT
                // =====================================================

                const uniqueCountMatch =
                    sharedStringsXML.match(
                        /uniqueCount="(\d+)"/
                    );

                if (!uniqueCountMatch) {
                    throw new Error(
                        "Could not find uniqueCount in sharedStrings.xml."
                    );
                }

                let nextSharedStringIndex =
                    parseInt(uniqueCountMatch[1], 10);


                // =====================================================
                // 7. ADD SHARED STRING
                // =====================================================

                function addSharedString(value) {

                    const index =
                        nextSharedStringIndex;

                    const escapedValue =
                        escapeXML(value);

                    const newSharedString =
                        `<si><t>${escapedValue}</t></si>`;

                    sharedStringsXML =
                        sharedStringsXML.replace(
                            "</sst>",
                            newSharedString + "</sst>"
                        );

                    nextSharedStringIndex++;

                    return index;
                }


                // =====================================================
                // 8. ADD FORM VALUES TO SHARED STRINGS
                // =====================================================

                const indexes = {

                    // Company
                    CIrep: addSharedString(formData.CIrep),
                    CIcompany: addSharedString(formData.CIcompany),
                    CIstreet: addSharedString(formData.CIstreet),
                    CIcity: addSharedString(formData.CIcity),
                    CIstate: addSharedString(formData.CIstate),
                    CIzip: addSharedString(formData.CIzip),
                    CIemail: addSharedString(formData.CIemail),
                    CIphone: addSharedString(formData.CIphone),

                    // Pick Up
                    PUcompany: addSharedString(formData.PUcompany),
                    PUstreet: addSharedString(formData.PUstreet),
                    PUcity: addSharedString(formData.PUcity),
                    PUstate: addSharedString(formData.PUstate),
                    PUzip: addSharedString(formData.PUzip),
                    PUdate: addSharedString(formData.PUdate),
                    PUtime: addSharedString(formData.PUtime),

                    // Delivery
                    Delcompany: addSharedString(formData.Delcompany),
                    Delstreet: addSharedString(formData.Delstreet),
                    Delcity: addSharedString(formData.Delcity),
                    Delstate: addSharedString(formData.Delstate),
                    Delzip: addSharedString(formData.Delzip),
                    Deldate: addSharedString(formData.Deldate),
                    Deltime: addSharedString(formData.Deltime),
                    DelPoc: addSharedString(formData.DelPoc),

                    // Services
                    SIservices: addSharedString(formData.SIservices),
                    SIdesc: addSharedString(formData.SIdesc),
                    SIdesc2: addSharedString(formData.SIdesc2),

                    // Financial
                    FImethod: addSharedString(formData.FImethod),

                    // Receipt Number
                    INfinal: addSharedString(formData.INfinal),

                    // Additional Info
                    Notes: addSharedString(formData.Notes),
                    SubDate: addSharedString(formData.SubDate)
                };


                // =====================================================
                // 9. UPDATE UNIQUE COUNT
                // =====================================================

                sharedStringsXML =
                    sharedStringsXML.replace(
                        /uniqueCount="\d+"/,
                        `uniqueCount="${nextSharedStringIndex}"`
                    );


                // =====================================================
                // 10. REPLACE EXISTING SHARED-STRING CELL
                // =====================================================

                function replaceExistingCellValue(
                    cellAddress,
                    newIndex
                ) {

                    const cellRegex = new RegExp(
                        `<c\\s+[^>]*r="${cellAddress}"[^>]*>[\\s\\S]*?<\\/c>`,
                        "i"
                    );

                    const match = worksheetXML.match(cellRegex);

                    if (!match) {
                        throw new Error(
                            `Could not find cell ${cellAddress}.`
                        );
                    }

                    const originalCell = match[0];

                    if (!/<v>[\s\S]*?<\/v>/i.test(originalCell)) {
                        throw new Error(
                            `Cell ${cellAddress} does not contain a value.`
                        );
                    }

                    const updatedCell =
                        originalCell.replace(
                            /<v>[\s\S]*?<\/v>/i,
                            `<v>${newIndex}</v>`
                        );

                    worksheetXML =
                        worksheetXML.replace(
                            originalCell,
                            updatedCell
                        );
                }


                // =====================================================
                // 11. REPLACE NUMERIC CELL
                // =====================================================

                function replaceNumericCellValue(
                    cellAddress,
                    value
                ) {

                    const cellRegex = new RegExp(
                        `<c\\s+[^>]*r="${cellAddress}"[^>]*>[\\s\\S]*?<\\/c>`,
                        "i"
                    );

                    const match =
                        worksheetXML.match(cellRegex);

                    if (!match) {
                        throw new Error(
                            `Could not find cell ${cellAddress}.`
                        );
                    }

                    const originalCell = match[0];

                    // Remove shared-string type
                    let updatedCell =
                        originalCell.replace(
                            /\s+t="[^"]*"/i,
                            ""
                        );

                    // Replace existing value
                    if (/<v>[\s\S]*?<\/v>/i.test(updatedCell)) {

                        updatedCell =
                            updatedCell.replace(
                                /<v>[\s\S]*?<\/v>/i,
                                `<v>${value}</v>`
                            );

                    } else {

                        // Add value to blank cell
                        updatedCell =
                            updatedCell.replace(
                                /<\/c>/i,
                                `<v>${value}</v></c>`
                            );
                    }

                    worksheetXML =
                        worksheetXML.replace(
                            originalCell,
                            updatedCell
                        );
                }


                // =====================================================
                // 12. REPLACE BLANK CELL
                // =====================================================

                function replaceBlankCellValue(
                    cellAddress,
                    newIndex
                ) {

                    const cellRegex = new RegExp(
                        `<c\\s+[^>]*r="${cellAddress}"[^>]*/>`,
                        "i"
                    );

                    const match = worksheetXML.match(cellRegex);

                    if (!match) {
                        throw new Error(
                            `Could not find blank cell ${cellAddress}.`
                        );
                    }

                    const originalCell = match[0];

                    const updatedCell =
                        originalCell.replace(
                            "/>",
                            ` t="s"><v>${newIndex}</v></c>`
                        );

                    worksheetXML = worksheetXML.replace(
                        originalCell,
                        updatedCell
                    );
                }


                // =====================================================
                // 13. UPDATE EXCEL CELLS
                // =====================================================

                // Company Info
                replaceExistingCellValue("A12", indexes.CIrep);
                replaceExistingCellValue("A13", indexes.CIcompany);
                replaceExistingCellValue("A14", indexes.CIstreet);
                replaceExistingCellValue("A15", indexes.CIcity);
                replaceExistingCellValue("C15", indexes.CIstate);
                replaceExistingCellValue("E15", indexes.CIzip);
                replaceBlankCellValue("B16", indexes.CIemail);
                replaceBlankCellValue("B17", indexes.CIphone);


                // Pick Up Info
                replaceExistingCellValue("B18", indexes.PUcompany);
                replaceExistingCellValue("B19", indexes.PUstreet);
                replaceExistingCellValue("B20", indexes.PUcity);
                replaceExistingCellValue("C20", indexes.PUstate);
                replaceExistingCellValue("E20", indexes.PUzip);
                replaceExistingCellValue("B21", indexes.PUdate);
                replaceExistingCellValue("C21", indexes.PUtime);


                // Delivery Info
                replaceExistingCellValue("G18", indexes.Delcompany);
                replaceExistingCellValue("G19", indexes.Delstreet);
                replaceExistingCellValue("G20", indexes.Delcity);
                replaceExistingCellValue("I20", indexes.Delstate);
                replaceExistingCellValue("K20", indexes.Delzip);
                replaceExistingCellValue("G21", indexes.Deldate);
                replaceExistingCellValue("I21", indexes.Deltime);
                replaceExistingCellValue("B31", indexes.DelPoc);


                // Service Info
                replaceExistingCellValue("F16", indexes.SIservices);
                replaceExistingCellValue("A23", indexes.SIdesc);
                replaceNumericCellValue("G23", formData.SIqty);
                replaceNumericCellValue("H23", formData.SIunitPrice);
                replaceExistingCellValue("A24", indexes.SIdesc2);
                replaceNumericCellValue("G24", formData.SIqty2);
                replaceNumericCellValue("H24", formData.SIunitPrice2);


                // Financial Info
                replaceExistingCellValue("F14", indexes.FImethod);
                replaceNumericCellValue("J28", formData.FIdiscount);
                replaceNumericCellValue("J30", formData.FIadjustments);


                // Receipt Number
                replaceBlankCellValue("I12", indexes.INfinal);


                // Additional Info
                replaceExistingCellValue("B28", indexes.Notes);
                replaceExistingCellValue("F9", indexes.SubDate);


                // =====================================================
                // 14. PUT XML BACK INTO XLSX
                // =====================================================

                zip.file("xl/worksheets/sheet1.xml", worksheetXML);

                zip.file("xl/sharedStrings.xml", sharedStringsXML);


                // =====================================================
                // 15. FORCE EXCEL TO RECALCULATE
                // =====================================================

                if (/<calcPr\b/i.test(workbookXML)) {

                    workbookXML =
                        workbookXML.replace(
                            /<calcPr\b[^>]*\/>/i,
                            '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
                        );

                } else {

                    workbookXML =
                        workbookXML.replace(
                            /<\/workbook>/i,
                            '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'
                        );
                }

                zip.file("xl/workbook.xml", workbookXML);


                // =====================================================
                // 16. GENERATE XLSX
                // =====================================================

                const outputBlob =
                    await zip.generateAsync({
                        type: "blob"
                    });


                // =====================================================
                // 17. DOWNLOAD RECEIPT
                // =====================================================

                const downloadURL = URL.createObjectURL(outputBlob);
                const link = document.createElement("a");

                link.href = downloadURL;

                link.download = `${formData.INfinal} Receipt.xlsx`;

                document.body.appendChild(link);

                link.click();

                document.body.removeChild(link);

                URL.revokeObjectURL(downloadURL);


                console.log("Receipt created successfully.");

            } catch (error) {

                console.error("Error creating receipt:", error);

                alert(
                    "There was a problem creating the receipt.\n\n" +
                    error.message
                );
            }
        }
    );


    // =========================================================
    // XML ESCAPING
    // =========================================================

    function escapeXML(value) {
        return String(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

});