function loadFile(url,callback){
    PizZipUtils.getBinaryContent(url,callback);
}
function generate() {

    loadFile("https://docxtemplater.com/tag-example.docx",function(error,content){
        if (error) { throw error };

        // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
        function replaceErrors(key, value) {
            if (value instanceof Error) {
                return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                    error[key] = value[key];
                    return error;
                }, {});
            }
            return value;
        }

        function errorHandler(error) {
            console.log(JSON.stringify({error: error}, replaceErrors));

            if (error.properties && error.properties.errors instanceof Array) {
                const errorMessages = error.properties.errors.map(function (error) {
                    return error.properties.explanation;
                }).join("\n");
                console.log('errorMessages', errorMessages);
                // errorMessages is a humanly readable message looking like this :
                // 'The tag beginning with "foobar" is unopened'
            }
            throw error;
        }

        var zip = new PizZip(content);
        var doc;
        try {
            doc=new window.docxtemplater(zip);
        } catch(error) {
            // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
            errorHandler(error);
        }

        doc.setData({
            doc_location: 'Mumbai',
            doc_day: 'Monday',
            doc_day_of: '21-06,1995',
            doc_landlord_detail: 'Dhanasekaran'
        });
        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render();
        }
        catch (error) {
            // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
            errorHandler(error);
        }

        var out=doc.getZip().generate({
            type:"blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }) //Output the document using Data-URI
        saveAs(out,"output.docx")
    })
}

function generateOnFile() {
    var docs = document.getElementById('fileToUpload');

    if (validateForm() == false) {
        alert("Fill all the Fields");
        return;
    }
    var reader = new FileReader();
    if (docs.files.length === 0) {
        alert("No files selected")
    }
    reader.readAsBinaryString(docs.files.item(0));

    reader.onerror = function (evt) {
        console.log("error reading file", evt);
        alert("error reading file" + evt)
    }
    reader.onload = function (evt) {
        const content = evt.target.result;
        var zip = new PizZip(content);
        var doc;
        try {
            doc=new window.docxtemplater(zip);
        } catch(error) {
            // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
            errorHandler(error);
        }

        var location = document.getElementById("doc_location").value;
        var documentedDay = document.getElementById("doc_day").value;

        var enterdDocumentDate = formatDate(documentedDay)

        var landlordDetail = document.getElementById("doc_landlord_detail").value;
        var landlordAge = document.getElementById("doc_landlord_age").value;
        var landlordResidingAt = document.getElementById("doc_landlord_residing_at").value; 

        var tenantdDetail = document.getElementById("doc_tenant_detail").value;
        var tenantAge = document.getElementById("doc_tenant_age").value;
        var tenantResidingAt = document.getElementById("doc_tenant_residing_at").value;  

        var addressOfRent = document.getElementById("doc_address_of_rented").value; 
        var numberOfBedRooms = document.getElementById("doc_number_of_bed_room").value; 
        var descriptionOfRentedPremise = document.getElementById("doc_description_of_rented_premise").value; 

        var demisedCommenceFrom = document.getElementById("doc_demised_premise_commence_from").value;
        var demisedValidTill = document.getElementById("doc_demised_premise_valid_till").value;
        var demisedPeriodYears = document.getElementById("doc_demised_premise_period_years").value;

        var enteredDemisedCommenceFrom = formatDate(demisedCommenceFrom);
        var enteredDemisedValidTill = formatDate(demisedValidTill);

        var monthlyRent = document.getElementById("doc_monthly_rent").value;
        var rentLastDate = document.getElementById("doc_rent_last_date").value;
        var maintanceCharge = document.getElementById("doc_maintainance_charge").value;
        var detailsOfFurnishing = document.getElementById("doc_detail_of_furnishing").value;
        var securityDeposit = document.getElementById("doc_security_deposit").value;

        var rentalAgreementLockinPeriod = document.getElementById("doc_rental_agreement_lockin_period").value;
        var rentalAgreementTerminatingPeriod = document.getElementById("doc_rental_agreement_terminate_period").value;
        var rentalAgreementTerminatingNoticePeriod = document.getElementById("doc_rental_agreement_terminate_notice_period").value;
        var arrearPaymentPeriod = document.getElementById("doc_arrear_payment_period").value;

        var subjectOfJuri = document.getElementById("doc_subj_of_juri").value;
        var executedDay = document.getElementById("doc_executed_day").value;

        var enteredExecutedDate = formatDate(executedDay)

        
        doc.setData({
            doc_location: String(location),
            doc_day: String(enterdDocumentDate),
            
            doc_landlord_detail: String(landlordDetail),
            doc_landlord_age: String(landlordAge),
            doc_landlord_residing_at: String(landlordResidingAt),

            doc_tenant_detail: String(tenantdDetail),
            doc_tenant_age: String(tenantAge),
            doc_tenant_residing_at: String(tenantResidingAt),

            doc_address_of_rented: String(addressOfRent),
            doc_number_of_bed_room: String(numberOfBedRooms),
            doc_description_of_rented_premise: String(descriptionOfRentedPremise),
            doc_address_of_rented_premise: String(addressOfRent),

            doc_demised_premise_commence_from: String(enteredDemisedCommenceFrom),
            doc_demised_premise_valid_till: String(enteredDemisedValidTill),
            doc_demised_premise_period_years: String(demisedPeriodYears),

            doc_monthly_rent: String(monthlyRent),
            doc_rent_last_date: String(rentLastDate),
            doc_maintainance_charge: String(maintanceCharge),
            doc_detail_of_furnishing: String(detailsOfFurnishing),
            doc_security_deposit: String(securityDeposit),

            doc_rental_agreement_lockin_period: String(rentalAgreementLockinPeriod),
            doc_rental_agreement_terminate_period: String(rentalAgreementTerminatingPeriod),
            doc_rental_agreement_terminate_notice_period: String(rentalAgreementTerminatingNoticePeriod),
            doc_arrear_payment_period: String(arrearPaymentPeriod),

            doc_subj_of_juri: String(subjectOfJuri),
            doc_executed_day: String(enteredExecutedDate),
        });
        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render();
        }
        catch (error) {
            // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
            errorHandler(error);
        }

        var out=doc.getZip().generate({
            type:"blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }) 
        var fileName = String(document.getElementById("exportingFilename").value);
        if(fileName) {
            fileName = fileName + ".docx"
            saveAs(out, fileName)
        } else {
            saveAs(out, "Rental Agreement.docx")
        }
        
        resetThisForm();
    }
}

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (day.length < 2) day = '0' + day;

    var dayString = ordinal_suffix_of(day)

return String(dayString) + " day of " + String(month_name(month)) + ", " + year;
}

function ordinal_suffix_of(i) {
    var j = i % 10,
        k = i % 100;
    if (j == 1 && k != 11) {
        return i + "st";
    }
    if (j == 2 && k != 12) {
        return i + "nd";
    }
    if (j == 3 && k != 13) {
        return i + "rd";
    }
    return i + "th";
}

function month_name(i) {
    switch(Number(i))
    {
        case 1: return "January";
        case 2: return "February";
        case 3: return "March"
        case 4: return "April";
        case 5: return "May";
        case 6: return "June";
        case 7: return "July";
        case 8: return "August";
        case 9: return "September";
        case 10: return "October";
        case 11: return "November";
        case 12: return "December";
        default: return "January"
    }
}

function resetThisForm() {
    var docs = document.getElementById("register-form");
    docs.reset();
}

function downloadTemplateDocument() {
    window.open("https://drive.google.com/uc?id=1ek5_LYQ2tqrEoAzr7-ty1GIy1s1q-fQS&export=download");
}

function validateForm() {
    var location = document.getElementById("doc_location");
    var documentedDay = document.getElementById("doc_day");

    var landlordDetail = document.getElementById("doc_landlord_detail");
    var landlordAge = document.getElementById("doc_landlord_age");
    var landlordResidingAt = document.getElementById("doc_landlord_residing_at"); 

    var tenantdDetail = document.getElementById("doc_tenant_detail");
    var tenantAge = document.getElementById("doc_tenant_age");
    var tenantResidingAt = document.getElementById("doc_tenant_residing_at");  

    var addressOfRent = document.getElementById("doc_address_of_rented"); 
    var numberOfBedRooms = document.getElementById("doc_number_of_bed_room"); 
    var descriptionOfRentedPremise = document.getElementById("doc_description_of_rented_premise"); 

    var demisedCommenceFrom = document.getElementById("doc_demised_premise_commence_from");
    var demisedValidTill = document.getElementById("doc_demised_premise_valid_till");
    var demisedPeriodYears = document.getElementById("doc_demised_premise_period_years");

    var monthlyRent = document.getElementById("doc_monthly_rent");
    var rentLastDate = document.getElementById("doc_rent_last_date");
    var maintanceCharge = document.getElementById("doc_maintainance_charge");
    var detailsOfFurnishing = document.getElementById("doc_detail_of_furnishing");
    var securityDeposit = document.getElementById("doc_security_deposit");

    var rentalAgreementLockinPeriod = document.getElementById("doc_rental_agreement_lockin_period");
    var rentalAgreementTerminatingPeriod = document.getElementById("doc_rental_agreement_terminate_period");
    var rentalAgreementTerminatingNoticePeriod = document.getElementById("doc_rental_agreement_terminate_notice_period");
    var arrearPaymentPeriod = document.getElementById("doc_arrear_payment_period");

    var subjectOfJuri = document.getElementById("doc_subj_of_juri");
    var executedDay = document.getElementById("doc_executed_day");

    if(location.value && documentedDay.value &&  landlordDetail.value && landlordAge.value && landlordResidingAt.value &&  tenantdDetail.value && tenantAge.value && tenantResidingAt.value && addressOfRent.value && descriptionOfRentedPremise.value && demisedCommenceFrom.value  && demisedValidTill.value  && demisedPeriodYears.value && monthlyRent.value && rentLastDate.value && maintanceCharge.value && securityDeposit.value && rentalAgreementLockinPeriod.value && rentalAgreementTerminatingPeriod.value && rentalAgreementTerminatingNoticePeriod.value && arrearPaymentPeriod.value && subjectOfJuri.value && executedDay.value) {
        return true;
    } else {
        return false;
    }
}