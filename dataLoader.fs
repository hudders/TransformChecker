module dataLoader

open System
open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open config
open test_utils
open csvLoader

let rowList = [3..125]

let loadData (dataType : string, xmlFile : string, a : int, b : int, c : int) =
    let dataSrc =
        if dataType = "Proposer" then
            xlsLoader.getXLS("2.1 Proposer");
        elif dataType = "Additional" then
            xlsLoader.getXLS("2.2 Additional Driver")
        elif dataType = "Claim" then
            xlsLoader.getXLS("2.3 Claims")
        elif dataType = "Conviction" then
            xlsLoader.getXLS("2.4 Convictions")
        elif dataType = "Vehicle" then
            xlsLoader.getXLS("2.5 Vehicle")
        elif dataType = "Policy" then
            xlsLoader.getXLS("2.6 Policy")
        else
            xlsLoader.getXLS("")
    for xlRow in rowList do
        if xlsLoader.cellValue(dataSrc, "A", xlRow) <> null && xlsLoader.cellValue(dataSrc, "A", xlRow).ToString() = personCollection.[c].TestId then
            let fName = xlsLoader.cellValue(dataSrc, "B", xlRow).ToString()
            let xVal = xlsLoader.cellValue(dataSrc, "D", xlRow).ToString()
            let xLoc = xlsLoader.cellValue(dataSrc, "E", xlRow)
            if xLoc = null then
                let notApplicableCell = "I" + xlRow.ToString()
                dataSrc.Range(notApplicableCell, notApplicableCell).Value2 <- "Not Applicable"
            else
                let xLoc = xLoc.ToString()
                let xLoc, xVal = xLoc.Split('\n'), xVal.Split('\n')
                let rec loop n =
                    let xValue =
                        if xVal.[n] = "as input" then
                            if dataType = "Claim" then
                                if fName = "Cost" then
                                    if claimCollection.[a].ClaimCost <> "" then
                                        if contains (xlsLoader.brandGroup(xlsLoader.brand)) ["CDL-FilterFree"] then
                                            claimCollection.[a].ClaimCost + ".00"
                                        else
                                            claimCollection.[a].ClaimCost
                                    else
                                        "0"
                                else
                                    ""
                            elif dataType = "Conviction" then
                                if fName = "Type" then
                                    convictionCollection.[a].Conviction
                                elif fName = "Licence Points" then
                                    convictionCollection.[a].NumberOfPoints
                                elif fName = "Ban (months)" then
                                    convictionCollection.[a].BanLength
                                elif fName = "Fine £" then
                                    convictionCollection.[a].FineAmount
                                elif fName = "If yes: what was the breathalyser reading?" then
                                    convictionCollection.[a].BreathalyserReading
                                else
                                    ""
                            elif dataType = "Vehicle" then
                                if fName = "Registration Year and Letter" then
                                    stripChars vehicleCollection.[a].RegistrationNumber " "
                                elif fName = "Manufacturer" then
                                    xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[1]
                                elif fName = "Model" then
                                    xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[2]
                                elif fName = "Style" then
                                    xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[3]
                                elif fName = "Engine" then
                                    xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[5]
                                elif fName = "Trim" then
                                    xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[6]
                                elif fName = "doors" then
                                    xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[7]
                                elif fName = "seats" then
                                    personCollection.[b].TestId
                                elif fName = "Private Mileage" then
                                    carUsageCollection.[b].AnnualMileage
                                elif fName = "Business Mileage" then
                                    carUsageCollection.[b].AnnualBusinessMileage
                                elif fName = "Vehicle Value" then
                                    vehicleCollection.[a].CurrentValue
                                else
                                    ""
                            elif dataType = "Proposer" then
                                if fName = "Firstname" || fName = "Forename" then
                                    personCollection.[a].FirstName
                                elif fName = "Lastname" || fName = "Surname" then
                                    personCollection.[a].LastName
                                elif fName = "Main telephone number" then
                                    stripChars personCollection.[a].MainTelephoneNumber " "
                                elif fName = "Email" || fName = "email" then
                                    personCollection.[a].Email
                                elif fName = "Occupation" then
                                    if personCollection.[a].OccupationTitleDescription <> "" then
                                        personCollection.[a].OccupationTitleDescription
                                    else
                                        "Unemployed"
                                 elif fName = "Type of Business" then
                                    if personCollection.[a].BusinessTypeDescription <> "" then
                                        personCollection.[a].BusinessTypeDescription
                                    else
                                        "Not In Employment"
                                else
                                    ""
                            elif dataType = "Additional" then
                                if fName = "FirstName" || fName = "Forename" then
                                    additionalCollection.[a].FirstName
                                elif fName = "Lastname" || fName = "Surname" then
                                    additionalCollection.[a].LastName
                                elif fName = "Occupation" then
                                    if additionalCollection.[a].OccupationTitleDescription <> "" then
                                        additionalCollection.[a].OccupationTitleDescription
                                    else
                                        "Unemployed"
                                elif fName = "Type of Business" then
                                    if additionalCollection.[a].BusinessTypeDescription <> "" then
                                        additionalCollection.[a].BusinessTypeDescription
                                    else
                                        "Not In Employment"
                                else
                                    ""
                            else
                                ""
                        elif contains xVal.[n] ["<conCode>"; "<carCode>"; "<styleCode>"; "<occCode>"; "<empCode>"] then
                            if dataType = "Conviction" then
                                if xlsLoader.codeLookup(convictionCollection.[a].Conviction, xVal.[n]) <> None then
                                    xlsLoader.codeLookup(convictionCollection.[a].Conviction, xVal.[n]).Value
                                else
                                    ""
                            elif dataType = "Vehicle" then
                                let strippedReg = (stripChars vehicleCollection.[a].RegistrationNumber " ")
                                if xlsLoader.codeLookup(strippedReg, xVal.[n]) <> None then
                                    xlsLoader.codeLookup(strippedReg, xVal.[n]).Value
                                else
                                    ""
                            elif dataType = "Proposer" then
                                if xVal.[n] = "<occCode>" then
                                    let occupation =
                                        if personCollection.[a].OccupationTitleDescription <> "" then
                                            personCollection.[a].OccupationTitleDescription
                                        else
                                            "Unemployed"
                                    if xlsLoader.codeLookup(occupation, xVal.[n]) <> None then
                                        xlsLoader.codeLookup(occupation, xVal.[n]).Value
                                    else
                                        ""
                                elif xVal.[n] = "<empCode>" then
                                    let business =
                                        if personCollection.[a].BusinessTypeDescription <> "" then
                                            personCollection.[a].BusinessTypeDescription
                                        else
                                            "Not In Employment"
                                    if xlsLoader.codeLookup(business, xVal.[n]) <> None then
                                        xlsLoader.codeLookup(business, xVal.[n]).Value
                                    else
                                        ""
                                else
                                    ""
                            elif dataType = "Additional" then
                                if xVal.[n] = "<occCode>" then
                                    let occupation =
                                        if additionalCollection.[a].OccupationTitleDescription <> "" then
                                            additionalCollection.[a].OccupationTitleDescription
                                        else
                                            "Unemployed"
                                    if xlsLoader.codeLookup(occupation, xVal.[n]) <> None then
                                        xlsLoader.codeLookup(occupation, xVal.[n]).Value
                                    else
                                        ""
                                elif xVal.[n] = "<empCode>" then
                                    let business =
                                        if additionalCollection.[a].BusinessTypeDescription <> "" then
                                            additionalCollection.[a].BusinessTypeDescription
                                        else
                                            "Not In Employment"
                                    if xlsLoader.codeLookup(business, xVal.[n]) <> None then
                                        xlsLoader.codeLookup(business, xVal.[n]).Value
                                    else
                                        ""
                                else
                                    ""
                            else
                                ""
                        elif contains xVal.[n] ["YYYY-MM-DD"; "YYYY-MM-01"; "DD-MM-YYYY"] then
                            let dateOffset = xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()
                            if dataType = "Claim" then
                                if fName = "Date" then
                                    dateTester (xVal.[n], "", claimCollection.[a].ClaimDateMonthCode, claimCollection.[a].ClaimDateYear)
                                else
                                    ""
                            elif dataType = "Conviction" then
                                if fName = "Date" then
                                    dateTester (xVal.[n], "", convictionCollection.[a].ConvictionDateMonthCode, convictionCollection.[a].ConvictionDateYear)
                                else
                                    ""
                            elif dataType = "Vehicle" then
                                if fName = "Date of Purchase" then
                                    dateTester (xVal.[n], todaysDay, carUsageCollection.[b].DateOfPurchaseMonthCode, carUsageCollection.[b].DateOfPurchaseYear)
                                elif fName = "Not purchased yet" then
                                    if contains (xlsLoader.brandGroup(xlsLoader.brand)) ["CDL-FilterFree"] then
                                        dateTester (xVal.[n], todaysDay, todaysMonth, todaysYear)
                                    else
                                        let dateFormat = Regex.Replace(xVal.[n].ToLower(),"mm","MM")
                                        ((DateTime.Parse(todaysDay + "/" + todaysMonth + "/" + todaysYear)).AddDays ((float personCollection.[b].PolicyStartDateOffset) - float 1)).ToString(dateFormat)
                                else
                                    ""
                            elif dataType = "Policy" then
                                if fName = "Cover Start Date" then
                                    let dateFormat = Regex.Replace(xVal.[n].ToLower(),"mm","MM")
                                    ((DateTime.Parse(todaysDay + "/" + todaysMonth + "/" + todaysYear)).AddDays ((float personCollection.[b].PolicyStartDateOffset) - float 1)).ToString(dateFormat)
                                else
                                    ""
                            elif dataType = "Proposer" then
                                if fName = "DOB" || fName = "Date of Birth" || dateOffset = "since birth" then
                                    dateTester (xVal.[n], personCollection.[a].DateOfBirthDay, personCollection.[a].DateOfBirthMonth, personCollection.[a].DateOfBirthYear)
                                elif dateOffset = "or since MM/YYYY" then
                                    dateTester (xVal.[n], "01", personCollection.[a].IsLivingInUkSinceMonthCode, personCollection.[a].IsLivingInUkSinceYear)
                                elif fName = "Period licence held for?" || fName = "Licence Date Obtained" then
                                    let licenceYear =
                                        if personCollection.[a].LicenceDateYear.Length > 0 then
                                            personCollection.[a].LicenceDateYear
                                        else
                                            (System.Convert.ToInt32(todaysYear) - System.Convert.ToInt32(personCollection.[a].LicenceHeldCode) + 1).ToString()
                                    dateTester (xVal.[n], personCollection.[a].LicenceDateDay, personCollection.[a].LicenceDateMonth, licenceYear)
                                else
                                    ""
                            elif dataType = "Additional" then
                                if fName = "DOB" || fName = "Date of Birth" || dateOffset = "since birth" then
                                    dateTester (xVal.[n], additionalCollection.[a].DateOfBirthDay, additionalCollection.[a].DateOfBirthMonth, additionalCollection.[a].DateOfBirthYear)
                                elif dateOffset = "or since MM/YYYY" then
                                    dateTester (xVal.[n], "01", additionalCollection.[a].IsLivingInUkSinceMonthCode, additionalCollection.[a].IsLivingInUkSinceYear)
                                elif fName = "Period licence held for?" || fName = "Licence Date Obtained" then
                                    let licenceYear =
                                        if additionalCollection.[a].LicenceDateYear.Length > 0 then
                                            additionalCollection.[a].LicenceDateYear
                                        else
                                            (System.Convert.ToInt32(todaysYear) - System.Convert.ToInt32(additionalCollection.[a].LicenceHeldCode) + 1).ToString()
                                    dateTester (xVal.[n], additionalCollection.[a].LicenceDateDay, additionalCollection.[a].LicenceDateMonth, licenceYear)
                                else
                                    ""
                            else ""
                        elif xVal.[n] = "<age>" then
                            let dateOffset = xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()
                            let coverStart = (((DateTime.Parse(todaysDay + "/" + todaysMonth + "/" + todaysYear)).AddDays ((float personCollection.[a].PolicyStartDateOffset) - float 1)).ToString("dd/MM/yyyy")).Split('/')
                            let todaysYearInt = System.Convert.ToInt32(coverStart.[2])
                            let todaysMonthInt = System.Convert.ToInt32(coverStart.[1])
                            let todaysDayInt = System.Convert.ToInt32(coverStart.[0])
                            if dataType = "Proposer" then
                                if dateOffset = "since birth" then
                                    let dateOfBirthYearInt = System.Convert.ToInt32(personCollection.[a].DateOfBirthYear)
                                    let dateOfBirthMonthInt = System.Convert.ToInt32(personCollection.[a].DateOfBirthMonth)
                                    let dateOfBirthDayInt = System.Convert.ToInt32(personCollection.[a].DateOfBirthDay)
                                    if todaysMonthInt < dateOfBirthMonthInt || ((todaysMonthInt = dateOfBirthMonthInt) && (todaysDayInt < dateOfBirthDayInt)) then
                                        (todaysYearInt - (dateOfBirthYearInt + 1)).ToString()
                                    else
                                        (todaysYearInt - dateOfBirthYearInt).ToString()
                                elif dateOffset = "or since MM/YYYY" then
                                    let todaysYear = System.Convert.ToInt32(todaysYear)
                                    let sinceYear = System.Convert.ToInt32(personCollection.[a].IsLivingInUkSinceYear)
                                    if todaysMonthInt < System.Convert.ToInt32(personCollection.[a].IsLivingInUkSinceMonthCode) then
                                        ((todaysYear - (sinceYear + 1)).ToString())
                                    else
                                        ((todaysYear - sinceYear).ToString())
                                else
                                    ""
                            elif dataType = "Additional" then
                                if dateOffset = "since birth" then
                                    let dateOfBirthYearInt = System.Convert.ToInt32(additionalCollection.[a].DateOfBirthYear)
                                    let dateOfBirthMonthInt = System.Convert.ToInt32(additionalCollection.[a].DateOfBirthMonth)
                                    let dateOfBirthDayInt = System.Convert.ToInt32(additionalCollection.[a].DateOfBirthDay)
                                    if todaysMonthInt < dateOfBirthMonthInt || ((todaysMonthInt = dateOfBirthMonthInt) && (todaysDayInt < dateOfBirthDayInt)) then
                                        (todaysYearInt - (dateOfBirthYearInt + 1)).ToString()
                                    else
                                        (todaysYearInt - dateOfBirthYearInt).ToString()
                                elif dateOffset = "or since MM/YYYY" then
                                    let todaysYear = System.Convert.ToInt32(todaysYear)
                                    let sinceYear = System.Convert.ToInt32(additionalCollection.[a].IsLivingInUkSinceYear)
                                    if todaysMonthInt < System.Convert.ToInt32(additionalCollection.[a].IsLivingInUkSinceMonthCode) then
                                        ((todaysYear - (sinceYear + 1)).ToString())
                                    else
                                        ((todaysYear - sinceYear).ToString())
                                else
                                    ""
                            else
                                ""
                        elif xVal.[n] = "YYYY" then
                            if dataType = "Vehicle" then
                                if fName = "Registration Year and Letter" || fName = "Manufacturer" then
                                    vehicleCollection.[a].RegYear
                                else
                                    ""
                            else
                                ""
                        else
                            xVal.[n]
                    xmlLoader.checkXml (xValue, xLoc.[n], xmlFile, dataSrc, xlRow)
                    //printfn "%i === %i  === %s" xLoc.Length n xVal.[n]
                    if xLoc.Length <> n + 1 then
                        loop (n + 1)
                loop 0
                let rec ResultLoop n =
                    if xlsLoader.cellValue(xlsLoader.dataExt, "K", n) <> null then
                        if xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString() = "Pass" then
                            fillCell(dataSrc, "G", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString())
                            fillCell(dataSrc, "H", xlRow, "")
                            fillCell(dataSrc, "J", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "I", n).ToString())
                            if xlsLoader.cellValue(xlsLoader.dataExt, "L", n) <> null then
                                fillCell(dataSrc, "K", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "L", n).ToString())
                        else
                            if xlsLoader.cellValue(dataSrc, "G", xlRow) = null then
                                fillCell(dataSrc, "H", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString())
                                fillCell(dataSrc, "J", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "I", n).ToString())
                                if xlsLoader.cellValue(xlsLoader.dataExt, "L", n) <> null then
                                    fillCell(dataSrc, "K", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "L", n).ToString())
                        if xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString() <> "Pass" then
                            ResultLoop (n + 1)
                ResultLoop 59
                let resultThing =
                    if xlsLoader.cellValue(dataSrc, "J", xlRow) <> null then
                        if xlsLoader.cellValue(dataSrc, "G", xlRow) <> null then
                            "."
                        else
                            "F"
                    else
                        "S"
                printf "%s" resultThing
                cleanUp()