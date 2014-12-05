module dataLoader

open System
open System.IO
open System.Text.RegularExpressions

open config
open csvLoader
open test_utils

let rowList = [3..118]//250]

let loadData (product : string, dataType : string, xmlFile : string, a : int, b : int, c : int) =
    let dataSrc =
        match product, dataType with
            | "PC", "Proposer" | "HH", "Proposer" -> xlsLoader.getXLS("2.1 " + dataType)
            | "PC", "Additional Driver" | "HH", "Joint Proposer" -> xlsLoader.getXLS("2.2 " + dataType)
            | "PC", "Claims" | "HH", "Property" -> xlsLoader.getXLS("2.3 " + dataType)
            | "PC", "Convictions" | "HH", "Contents Cover" -> xlsLoader.getXLS("2.4 " + dataType)
            | "PC", "Vehicle" | "HH", "Buildings Cover" -> xlsLoader.getXLS("2.5 " + dataType)
            | "PC", "Policy" | "HH", "Locks and Security" -> xlsLoader.getXLS("2.6 " + dataType)
            | "HH", "Claims" -> xlsLoader.getXLS("2.7 " + dataType)
            | "HH", "Price Page" -> xlsLoader.getXLS("2.8 " + dataType)
            | _ -> xlsLoader.getXLS("")

    for xlRow in rowList do
        if xlsLoader.cellValue(dataSrc, "A", xlRow) <> null && xlsLoader.cellValue(dataSrc, "A", xlRow).ToString() = personCollection.[c].TestId then
            let fName = xlsLoader.cellValue(dataSrc, "B", xlRow).ToString()
            let xVal = xlsLoader.cellValue(dataSrc, "D", xlRow)
            let xLoc = xlsLoader.cellValue(dataSrc, "E", xlRow)
            if xLoc = null || xVal = null then
                dataSrc.Range(("I" + xlRow.ToString()), ("I" + xlRow.ToString())).Value2 <- "Not Applicable"
            else
                let xLoc, xVal = (xLoc.ToString()).Split('\n'), (xVal.ToString()).Split('\n')
                let rec loop n =
                    //printfn "%s" xLoc.[n]
                    let xValue =
                        match xVal.[n] with
                        | "as input" -> match fName with
                                        // Proposer & Additional
                                        | "Firstname" | "First Name" | "Forename"   -> match dataType with
                                                                                       | "Proposer"          -> personCollection.[a].FirstName
                                                                                       | "Additional Driver" -> additionalCollection.[a].FirstName
                                                                                       | _ -> ""
                                        | "Lastname" | "Last Name" | "Surname"      -> match dataType with
                                                                                       | "Proposer"          -> personCollection.[a].LastName
                                                                                       | "Additional Driver" -> additionalCollection.[a].LastName
                                                                                       | _ -> ""
                                        | "Main Telephone Number"
                                        | "Main telephone number"                   -> personCollection.[a].MainTelephoneNumber
                                        | "Email" | "E-mail"                        -> personCollection.[a].Email
                                        | "Occupation"                              -> match dataType with
                                                                                       | "Proposer"           -> if personCollection.[a].OccupationTitleDescription <> ""
                                                                                                                 then personCollection.[a].OccupationTitleDescription
                                                                                                                 else "Unemployed"
                                                                                       | "Additional Driver"  -> if additionalCollection.[a].OccupationTitleDescription <> ""
                                                                                                                 then additionalCollection.[a].OccupationTitleDescription
                                                                                                                 else "Unemployed"
                                                                                       | _ -> ""
                                        | "Type of Business"                        -> match dataType with
                                                                                       | "Proposer"           -> if personCollection.[a].BusinessTypeDescription <> ""
                                                                                                                 then personCollection.[a].BusinessTypeDescription
                                                                                                                 else "Not In Employment"
                                                                                       | "Additional Driver"  -> if additionalCollection.[a].BusinessTypeDescription <> ""
                                                                                                                 then additionalCollection.[a].BusinessTypeDescription
                                                                                                                 else "Not In Employment"
                                                                                       | _ -> ""
                                        // Claim
                                        | "Cost"                                    -> if claimCollection.[a].ClaimCost <> ""
                                                                                       then match (xlsLoader.brandGroup(xlsLoader.brand)) with
                                                                                            | "CDL-FilterFree" | "OpenGI" -> claimCollection.[a].ClaimCost + ".00"
                                                                                            | _                           -> claimCollection.[a].ClaimCost
                                                                                       else "0"
                                        // Conviction
                                        | "Type" -> convictionCollection.[a].Conviction
                                        | "Licence Points" -> convictionCollection.[a].NumberOfPoints
                                        | "Ban (months)" -> convictionCollection.[a].BanLength
                                        | "Fine £" -> convictionCollection.[a].FineAmount
                                        | "If yes: what was the breathalyser reading?" -> convictionCollection.[a].BreathalyserReading
                                        // Vehicle
                                        | "Registration Year and Letter" -> stripChars vehicleCollection.[a].RegistrationNumber " "
                                        | "ABI Code" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[0]
                                        | "Manufacturer" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[1]
                                        | "Model" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[2]
                                        | "Style" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[3]
                                        | "Engine Capacity" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[5]
                                        | "Trim" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[6]
                                        | "doors" -> xlsLoader.carDetailLookup((stripChars vehicleCollection.[a].RegistrationNumber " ")).[7]
                                        | "seats" -> personCollection.[b].TestId
                                        | "Private Mileage" -> carUsageCollection.[b].AnnualMileage
                                        | "Business Mileage" -> carUsageCollection.[b].AnnualBusinessMileage
                                        | "Vehicle Value" -> vehicleCollection.[a].CurrentValue
                                        //| "Vehicle kept at <Postcode>" -> personCollection.[c].
                                        | _ -> ""
                        // Codes
                        | "<conCode>" | "<conCode1>" | "<conCode2>"
                                                        -> if xlsLoader.codeLookup(convictionCollection.[a].Conviction, xVal.[n]) <> None
                                                           then xlsLoader.codeLookup(convictionCollection.[a].Conviction, xVal.[n]).Value
                                                           else ""
                        | "<carCode>" | "<styleCode>"   -> let strippedReg = (stripChars vehicleCollection.[a].RegistrationNumber " ")
                                                           if xlsLoader.codeLookup(strippedReg, xVal.[n]) <> None
                                                           then xlsLoader.codeLookup(strippedReg, xVal.[n]).Value
                                                           else ""
                        | "<occCode>"                   -> let occupation = match dataType with
                                                                            | "Proposer"          -> if personCollection.[a].OccupationTitleDescription <> "" 
                                                                                                     then personCollection.[a].OccupationTitleDescription 
                                                                                                     else "Unemployed"
                                                                            | "Additional Driver" -> if additionalCollection.[a].OccupationTitleDescription <> "" 
                                                                                                     then additionalCollection.[a].OccupationTitleDescription
                                                                                                     else "Unemployed"
                                                                            | _                   -> "Unemployed"
                                                           if xlsLoader.codeLookup(occupation, xVal.[n]) <> None
                                                           then xlsLoader.codeLookup(occupation, xVal.[n]).Value
                                                           else ""
                        | "<empCode>"                   -> let business = match dataType with
                                                                            | "Proposer"          -> if personCollection.[a].BusinessTypeDescription <> ""
                                                                                                     then personCollection.[a].BusinessTypeDescription
                                                                                                     else "Not In Employment"
                                                                            | "Additional Driver" -> if additionalCollection.[a].BusinessTypeDescription <> ""
                                                                                                     then additionalCollection.[a].BusinessTypeDescription
                                                                                                     else "Not In Employment"
                                                                            | _                   -> "Not In Employment"
                                                           if xlsLoader.codeLookup(business, xVal.[n]) <> None
                                                           then xlsLoader.codeLookup(business, xVal.[n]).Value
                                                           else ""
                        | "<MM>" | "<YY>"               -> codeConvert(dataType, fName, xVal.[n], personCollection.[c].TestId)
                        // Dates
                        | "dd/mm/yyyy" | "DD/MM/YYYY" 
                        | "yyyy/mm/dd" | "YYYY/MM/DD"
                        | "dd-mm-yyyy" | "DD-MM-YYYY"
                        | "yyyy-mm-dd" | "YYYY-MM-DD"
                        | "01/mm/yyyy" | "yyyy-mm-01"
                        | "01/MM/YYYY" | "YYYY-MM-01"
                        | "01-mm-yyyy" | "01-MM-YYYY"
                        | "YYYY-MM-DDT00:00:00"          -> let datePart, timePart = if xVal.[n] = "YYYY-MM-DDT00:00:00"
                                                                                     then "YYYY-MM-DD", "T00:00:00"
                                                                                     else xVal.[n], ""
                                                            let dateFormat = Regex.Replace(datePart.ToLower(),"mm","MM")
                                                            let dateTime = let dateOffset = xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()
                                                                           match dataType, fName with
                                                                           | "Vehicle", "Date of Purchase"                                          -> dateTester (dateFormat, todaysDay, carUsageCollection.[b].DateOfPurchaseMonthCode, carUsageCollection.[b].DateOfPurchaseYear)
                                                                           | "Vehicle", "Not purchased yet"                                         -> match (xlsLoader.brandGroup(xlsLoader.brand)) with
                                                                                                                                                       | "CDL-FilterFree" | "OpenGI"    -> dateTester (dateFormat, todaysDay, todaysMonth, todaysYear)
                                                                                                                                                       | _                              -> ((DateTime.Parse(todaysDay + "/" + todaysMonth + "/" + todaysYear)).AddDays ((float personCollection.[b].PolicyStartDateOffset) - float 1)).ToString(dateFormat)
                                                                           | "Policy", "Cover Start Date"                                           -> ((DateTime.Parse(todaysDay + "/" + todaysMonth + "/" + todaysYear)).AddDays ((float personCollection.[b].PolicyStartDateOffset) - float 1)).ToString(dateFormat)
                                                                           | "Proposer", ("DOB" | "Date of Birth" | "since birth")                  -> dateTester (dateFormat, personCollection.[a].DateOfBirthDay, personCollection.[a].DateOfBirthMonth, personCollection.[a].DateOfBirthYear)
                                                                           | "Proposer", ("Period licence held for?" | "Licence Date Obtained")     -> let licenceYear = if personCollection.[a].LicenceDateYear.Length > 0
                                                                                                                                                                         then personCollection.[a].LicenceDateYear
                                                                                                                                                                         else (System.Convert.ToInt32(todaysYear) - System.Convert.ToInt32(personCollection.[a].LicenceHeldCode) + 1).ToString()
                                                                                                                                                       dateTester (dateFormat, personCollection.[a].LicenceDateDay, personCollection.[a].LicenceDateMonth, licenceYear)
                                                                           | "Proposer", ("Have You Passed Any Driving Qualifications" | "Date Obtained (driving qualification)")    
                                                                                                                                                    -> dateTester (dateFormat, "01", personCollection.[a].ObtainedMonth, personCollection.[a].ObtainedYear)
                                                                           | "Proposer", _                                                          -> dateTester (dateFormat, personCollection.[a].IsLivingInUkSinceDay, personCollection.[a].IsLivingInUkSinceMonthCode, personCollection.[a].IsLivingInUkSinceYear)
                                                                           | "Additional Driver", ("DOB" | "Date of Birth" | "since birth")         -> dateTester (dateFormat, additionalCollection.[a].DateOfBirthDay, additionalCollection.[a].DateOfBirthMonth, additionalCollection.[a].DateOfBirthYear)
                                                                           | "Additional Driver", ("Period licence held for?" | "Licence Date Obtained")   
                                                                                                                                                    -> let licenceYear = if additionalCollection.[a].LicenceDateYear.Length > 0
                                                                                                                                                                         then additionalCollection.[a].LicenceDateYear
                                                                                                                                                                         else (System.Convert.ToInt32(todaysYear) - System.Convert.ToInt32(additionalCollection.[a].LicenceHeldCode) + 1).ToString()
                                                                                                                                                       dateTester (dateFormat, additionalCollection.[a].LicenceDateDay, additionalCollection.[a].LicenceDateMonth, licenceYear)
                                                                           
                                                                           | "Additional Driver", _                                                 -> dateTester (dateFormat, additionalCollection.[a].IsLivingInUkSinceDay, additionalCollection.[a].IsLivingInUkSinceMonthCode, additionalCollection.[a].IsLivingInUkSinceYear)
                                                                           | "Convictions", _                                                       -> dateTester (dateFormat, "01", convictionCollection.[a].ConvictionDateMonthCode, convictionCollection.[a].ConvictionDateYear)
                                                                           | "Claims", _                                                             -> dateTester (dateFormat, "01", claimCollection.[a].ClaimDateMonthCode, claimCollection.[a].ClaimDateYear)
                                                                           | _ -> ""
                                                            dateTime + timePart

                        | "<age>"                       -> let dateOffset = xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()
                                                           let coverStart = (((DateTime.Parse(todaysDay + "/" + todaysMonth + "/" + todaysYear)).AddDays ((float personCollection.[a].PolicyStartDateOffset) - float 1)).ToString("dd/MM/yyyy")).Split('/')
                                                           let todaysYearInt = System.Convert.ToInt32(coverStart.[2])
                                                           let todaysMonthInt = System.Convert.ToInt32(coverStart.[1])
                                                           let todaysDayInt = System.Convert.ToInt32(coverStart.[0])
                                                           match dataType, dateOffset with
                                                           | "Proposer", "since birth"        -> let dateOfBirthYearInt = toInt(personCollection.[a].DateOfBirthYear)
                                                                                                 let dateOfBirthMonthInt = toInt(personCollection.[a].DateOfBirthMonth)
                                                                                                 let dateOfBirthDayInt = toInt(personCollection.[a].DateOfBirthDay)
                                                                                                 if todaysMonthInt < dateOfBirthMonthInt || ((todaysMonthInt = dateOfBirthMonthInt) && (todaysDayInt < dateOfBirthDayInt))
                                                                                                 then (todaysYearInt - (dateOfBirthYearInt + 1)).ToString()
                                                                                                 else (todaysYearInt - dateOfBirthYearInt).ToString()
                                                           | "Proposer", "or since MM/YYYY"   -> let todaysYear = toInt(todaysYear)
                                                                                                 let sinceYear = toInt(personCollection.[a].IsLivingInUkSinceYear)
                                                                                                 if todaysMonthInt < System.Convert.ToInt32(personCollection.[a].IsLivingInUkSinceMonthCode)
                                                                                                 then ((todaysYear - (sinceYear + 1)).ToString())
                                                                                                 else ((todaysYear - sinceYear).ToString())
                                                           | "Additional Driver", "since birth"  
                                                                                              -> let dateOfBirthYearInt = toInt(additionalCollection.[a].DateOfBirthYear)
                                                                                                 let dateOfBirthMonthInt = toInt(additionalCollection.[a].DateOfBirthMonth)
                                                                                                 let dateOfBirthDayInt = toInt(additionalCollection.[a].DateOfBirthDay)
                                                                                                 if todaysMonthInt < dateOfBirthMonthInt || ((todaysMonthInt = dateOfBirthMonthInt) && (todaysDayInt < dateOfBirthDayInt))
                                                                                                 then (todaysYearInt - (dateOfBirthYearInt + 1)).ToString()
                                                                                                 else (todaysYearInt - dateOfBirthYearInt).ToString()
                                                           | "Additional Driver", "or since MM/YYYY"
                                                                                              -> let todaysYear = System.Convert.ToInt32(todaysYear)
                                                                                                 let sinceYear = System.Convert.ToInt32(additionalCollection.[a].IsLivingInUkSinceYear)
                                                                                                 if todaysMonthInt < System.Convert.ToInt32(additionalCollection.[a].IsLivingInUkSinceMonthCode)
                                                                                                 then ((todaysYear - (sinceYear + 1)).ToString())
                                                                                                 else ((todaysYear - sinceYear).ToString())
                                                           | _ -> ""
                        // MISC
                        | "YYYY"    -> match dataType, fName with
                                       | "Vehicle", ("Registration Year and Letter" | "Manufacturer") -> vehicleCollection.[a].RegYear
                                       | _ -> ""
                        | _         -> (xVal.[n]).Trim()

                    xmlLoader.checkXml (xValue, xLoc.[n], xmlFile, dataSrc, xlRow, dataType)
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