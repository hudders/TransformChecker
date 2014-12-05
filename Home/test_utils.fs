module test_utils

open System
open System.IO
open Microsoft.Office.Interop

open config
open jsnLoader
open xlsLoader

let toInt (str : string) =
    System.Convert.ToInt32(str)

let stripChars text (chars:string) =
    Array.fold (fun (s:string) c -> s.Replace(c.ToString(),"")) text (chars.ToCharArray())

let todaysDate = (System.DateTime.Now.ToString("dd/MM/yyyy")).Split('/')
let todaysDay, todaysMonth, todaysYear = todaysDate.[0], todaysDate.[1], todaysDate.[2]
let passCell, failCell, resultCell = "G","H","J"

let mutable fillCell =
        (fun (xlFile : Excel.Worksheet, xlColumn : string, xlRow : int, value : string) -> xlFile.Range((xlColumn + xlRow.ToString()),(xlColumn + xlRow.ToString())).Value2 <- value)

let dateTester (dateFormat : string, day : string, month : string, year : string) =
    let date =
        if month.Length > 0 then
            if day.Length > 0 then
                day + "/" + month + "/" + year
            else
                "01" + "/" + month + "/" + year
        else
            todaysDay.ToString() + "/" + todaysMonth.ToString() + "/" + year
    let parsedDate = try DateTime.Parse date with | :? System.FormatException -> DateTime.Parse "01/01/01"
    let newDateFormat = (dateFormat.ToLower()).Replace("mm","MM")
    parsedDate.ToString(newDateFormat)

let matchToExpected(dataSrc : Excel.Worksheet, acVal : string, exVal : string, exLoc : string, xlRow : int) =
    let acVal, exVal, exLoc, eCol = acVal.Replace("–","-"), exVal.Replace("–","-"), exLoc.Replace("–","-"), cellValue(dataSrc,"E",xlRow).ToString().Replace("–","-")
    let lineNumber =
        let rec loop n =
            if eCol.Split('\n').[n] = exLoc || (eCol.Split('\n').[n]).Split('[').[0] = exLoc.Split('[').[0] then
                n
            else
                loop (n + 1)
        loop 0
    let rec DataExtractionLoop row =
        if cellValue(dataExt, "B", row) = null then
            fillCell(dataExt, "C", row, lineNumber.ToString())
            let runNumber =
                if cellValue(dataExt, "C", row).ToString() = cellValue(dataExt, "C", row - 1).ToString() then
                    let x = cellValue(dataExt, "B", row - 1).ToString()
                    (System.Convert.ToInt32(x:string) + 1).ToString()
                else
                    "1"
            let matchResult =
                if exVal = "[ANYTHING]" then
                    if acVal <> "" && acVal <> "[MISSING]" then
                        "Y"
                    else
                        "N"
                else
                    if acVal = exVal || acVal.Replace(" ","") = exVal.Replace(" ","") then //Make sure phone numbers are matched.
                        "Y"
                    else
                        "N"
            fillCell(dataExt, "B", row, runNumber)
            fillCell(dataExt, "D", row, exVal)
            fillCell(dataExt, "E", row, acVal)
            fillCell(dataExt, "F", row, matchResult)
            let rec DataMergeLoop n =
                if cellValue(dataExt, "H", n) = null || cellValue(dataExt, "H", n).ToString() = (cellValue(dataExt, "B", row).ToString()) then
                    fillCell(dataExt, "H", n, (cellValue(dataExt, "B", row).ToString()))
                    let dataFill(xCol : string, mCol : string) =
                        if cellValue(dataExt, mCol, n) = null then
                            cellValue(dataExt, xCol, row).ToString()
                        else
                            cellValue(dataExt, mCol, n).ToString() + "\n" + cellValue(dataExt, xCol, row).ToString()
                    fillCell(dataExt, "I", n, dataFill("E","I"))
                    fillCell(dataExt, "J", n, dataFill("F","J"))
                    fillCell(dataExt, "L", n, dataFill("D","L"))
                else
                    DataMergeLoop (n + 1)
                if contains "N" ((cellValue(dataExt, "J", n).ToString()).Split('\n')) then
                    fillCell(dataExt, "K", n, "Fail")
                else
                    fillCell(dataExt, "K", n, "Pass")
            DataMergeLoop startCell
        else
            DataExtractionLoop (row + 1)
    DataExtractionLoop startCell

let cleanUp() =
    let rec cleanMeLoop n =
        fillCell(dataExt, "B", n, null)
        fillCell(dataExt, "C", n, null)
        fillCell(dataExt, "D", n, null)
        fillCell(dataExt, "E", n, null)
        fillCell(dataExt, "F", n, null)
        fillCell(dataExt, "H", n, null)
        fillCell(dataExt, "I", n, null)
        fillCell(dataExt, "J", n, null)
        fillCell(dataExt, "K", n, null)
        fillCell(dataExt, "L", n, null)
        if cellValue(dataExt, "B", n + 1) <> null then
            cleanMeLoop (n + 1)
    cleanMeLoop startCell

let deleteExt(path : string, extension : string) =
    for file in System.IO.Directory.EnumerateFiles(path + "/", "*." + extension) do
        let tempPath = System.IO.Path.Combine(path, file)
        //printf "%s" tempPath
        System.IO.File.Delete(tempPath)

let convertMapping (valueType : string, dataType : string, fName : string, journeyNumber : string, risk : System.Collections.Generic.List<Risk>) =
    let j = toInt(journeyNumber)-1
    match valueType with
    | "as input"        -> match fName with
                           | "FirstName" | "First Name" | "Forename"                  -> match dataType with
                                                                                         | "Proposer" -> risk.[j].firstName
                                                                                         | "Joint Proposer" -> risk.[j].jp_firstName
                                                                                         | _ -> ""
                           | "Lastname" | "Surname" | "Last Name"                     -> match dataType with
                                                                                         | "Proposer" -> risk.[j].lastName
                                                                                         | "Joint Proposer" -> risk.[j].jp_lastName
                                                                                         | _ -> ""
                           | "House Number / Name"                                    -> match dataType with
                                                                                         | "Proposer" -> if risk.[j].postalAddress_number <> ""
                                                                                                         then risk.[j].postalAddress_number
                                                                                                         else risk.[j].postalAddress_building
                                                                                         | "Property" -> if risk.[j].riskAddress_number <> ""
                                                                                                         then risk.[j].riskAddress_number
                                                                                                         else risk.[j].riskAddress_building
                                                                                         | _ -> ""
                           | "Address Line 1"                                         -> match dataType with
                                                                                         | "Proposer" -> risk.[j].postalAddress_lineOne
                                                                                         | "Property" -> risk.[j].riskAddress_lineOne
                                                                                         | _ -> ""
                           | "Address Line 2"                                         -> match dataType with
                                                                                         | "Proposer" -> risk.[j].postalAddress_thoroughfare
                                                                                         | "Property" -> risk.[j].riskAddress_thoroughfare
                                                                                         | _ -> ""
                           | "Address Line 3"                                         -> match dataType with
                                                                                         | "Proposer" -> risk.[j].postalAddress_dependentLocality
                                                                                         | "Property" -> risk.[j].riskAddress_dependentLocality
                                                                                         | _ -> ""
                           | "Address Line 4"                                         -> match dataType with
                                                                                         | "Proposer" -> risk.[j].postalAddress_town
                                                                                         | "Property" -> risk.[j].riskAddress_town
                                                                                         | _ -> ""
                           | "Address Line 5"                                         -> match dataType with
                                                                                         | "Proposer" -> if risk.[j].postalAddress_postalCounty <> ""
                                                                                                         then risk.[j].postalAddress_postalCounty
                                                                                                         elif risk.[j].postalAddress_optionalCounty <> ""
                                                                                                         then risk.[j].postalAddress_optionalCounty
                                                                                                         elif risk.[j].postalAddress_traditionalCounty <> ""
                                                                                                         then risk.[j].postalAddress_traditionalCounty
                                                                                                         else risk.[j].postalAddress_administrativeCounty
                                                                                         | "Property" -> if risk.[j].riskAddress_postalCounty <> ""
                                                                                                         then risk.[j].riskAddress_postalCounty
                                                                                                         elif risk.[j].riskAddress_optionalCounty <> ""
                                                                                                         then risk.[j].riskAddress_optionalCounty
                                                                                                         elif risk.[j].riskAddress_traditionalCounty <> ""
                                                                                                         then risk.[j].riskAddress_traditionalCounty
                                                                                                         else risk.[j].riskAddress_administrativeCounty
                                                                                         | _ -> ""
                           | "Address Line 6"                                         -> match dataType with
                                                                                         | "Proposer" -> risk.[j].postalAddress_administrativeCounty
                                                                                         | _          -> risk.[j].riskAddress_administrativeCounty
                           | "Postcode"                                               -> match dataType with
                                                                                         | "Proposer" -> risk.[j].postalAddress_postcode
                                                                                         | "Property" -> risk.[j].riskAddress_postcode
                                                                                         | _ -> ""
                           | "Main Telephone Number"                                  -> risk.[j].telephoneNumber
                           | "Email" | "E-mail"                                        -> risk.[j].emailAddress
                           | "How many children are normally minded at this property"  -> risk.[j].childrenMindedAtProperty.ToString()
                           | "What amount of total cover do you require?"              -> risk.[j].contents_awayFromHome.ToString()
                           | "Total value of the contents to be insured"               -> risk.[j].contents_sumInsured.ToString()
                           | "Total value of the high risk items of your home"         -> risk.[j].contents_highRisk.ToString()
                           | "Value of the most expensive high risk item in your home" -> risk.[j].contents_mostExpensive.ToString()
                           | "What is the rebuild cost of the property?"               -> risk.[j].buildings_sumInsured.ToString()
                           | "Days consecutively left unoccupied"                      -> risk.[j].unoccupiedDays
                           | "Item Value £" | "Item Description" | "Cost (best estimate)" | "Bicycle Value £" | "Bicycle Description"
                                                                                       -> match dataType with
                                                                                          | "Contents Cover" -> let rec itemLoop itemID =
                                                                                                                    if itemCollection.[itemID].riskID = risk.[j].testID
                                                                                                                    then match fName, itemCollection.[itemID].itemType with
                                                                                                                         | "Bicycle Description", "Bicycle" | "Item Description", _ -> itemCollection.[itemID].itemDesc
                                                                                                                         | "Bicycle Value £", "Bicycle" | "Item Value £", _         -> itemCollection.[itemID].itemValue
                                                                                                                         | _ -> if itemCollection.Count > itemID + 1
                                                                                                                                then itemLoop (itemID + 1)
                                                                                                                                else ""
                                                                                                                    elif itemCollection.Count > itemID + 1
                                                                                                                    then itemLoop (itemID + 1)
                                                                                                                    else ""
                                                                                                                itemLoop 0
                                                                                          | "Claims"          -> let rec claimLoop claimID =
                                                                                                                     if claimCollection.[claimID].riskID = risk.[j].testID
                                                                                                                     then claimCollection.[claimID].claim_cost.ToString()
                                                                                                                     elif claimCollection.Count > claimID + 1
                                                                                                                     then claimLoop (claimID + 1)
                                                                                                                     else ""
                                                                                                                 claimLoop 0
                                                                                          | _ -> ""

                           | _ -> ""
    | "<conCode>" | "<carCode>" | "<styleCode>" | "<occCode>" | "<empCode>"            -> let variableName =
                                                                                              if dataType = "Proposer" && valueType = "<occCode>" 
                                                                                              then risk.[j].occupationCode
                                                                                              elif dataType = "Proposer" && valueType = "<empCode>"
                                                                                              then risk.[j].businessType
                                                                                              elif dataType = "Joint Proposer" && valueType = "<occCode>"
                                                                                              then risk.[j].jp_occupationCode
                                                                                              elif dataType = "Joint Proposer" && valueType = "<empCode>"
                                                                                              then risk.[j].jp_businessType
                                                                                              else ""
                                                                                          if xlsLoader.codeLookup(variableName, valueType) <> None
                                                                                          then xlsLoader.codeLookup(variableName, valueType).Value
                                                                                          else ""
    | "dd/mm/yyyy" | "DD/MM/YYYY" | "yyyy/mm/dd" | "YYYY/MM/DD" | "dd-mm-yyyy"
    | "DD-MM-YYYY" | "yyyy-mm-dd" | "YYYY-MM-DD" | "01/mm/yyyy" | "yyyy-mm-01"
    | "01/MM/YYYY" | "YYYY-MM-01" | "01-mm-yyyy" | "01-MM-YYYY" | "yyyy/mm/01"
    | "YYYY/MM/01" | "yyyy/01/01" | "YYYY/01/01" | "YYYY-MM-DDT00:00:00"               -> if dataType = "Claim" || dataType = "Claims"
                                                                                          then let rec claimLoop claimID =
                                                                                                   if claimCollection.[claimID].riskID = risk.[j].testID
                                                                                                   then dateTester (valueType, (claimCollection.[claimID].claim_date.Day).ToString(), (claimCollection.[claimID].claim_date.Month).ToString(), (claimCollection.[claimID].claim_date.Year).ToString())
                                                                                                   elif claimCollection.Count > claimID
                                                                                                   then claimLoop (claimID + 1)
                                                                                                   else ""
                                                                                               claimLoop 0
                                                                                          elif dataType = "Property"
                                                                                          then dateTester (valueType, "01", "01", risk.[j].dateRoofRecovered)
                                                                                          //else let dateOffset = xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()
                                                                                          else if fName = "DOB" || fName = "Date of Birth" //|| dateOffset = "since birth"
                                                                                               then if dataType = "Joint Proposer"
                                                                                                    then dateTester (valueType, risk.[j].jp_dateOfBirthDay, risk.[j].jp_dateOfBirthMonth, risk.[j].jp_dateOfBirthYear)
                                                                                                    else dateTester (valueType, risk.[j].dateOfBirthDay, risk.[j].dateOfBirthMonth, risk.[j].dateOfBirthYear)
                                                                                               else dateTester (valueType, todaysDay, todaysMonth, todaysYear)

    | "YYYY"                                                                           -> risk.[j].yearBuilt
    | _                                                                                -> valueType