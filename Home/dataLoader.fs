module dataLoader

open config
open test_utils
open jsnLoader

let rowList = [3..125]

let loadData (product : string, dataType : string, xmlFile : string, a : int) =
    let dataSrc =
        if dataType = "Proposer" then
            xlsLoader.getXLS("2.1 Proposer");
        elif dataType = "Additional Driver" || dataType = "Joint Proposer" then
            xlsLoader.getXLS("2.2 " + dataType)
        elif (dataType = "Claim" && product = "PC") || dataType = "Property" then
            xlsLoader.getXLS("2.3 " + dataType)
        elif dataType = "Conviction" then
            xlsLoader.getXLS("2.4 Convictions")
        elif dataType = "Contents"then
            xlsLoader.getXLS("2.4 Contents Cover")
        elif dataType = "Vehicle" then
            xlsLoader.getXLS("2.5 Vehicle")
        elif dataType = "Buildings" then
            xlsLoader.getXLS("2.5 Buildings Cover")
        elif dataType = "Policy" || dataType = "Locks and Security" then
            xlsLoader.getXLS("2.6 " + dataType)
        elif dataType = "Claim" && product = "HH" then
            xlsLoader.getXLS("2.7 Claims")
        elif dataType = "Price" then
            xlsLoader.getXLS("2.8 Price Page")
        else
            xlsLoader.getXLS("")
    for xlRow in rowList do
        if xlsLoader.cellValue(dataSrc, "A", xlRow) <> null && xlsLoader.cellValue(dataSrc, "A", xlRow).ToString() = (risk.[a].testID).ToString() then
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
                            if fName = "FirstName" || fName = "First Name" || fName = "Forename" then
                                if dataType = "Joint Proposer" then
                                    risk.[a].jp_firstName
                                else
                                    risk.[a].firstName
                            elif fName = "Lastname" || fName = "Surname" then
                                if dataType = "Joint Proposer" then
                                    risk.[a].jp_lastName
                                else
                                    risk.[a].lastName
                            elif fName = "Address" || fName = "Postcode" || fName = "House Number / Name" then
                                let addressParts = (xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()).Split('\n')
                                if dataType = "Proposer" then
                                    if addressParts.[n] = "Address Line 1" || fName = "House Number / Name" then
                                        if risk.[a].postalAddress_number = "" then
                                            risk.[a].postalAddress_building
                                        else
                                            risk.[a].postalAddress_number
                                    elif addressParts.[n] = "Address Line 2" then
                                        risk.[a].postalAddress_thoroughfare
                                    elif addressParts.[n] = "Address Line 3" then
                                        risk.[a].postalAddress_dependentLocality
                                    elif addressParts.[n] = "Address Line 4" then
                                        risk.[a].postalAddress_town
                                    elif addressParts.[n] = "Address Line 5" then
                                        risk.[a].postalAddress_postalCounty
                                    elif addressParts.[n] = "Postcode" || fName = "Postcode" then
                                        risk.[a].postalAddress_postcode
                                    else
                                        ""
                                else
                                    if addressParts.[n] = "Address Line 1" || fName = "House Number / Name" then
                                        if risk.[a].riskAddress_number = "" then
                                            risk.[a].riskAddress_building
                                        else
                                            risk.[a].riskAddress_number
                                    elif addressParts.[n] = "Address Line 2" then
                                        risk.[a].riskAddress_thoroughfare
                                    elif addressParts.[n] = "Address Line 3" then
                                        risk.[a].riskAddress_dependentLocality
                                    elif addressParts.[n] = "Address Line 4" then
                                        risk.[a].riskAddress_town
                                    elif addressParts.[n] = "Address Line 5" then
                                        risk.[a].riskAddress_postalCounty
                                    elif addressParts.[n] = "Postcode" || fName = "Postcode" then
                                        risk.[a].riskAddress_postcode
                                    else
                                        ""
                            elif fName = "Total value of the contents to be insured" then
                                risk.[a].contents_sumInsured.ToString()
                            elif fName = "Total value of the high risk items of your home" then
                                risk.[a].contents_highRisk.ToString()
                            elif fName = "Value of the most expensive high risk item in your home" then
                                risk.[a].contents_mostExpensive.ToString()
                            elif fName = "What is the rebuild cost of the property?" then
                                risk.[a].buildings_sumInsured.ToString()
                            elif dataType = "Contents" && fName = "Description" then
                                let rec itemLoop itemID =
                                    if itemCollection.[itemID].riskID = risk.[a].testID then
                                        itemCollection.[itemID].itemDesc
                                    elif itemCollection.Count > itemID then
                                        itemLoop (itemID + 1)
                                    else
                                        ""
                                itemLoop 0
                            elif dataType = "Contents" && fName = "Value £" then
                                let rec itemLoop itemID =
                                    if itemCollection.[itemID].riskID = risk.[a].testID then
                                        itemCollection.[itemID].itemValue
                                    elif itemCollection.Count > itemID then
                                        itemLoop (itemID + 1)
                                    else
                                        ""
                                itemLoop 0
                             elif dataType = "Claim" && fName = "Cost (best estimate)" then
                                let rec claimLoop claimID =
                                    if claimCollection.[claimID].riskID = risk.[a].testID then
                                        (claimCollection.[claimID].claim_cost).ToString()
                                    elif claimCollection.Count > claimID then
                                        claimLoop (claimID + 1)
                                    else
                                        ""
                                claimLoop 0
                            else
                                ""
                        elif contains xVal.[n] ["<conCode>"; "<carCode>"; "<styleCode>"; "<occCode>"; "<empCode>"] then
                            "Aaron"
                        elif contains xVal.[n] ["YYYY-MM-DD"; "YYYY-MM-01"; "DD-MM-YYYY"; "DD MM YYYY"; "YYYY/MM/DD"] then
                            if dataType = "Claim" then
                                let rec claimLoop claimID =
                                    if claimCollection.[claimID].riskID = risk.[a].testID then
                                        dateTester (xVal.[n], (claimCollection.[claimID].claim_date.Day).ToString(), (claimCollection.[claimID].claim_date.Month).ToString(), (claimCollection.[claimID].claim_date.Year).ToString())
                                    elif claimCollection.Count > claimID then
                                        claimLoop (claimID + 1)
                                    else
                                        ""
                                claimLoop 0
                            else
                                let dateOffset = xlsLoader.cellValue(dataSrc, "C", xlRow).ToString()
                                if fName = "DOB" || fName = "Date of Birth" || dateOffset = "since birth" then
                                    if dataType = "Joint Proposer" then
                                        dateTester (xVal.[n], risk.[a].jp_dateOfBirthDay, risk.[a].jp_dateOfBirthMonth, risk.[a].jp_dateOfBirthYear)
                                    else
                                        dateTester (xVal.[n], risk.[a].dateOfBirthDay, risk.[a].dateOfBirthMonth, risk.[a].dateOfBirthYear)
                                else
                                    ""
                        elif xVal.[n] = "<age>" then
                            "Aaron" //TODO
                        elif xVal.[n] = "YYYY" then
                            "Aaron" //TODO
                        else
                            xVal.[n]
                    xmlLoader.checkXml (xValue, xLoc.[n], xmlFile, dataSrc, xlRow)
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
                ResultLoop startCell
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