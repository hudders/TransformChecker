module xmlLoader

open canopy
open System.IO
open System.Xml
open System.Xml.Linq
open System.Xml.XPath
open Microsoft.Office.Interop

open config
open test_utils
open xlsLoader

// processXml
// Sorts out any problems with the xml on a brand basis
let processXml (xmlString : string) =
    let xmlString = xmlString.Replace("xmlns=","x=").Replace("tem:","").Replace("web:","").Replace("&lt;","<").Replace("&gt;",">")
    if brandGroup(brand) = "SSP Multi Quote" then
        xmlString.Replace("pmq:","")
    elif brandGroup(brand) = "QuoteExchange" then
        xmlString.Replace("a:","")
    else
        xmlString

// getXML
// Downloads the appropriate XML files and returns a linkID
// so we can find the files later.
let getXML(environment, product, brand : string, lastName, email, xmlFile) =
    if product <> "HH" || brandGroup(brand) <> "CDL-FilterFree" then
        start chrome
        let browser1 = browser
        let environmentURL = match environment with
                                | "UAT" -> "http://peg-ctmcoruqt02.comparethemarket.local:65000#"
                                | "REG" -> "http://peg-ctmcorrqt01.comparethemarket.local:65000#"
                                | "UAT_BLUE" -> "http://peg-ctmhoruwb01.comparethemarket.local:65000/"
                                | _ -> "http://peg-ctmcorqqt01.comparethemarket.local:65000#"
        let productClass, xmlFolder = match product with
                                        | Prefix "PC" rest -> "PrivateCar", xlsRtFolder + "\Car\\" + brandX + "\xml"
                                        | Prefix "LC" rest -> "LightCommercial", xlsRtFolder + "\Van\\" + brandX + "\xml"
                                        | Prefix "HH" rest -> "Household_Blue", xlsRtFolder + "\Home\\" + brandX + "\xml"
                                        | _ -> "Household", ""
        url environmentURL
        "#DropDownListProducts" << productClass
        click "#DropDownListProducts"
        click brandCode
        "#TextBoxSurname" << lastName
        "#TextBoxEmailAdr" << email
        click "Search"

        let elementPath = "#DataListFiles tbody tr td a:link"

        let linkID =
            if (elements elementPath).Length > 0 then

                let listOfXmlFileLinks = // Creates a list of all datetime stamps, orders by age, newest to oldest
                    seq { for x in elements elementPath -> System.DateTime.Parse(x.GetAttribute("text").[..18]) }
                    |> List.ofSeq
                    |> List.sortBy(fun x -> x)
                    |> List.rev

                let xmlLink = // Iterates through links on page, finds the link with the newest date
                    let result = ref None in
                    seq { for x in elements elementPath -> x.GetAttribute("text") }
                    |> Seq.iter (fun x -> if x.StartsWith(listOfXmlFileLinks.[0].ToString()) then result := Some x)
                    !result

                click xmlLink.Value

                let linkID = ((element "#GridView1 tbody tr td a:link").GetAttribute("href")).Split('_')

                let clickAll selector =
                    elements selector
                    |> List.iter (fun element -> click element)

                clickAll "#GridView1 input"
                printfn "Successfully downloaded XML"

                linkID
            else
                printfn "No XML found in QuoteFinder. Skipping tests."
                null
        if linkID <> null then // move the request after to the brand's xml folder in panel.testplans and rename it something sensible
            File.Copy((dloadFolder + "Request_" + brandGroup(brand) + "_After_" + linkID.[3] + "_2.xml").ToString(), xmlFile)
            deleteExt(dloadFolder, "xml")
            deleteExt(dloadFolder, "tmp")
        quit browser1
        printfn "Timestamp: %s" (System.DateTime.Now.ToString("hh:mm:ss"))

let checkXml(expectedVal : string, expectedLoc : string, xmlFile : string, xlsFile : Excel.Worksheet, xlsNode : int) =
    if File.Exists(xmlFile) then
        let xml = XDocument.Load(xmlFile).ToString()
        let doc = new XmlDocument() in doc.LoadXml (processXml(xml))
        
        let xSeq = doc.SelectNodes expectedLoc
                    |> Seq.cast<XmlNode>
        
        if Seq.isEmpty xSeq then
            matchToExpected(xlsFile, "[MISSING]", expectedVal, expectedLoc, xlsNode)
        else
            xSeq
                |> Seq.iter (fun node -> matchToExpected(xlsFile, (if node.InnerXml = "" then "[EMPTY]" else node.InnerXml), expectedVal, expectedLoc, xlsNode))
                //|> Seq.iter (fun node -> printfn "%s - %s" node.InnerXml expectedVal)
    else
        printfn "Skipped test because xml file doesn't exist - check your filters!"