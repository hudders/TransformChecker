module xmlLoader

open canopy
open System.IO
open System.Xml
open System.Xml.Linq
open System.Xml.XPath
open Microsoft.Office.Interop

open config
open csvLoader
open xlsLoader
open test_utils

// processXml
// Sorts out any problems with the xml on a brand basis
let processXml (xmlString : string) =
    let xmlString = xmlString.Replace("xmlns=","x=").Replace("tem:","").Replace("web:","").Replace("params:","").Replace("&lt;","<").Replace("&gt;",">")
    if brandGroup(brand) = "SSP Multi Quote" then
        xmlString.Replace("pmq:","")
    elif brandGroup(brand) = "QuoteExchange" then
        xmlString.Replace("a:","")
    else
        xmlString

// hasAttribute
// Determine whether value contains an attribute declaration
let hasAttribute (value : string) =
    match value with
        | RegExParse "([\s\S]*\[@[\s\S]*=[\s\S]*\])" value -> true
        | _ -> false

// getXML
// Downloads the appropriate XML files and returns a linkID
// so we can find the files later.
let getXML(environment, product, brand : string, lastName, email, xmlFile) =
    let xmlFolder = xlsRtFolder + "\Car\\" + brandX + "\xml"
    start chrome
    let browser1 = browser
    let environmentURL = match environment with
                            | "UAT" -> "http://peg-ctmcoruqt01.comparethemarket.local:65000#"
                            | "REG" -> "http://peg-ctmcorrqt01.comparethemarket.local:65000#"
                            | _ -> "http://peg-ctmcorqqt01.comparethemarket.local:65000#"
    let productClass = match product with
                            | Prefix "PC" rest -> "PrivateCar"
                            | Prefix "LC" rest -> "LightCommercial"
                            | _ -> "Household"
    url environmentURL
    "#DropDownListProducts" << productClass
    //"#DropDownListBrands" << brand
    click brandCode
    "#TextBoxSurname" << lastName
    "#TextBoxEmailAdr" << email
    click "Search"

    let elementPath = "#DataListFiles tbody tr td a:link"

    let linkID =
        if (elements elementPath).Length > 0 then

            let listOfXmlFileLinks = 
                seq { for x in elements elementPath -> System.DateTime.Parse(x.GetAttribute("text").[..18]) }
                |> List.ofSeq
                |> List.sortBy(fun x -> x)
                |> List.rev

            let xmlLink =
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
            linkID
        else
            printfn "No XML found in QuoteFinder. Skipping tests."
            null
    let srcFile = ((dloadFolder + "Request_" + linkID.[1] + "_After_" + linkID.[3] + "_2.xml").ToString())
//    let errFile = ((dloadFolder + "Error_" + brandGroup(brand) + "_After_" + linkID.[3] + "_3.xml").ToString())
    if linkID <> null && File.Exists(srcFile) then
        printfn "Successfully downloaded XML"
        printfn "%s" srcFile
        File.Copy(srcFile, xmlFile)
//    elif File.Exists(errFile) then
//        printfn "XML downloaded but contains errors:"
//        let xml = XDocument.Load(errFile).ToString()
//        let doc = new XmlDocument() in doc.LoadXml (processXml(xml))
//        let xSeq = doc.SelectNodes "//Messages/GeneralMessages/Message/MessageText" |> Seq.cast<XmlNode>
//        if Seq.isEmpty xSeq = false then
//            printfn ""
//            xSeq |> Seq.iter (fun node -> printfn "%s" node.InnerXml)
//            printfn ""
    quit browser1
    printfn "Timestamp: %s" (System.DateTime.Now.ToString("hh:mm:ss"))

let checkXml(expectedVal : string, expectedLoc : string, xmlFile : string, xlsFile : Excel.Worksheet, xlsNode : int, dataType : string) =
    if File.Exists(xmlFile) then
        let f, j = xlsLoader.cellValue(xlsFile, "B", xlsNode).ToString(), xlsLoader.cellValue(xlsFile, "A", xlsNode).ToString()
        //let expectedVal = codeConvert(dataType, f, expectedVal, j)
        
        let xml = XDocument.Load(xmlFile).ToString()
        let doc = new XmlDocument() in doc.LoadXml (processXml(xml))
        
        let xSeq(location : string) = 
            doc.SelectNodes location
                |> Seq.cast<XmlNode>

        let expectedLoc, expectedVal =
            if hasAttribute expectedVal then
                // value[@attribute_name="attribute_val]"
                let value = expectedVal.Split('[').[0]
                let attribute_name = (expectedVal.Split('[').[1]).Split('"').[0]
                let attribute_val = codeConvert(dataType,f, (expectedVal.Split('[').[1]).Split('"').[1], j)
                expectedLoc + "[" + attribute_name + "\"" + attribute_val + "\"]", value
            else
                expectedLoc, expectedVal

        printfn "%s %s" expectedLoc expectedVal

        if Seq.isEmpty (xSeq(expectedLoc)) then
            matchToExpected(xlsFile, "[MISSING]", (if expectedVal = "" then "[EMPTY]" else expectedVal), expectedLoc, xlsNode)
        else
            xSeq(expectedLoc)
                |> Seq.iter (fun node -> matchToExpected(xlsFile, (if node.InnerXml = "" then "[EMPTY]" else node.InnerXml), (if expectedVal = "" then "[EMPTY]" else expectedVal), expectedLoc, xlsNode))
    else
        printfn "Skipped test because xml file doesn't exist - check your filters!"