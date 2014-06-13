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
    let xmlString = xmlString.Replace("xmlns=","x=")
    if brand = "Drivology" then
        xmlString.Replace("pmq:","")
    else
        xmlString

// getXML
// Downloads the appropriate XML files and returns a linkID
// so we can find the files later.
let getXML(environment, brand, lastName, email, vehicleId) =
    start chrome
    let environmentURL = match environment with
                            | "UAT" -> "http://peg-ctmcoruqt01.comparethemarket.local:65000#"
                            | "REG" -> "http://peg-ctmcorrqt01.comparethemarket.local:65000#"
                            | _ -> "http://peg-ctmcorqqt01.comparethemarket.local:65000#"
    let productClass = match vehicleId with
                            | Prefix "PC" rest -> "PrivateCar"
                            | Prefix "LC" rest -> "LightCommercial"
                            | _ -> "Household"
    url environmentURL
    "#DropDownListProducts" << productClass
    "#DropDownListBrands" << brand
    "#TextBoxSurname" << lastName
    "#TextBoxEmailAdr" << email
    click "Search"

    let elementPath = "#DataListFiles tbody tr td a:link"

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

    quit()
    linkID

let checkXml(expectedVal : string, expectedLoc : string, xmlFile : string, xlsFile : Excel.Worksheet, xlsNode : int) =
    if File.Exists(xmlFile) then
        let xml = XDocument.Load(xmlFile).ToString()
        let doc = new XmlDocument() in doc.LoadXml (processXml(xml))
        doc.SelectNodes expectedLoc
            |> Seq.cast<XmlNode>
            |> Seq.iter (fun node -> matchToExpected(xlsFile, node.InnerXml, expectedVal, expectedLoc, xlsNode))
        

    else
        printfn "Skipped test because xml file doesn't exist - check your filters!"