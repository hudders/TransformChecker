module jsnLoader

open System
open System.Collections.Generic
open FSharp.Data
open FSharp.Data.Json.Extensions

open config

type riskMaster = JsonProvider< "ref\json\example\01_example.json" >
let fileArray = System.IO.Directory.GetFiles(jsnRtFolder, "*.json")
let numberOfFiles = fileArray.Length

// RISK //
type Risk() =
    class
        [<DefaultValue>] val mutable testID : int
        // Proposer
        [<DefaultValue>] val mutable firstName : string
        [<DefaultValue>] val mutable lastName : string
        [<DefaultValue>] val mutable dateOfBirthDay : string
        [<DefaultValue>] val mutable dateOfBirthMonth : string
        [<DefaultValue>] val mutable dateOfBirthYear : string
        [<DefaultValue>] val mutable postalAddress_organisationName : string
        [<DefaultValue>] val mutable postalAddress_department : string
        [<DefaultValue>] val mutable postalAddress_subBuilding : string
        [<DefaultValue>] val mutable postalAddress_building : string
        [<DefaultValue>] val mutable postalAddress_number : string
        [<DefaultValue>] val mutable postalAddress_dependentThoroughfare : string
        [<DefaultValue>] val mutable postalAddress_thoroughfare : string
        [<DefaultValue>] val mutable postalAddress_doubleDependentLocality : string
        [<DefaultValue>] val mutable postalAddress_dependentLocality : string
        [<DefaultValue>] val mutable postalAddress_town : string
        [<DefaultValue>] val mutable postalAddress_traditionalCounty : string
        [<DefaultValue>] val mutable postalAddress_administrativeCounty : string
        [<DefaultValue>] val mutable postalAddress_optionalCounty : string
        [<DefaultValue>] val mutable postalAddress_postalCounty : string
        [<DefaultValue>] val mutable postalAddress_postcode : string
        [<DefaultValue>] val mutable telephoneNumber : int
        [<DefaultValue>] val mutable emailAddress : string
        [<DefaultValue>] val mutable occupationCode : string
        [<DefaultValue>] val mutable businessType : string
        [<DefaultValue>] val mutable childrenMindedAtProperty : int
        // Joint Proposer
        [<DefaultValue>] val mutable jp_firstName : string
        [<DefaultValue>] val mutable jp_lastName : string
        [<DefaultValue>] val mutable jp_dateOfBirthDay : string
        [<DefaultValue>] val mutable jp_dateOfBirthMonth : string
        [<DefaultValue>] val mutable jp_dateOfBirthYear : string
        [<DefaultValue>] val mutable jp_occupationCode : string
        [<DefaultValue>] val mutable jp_businessType : string
        // Property
        [<DefaultValue>] val mutable riskAddress_organisationName : string
        [<DefaultValue>] val mutable riskAddress_department : string
        [<DefaultValue>] val mutable riskAddress_subBuilding : string
        [<DefaultValue>] val mutable riskAddress_building : string
        [<DefaultValue>] val mutable riskAddress_number : string
        [<DefaultValue>] val mutable riskAddress_dependentThoroughfare : string
        [<DefaultValue>] val mutable riskAddress_thoroughfare : string
        [<DefaultValue>] val mutable riskAddress_doubleDependentLocality : string
        [<DefaultValue>] val mutable riskAddress_dependentLocality : string
        [<DefaultValue>] val mutable riskAddress_town : string
        [<DefaultValue>] val mutable riskAddress_traditionalCounty : string
        [<DefaultValue>] val mutable riskAddress_administrativeCounty : string
        [<DefaultValue>] val mutable riskAddress_optionalCounty : string
        [<DefaultValue>] val mutable riskAddress_postalCounty : string
        [<DefaultValue>] val mutable riskAddress_postcode : string
        // Contents Cover
        [<DefaultValue>] val mutable contents_sumInsured : int
        [<DefaultValue>] val mutable contents_highRisk : int
        [<DefaultValue>] val mutable contents_mostExpensive : int
        // Buildings Cover
        [<DefaultValue>] val mutable buildings_sumInsured : int
        // Locks & Security
        // Claims
        // Price Page
    end

let mutable risk = new List<Risk>()

let riskLoad() =
    let rec loop n =
        let jsonFile = riskMaster.Load(fileArray.[n-1])
        let obj = new Risk()
        obj.testID <- n
        obj.firstName <- jsonFile.PolicyHolder.FirstName
        obj.lastName <- jsonFile.PolicyHolder.Surname
        obj.dateOfBirthDay <- (DateTime.Parse((jsonFile.PolicyHolder.DateOfBirth).ToString()).Day).ToString()
        obj.dateOfBirthMonth <- (DateTime.Parse((jsonFile.PolicyHolder.DateOfBirth).ToString()).Month).ToString()
        obj.dateOfBirthYear <- (DateTime.Parse((jsonFile.PolicyHolder.DateOfBirth).ToString()).Year).ToString()
        obj.postalAddress_organisationName <- jsonFile.ContactDetails.PostalAddress.OrganisationName
        obj.postalAddress_department <- jsonFile.ContactDetails.PostalAddress.Department
        obj.postalAddress_subBuilding <- jsonFile.ContactDetails.PostalAddress.SubBuilding
        obj.postalAddress_building <- jsonFile.ContactDetails.PostalAddress.Building
        obj.postalAddress_number <- jsonFile.ContactDetails.PostalAddress.Number
        obj.postalAddress_dependentThoroughfare <- jsonFile.ContactDetails.PostalAddress.DependentThoroughfare
        obj.postalAddress_thoroughfare <- jsonFile.ContactDetails.PostalAddress.Thoroughfare
        obj.postalAddress_doubleDependentLocality <- jsonFile.ContactDetails.PostalAddress.DoubleDependentLocality
        obj.postalAddress_dependentLocality <- jsonFile.ContactDetails.PostalAddress.DependentLocality
        obj.postalAddress_town <- jsonFile.ContactDetails.PostalAddress.Town
        obj.postalAddress_traditionalCounty <- jsonFile.ContactDetails.PostalAddress.TraditionalCounty
        obj.postalAddress_administrativeCounty <- jsonFile.ContactDetails.PostalAddress.AdministrativeCounty
        obj.postalAddress_optionalCounty <- jsonFile.ContactDetails.PostalAddress.OptionalCounty
        obj.postalAddress_postalCounty <- jsonFile.ContactDetails.PostalAddress.PostalCounty
        obj.postalAddress_postcode <- jsonFile.ContactDetails.PostalAddress.Postcode
        obj.telephoneNumber <- jsonFile.ContactDetails.Telephone
        obj.emailAddress <- jsonFile.ContactDetails.Email
        // Annoyingly, the json doesn't include occupation and business codes for the "not in employment" types, so we have to add them:
        let occCode, busCode =
            if jsonFile.PolicyHolder.EmploymentStatus = "retired" then "R09", "947"
            elif jsonFile.PolicyHolder.EmploymentStatus = "houseperson" then "H09", "948"
            elif jsonFile.PolicyHolder.EmploymentStatus = "unemployed" then "U03", "747"
            elif jsonFile.PolicyHolder.EmploymentStatus = "unableToWork" then "42D", "949"
            else (jsonFile.PolicyHolder.OccupationCode).ToString(), (jsonFile.PolicyHolder.BusinessCode).ToString()
        obj.occupationCode <- occCode
        obj.businessType <- busCode
        if (jsonFile.Property.Residents.JsonValue.ToString()).Contains("childMindingNumberOfChildren") && jsonFile.Property.Residents.ChildMinding.ChildMindingProperty then
            obj.childrenMindedAtProperty <- jsonFile.Property.Residents.ChildMinding.ChildMindingNumberOfChildren
        else
            obj.childrenMindedAtProperty <- 0
        // Joint Proposer is not mandatory. Here we prevent looking for their details if they don't exist.
        if (jsonFile.JsonValue.TryGetProperty("jointPolicyHolder")).IsSome then
            obj.jp_firstName <- jsonFile.JointPolicyHolder.FirstName
            obj.jp_lastName <- jsonFile.JointPolicyHolder.Surname
            obj.jp_dateOfBirthDay <- (DateTime.Parse((jsonFile.JointPolicyHolder.DateOfBirth).ToString()).Day).ToString()
            obj.jp_dateOfBirthMonth <- (DateTime.Parse((jsonFile.JointPolicyHolder.DateOfBirth).ToString()).Month).ToString()
            obj.jp_dateOfBirthYear <- (DateTime.Parse((jsonFile.JointPolicyHolder.DateOfBirth).ToString()).Year).ToString()
            let occCode, busCode =
                if jsonFile.JointPolicyHolder.EmploymentStatus = "retired" then "R09", "947"
                elif jsonFile.JointPolicyHolder.EmploymentStatus = "houseperson" then "H09", "948"
                elif jsonFile.JointPolicyHolder.EmploymentStatus = "unemployed" then "U03", "747"
                elif jsonFile.JointPolicyHolder.EmploymentStatus = "unableToWork" then "42D", "949"
                else (jsonFile.JointPolicyHolder.OccupationCode).ToString(), (jsonFile.JointPolicyHolder.BusinessCode).ToString()
            obj.jp_occupationCode <- occCode
            obj.jp_businessType <- busCode
        // Property
        obj.riskAddress_organisationName <- jsonFile.Property.InsuredAddress.OrganisationName
        obj.riskAddress_department <- jsonFile.Property.InsuredAddress.Department
        obj.riskAddress_subBuilding <- jsonFile.Property.InsuredAddress.SubBuilding
        obj.riskAddress_building <- jsonFile.Property.InsuredAddress.Building
        obj.riskAddress_number <- jsonFile.Property.InsuredAddress.Number
        obj.riskAddress_dependentThoroughfare <- jsonFile.Property.InsuredAddress.DependentThoroughfare
        obj.riskAddress_thoroughfare <- jsonFile.Property.InsuredAddress.Thoroughfare
        obj.riskAddress_doubleDependentLocality <- jsonFile.Property.InsuredAddress.DoubleDependentLocality
        obj.riskAddress_dependentLocality <- jsonFile.Property.InsuredAddress.DependentLocality
        obj.riskAddress_town <- jsonFile.Property.InsuredAddress.Town
        obj.riskAddress_traditionalCounty <- jsonFile.Property.InsuredAddress.TraditionalCounty
        obj.riskAddress_administrativeCounty <- jsonFile.Property.InsuredAddress.AdministrativeCounty
        obj.riskAddress_optionalCounty <- jsonFile.Property.InsuredAddress.OptionalCounty
        obj.riskAddress_postalCounty <- jsonFile.Property.InsuredAddress.PostalCounty
        obj.riskAddress_postcode <- jsonFile.Property.InsuredAddress.Postcode
        // Contents Cover is not mandatory either:
        if (jsonFile.JsonValue.TryGetProperty("contentsCover")).IsSome then
            obj.contents_sumInsured <- jsonFile.Property.ContentsCover.CoverAmount
            obj.contents_highRisk <- jsonFile.Property.ContentsCover.HighRiskAmount
            obj.contents_mostExpensive <- jsonFile.Property.ContentsCover.MostExpensiveHighRiskItemAmount
        // Buildings Cover
        if (jsonFile.JsonValue.TryGetProperty("buildingsCover")).IsSome then
            obj.buildings_sumInsured <- jsonFile.Property.BuildingsCover.RebuildCost
        // Locks & Security
        // Claims
        // Price Page
        risk.Add(obj)
        if numberOfFiles > n then
            loop (n + 1)
    loop 1
// --------- //

// CONTENTS //
type SpecifiedItem() =
    class
        [<DefaultValue>] val mutable riskID : int
        [<DefaultValue>] val mutable itemType : string
        [<DefaultValue>] val mutable itemValue : string
        [<DefaultValue>] val mutable itemDesc : string
    end

let mutable itemCollection = new List<SpecifiedItem>()

let itemCollectionLoad() =
    let rec loop n =
        let jsonFile = riskMaster.Load(fileArray.[n-1])
        if (jsonFile.Property.JsonValue.TryGetProperty("contentsCover")).IsSome then
            if (jsonFile.Property.ContentsCover.JsonValue.TryGetProperty("specifiedItems")).IsSome then
                let rec contentsLoop x =
                    let obj = new SpecifiedItem()
                    obj.riskID <- n
                    obj.itemType <- jsonFile.Property.ContentsCover.SpecifiedItems.[x].Type
                    obj.itemValue <- (jsonFile.Property.ContentsCover.SpecifiedItems.[x].Value).ToString()
                    obj.itemDesc <- jsonFile.Property.ContentsCover.SpecifiedItems.[x].Description
                    itemCollection.Add(obj)
                    if (jsonFile.JsonValue.TryGetProperty("jsonFile.property.contentsCover.specifiedItems.[x+1].type")).IsSome then
                        contentsLoop (x + 1)
                contentsLoop 0
            if (jsonFile.Property.ContentsCover.JsonValue.TryGetProperty("bicycles")).IsSome then
                let rec contentsLoop x =
                    let obj = new SpecifiedItem()
                    obj.riskID <- n
                    obj.itemType <- (jsonFile.Property.ContentsCover.Bicycles.[x].Type).ToString()
                    obj.itemValue <- (jsonFile.Property.ContentsCover.Bicycles.[x].Value).ToString()
                    obj.itemDesc <- jsonFile.Property.ContentsCover.Bicycles.[x].Description
                    itemCollection.Add(obj)
                    if (jsonFile.JsonValue.TryGetProperty("jsonFile.property.contentsCover.bicycles.[x+1].type")).IsSome then
                        contentsLoop (x + 1)
                contentsLoop 0
        if numberOfFiles > n then
            loop (n + 1)
    loop 1
// --------- //

// CLAIMS //
type Claim() =
    class
        [<DefaultValue>] val mutable riskID : int
        [<DefaultValue>] val mutable claim_date : DateTime
        [<DefaultValue>] val mutable claim_cost : int
    end

let mutable claimCollection = new List<Claim>()

let claimLoad() =
    let rec loop n =
        let jsonFile = riskMaster.Load(fileArray.[n-1])
        if (jsonFile.Property.Residents.JsonValue.TryGetProperty("previousClaims")).IsSome then
            let rec claimLoop x =
                let obj = new Claim()
                obj.riskID <- n
                obj.claim_date <- jsonFile.Property.Residents.PreviousClaims.[x].ClaimDate
                obj.claim_cost <- jsonFile.Property.Residents.PreviousClaims.[x].DamageAmount
                claimCollection.Add(obj)
                if (jsonFile.Property.Residents.JsonValue.TryGetProperty("previousClaims.[x+1].claimID")).IsSome then
                    claimLoop (x + 1)
            claimLoop 0
        if numberOfFiles > n then
            loop (n + 1)
    loop 1
// --------- //