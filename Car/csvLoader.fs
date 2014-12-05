module csvLoader

open System
open System.Collections.Generic
open FSharp.Data.Csv

open config

// PERSON LOADER
type Person() =
    class
        [<DefaultValue>] val mutable  TestId : string
        [<DefaultValue>] val mutable  Environment : string
        [<DefaultValue>] val mutable  Execution : string
        [<DefaultValue>] val mutable  VehicleId : string
        [<DefaultValue>] val mutable  VehicleUsage : string
        [<DefaultValue>] val mutable  ModificationId : string
        [<DefaultValue>] val mutable  FirstName : string
        [<DefaultValue>] val mutable  LastName : string
        [<DefaultValue>] val mutable  DateOfBirthDay : string
        [<DefaultValue>] val mutable  DateOfBirthMonth : string
        [<DefaultValue>] val mutable  DateOfBirthYear : string
        [<DefaultValue>] val mutable  OccupationTitleDescription : string
        [<DefaultValue>] val mutable  BusinessTypeDescription : string
        [<DefaultValue>] val mutable  IsLivingInUkSinceDay : string
        [<DefaultValue>] val mutable  IsLivingInUkSinceMonthCode : string
        [<DefaultValue>] val mutable  IsLivingInUkSinceYear : string
        [<DefaultValue>] val mutable  LicenceDateDay : string
        [<DefaultValue>] val mutable  LicenceDateMonth : string
        [<DefaultValue>] val mutable  LicenceDateYear : string
        [<DefaultValue>] val mutable  LicenceHeldCode : string
        [<DefaultValue>] val mutable  ObtainedMonth : string
        [<DefaultValue>] val mutable  ObtainedYear : string
        [<DefaultValue>] val mutable  ClaimIds : string
        [<DefaultValue>] val mutable  ConvictionIds : string
        [<DefaultValue>] val mutable  PolicyStartDateOffset : string
        [<DefaultValue>] val mutable  Email : string
        [<DefaultValue>] val mutable  MainTelephoneNumber : string
        [<DefaultValue>] val mutable  AdditionalDriverIds : string
        [<DefaultValue>] val mutable  Postcode : string
    end

let mutable personCollection = new List<Person>()

let personLoad (csv_filename : string) = 
    let csv_file = CsvFile.Load(csv_filename).Cache()
    for row in csv_file.Data do
        let obj = new Person()
        obj.TestId <- row.GetColumn("TestId")
        obj.Environment <- row.GetColumn("Environment")
        obj.Execution <- row.GetColumn("Execution")
        obj.VehicleId <- row.GetColumn("VehicleId")
        obj.VehicleUsage <- row.GetColumn("VehicleUsage")
        obj.ModificationId <- row.GetColumn("ModificationId")
        obj.FirstName <- row.GetColumn("FirstName")
        obj.LastName <- row.GetColumn("LastName")
        obj.DateOfBirthDay <- row.GetColumn("DateOfBirthDayCode")
        obj.DateOfBirthMonth <- row.GetColumn("DateOfBirthMonthCode")
        obj.DateOfBirthYear <- row.GetColumn("DateOfBirthYear")
        let occDesc, busDesc =
            if row.GetColumn("EmploymentStatusCode") = "currentlynotworking" then
                row.GetColumn("WhyNotWorking"), if row.GetColumn("WhyNotWorking") = "Unemployed" then
                                                    "None - Unemployed"
                                                elif row.GetColumn("WhyNotWorking") = "Houseperson" then
                                                    "None - Household Duties"
                                                elif row.GetColumn("WhyNotWorking") = "Retired" then
                                                    "None - Retired"
                                                else
                                                    "None - Not Employed due to Disability"
            elif row.GetColumn("EmploymentStatusCode") = "F" then
                row.GetColumn("StudentType"), "None - Student"
            else
                row.GetColumn("OccupationTitleDescription"), row.GetColumn("BusinessTypeDescription")
        obj.OccupationTitleDescription <- occDesc
        obj.BusinessTypeDescription <- busDesc
        if row.GetColumn("IsLivingInUkSinceBirth") = "No" then
            obj.IsLivingInUkSinceDay <- (DateTime(toInt(row.GetColumn("IsLivingInUkSinceYear")), toInt(row.GetColumn("IsLivingInUkSinceMonthCode")), 1).AddMonths(1).AddDays(float(-1))).Day.ToString()
            obj.IsLivingInUkSinceMonthCode <- row.GetColumn("IsLivingInUkSinceMonthCode")
            obj.IsLivingInUkSinceYear <- row.GetColumn("IsLivingInUkSinceYear")
        else
            obj.IsLivingInUkSinceDay <- row.GetColumn("DateOfBirthDayCode")
            obj.IsLivingInUkSinceMonthCode <- row.GetColumn("DateOfBirthMonthCode")
            obj.IsLivingInUkSinceYear <- row.GetColumn("DateOfBirthYear")
        obj.LicenceDateDay <- row.GetColumn("LicenceDateDayCode")
        obj.LicenceDateMonth <- row.GetColumn("LicenceDateMonthCode")
        obj.LicenceDateYear <- row.GetColumn("LicenceDateYear")
        obj.LicenceHeldCode <- row.GetColumn("LicenceHeldCode")
        obj.ObtainedMonth <- row.GetColumn("ObtainedMonthCode")
        obj.ObtainedYear <- row.GetColumn("ObtainedYear")
        obj.ClaimIds <- row.GetColumn("ClaimIds")
        obj.ConvictionIds <- row.GetColumn("ConvictionIds")
        obj.PolicyStartDateOffset <- row.GetColumn("PolicyStartDateOffset")
        obj.Email <- row.GetColumn("Email")
        obj.MainTelephoneNumber <- row.GetColumn("MainTelephoneNumber").Replace(" ","")
        obj.AdditionalDriverIds <- row.GetColumn("AdditionalDriverIds")
//        obj.Postcode <- TODO
        personCollection.Add(obj)
//
// ADDITIONAL DRIVER LOADER
type Additional() =
    class
        [<DefaultValue>] val mutable  AdditionalDriverId : string
        [<DefaultValue>] val mutable  FirstName : string
        [<DefaultValue>] val mutable  LastName : string
        [<DefaultValue>] val mutable  DateOfBirthDay : string
        [<DefaultValue>] val mutable  DateOfBirthMonth : string
        [<DefaultValue>] val mutable  DateOfBirthYear : string
        [<DefaultValue>] val mutable  Relationship : string
        [<DefaultValue>] val mutable  OccupationTitleDescription : string
        [<DefaultValue>] val mutable  BusinessTypeDescription : string
        [<DefaultValue>] val mutable  IsLivingInUkSinceDay : string
        [<DefaultValue>] val mutable  IsLivingInUkSinceMonthCode : string
        [<DefaultValue>] val mutable  IsLivingInUkSinceYear : string
        [<DefaultValue>] val mutable  LicenceDateDay : string
        [<DefaultValue>] val mutable  LicenceDateMonth : string
        [<DefaultValue>] val mutable  LicenceDateYear : string
        [<DefaultValue>] val mutable  LicenceHeldCode : string
        [<DefaultValue>] val mutable  ClaimIds : string
        [<DefaultValue>] val mutable  ConvictionIds : string
    end

let mutable additionalCollection = new List<Additional>()

let additionalLoad (csv_filename : string) = 
    let csv_file = CsvFile.Load(csv_filename).Cache()
    for row in csv_file.Data do
        let obj = new Additional()
        obj.AdditionalDriverId <- row.GetColumn("AdditionalDriverId")
        obj.FirstName <- row.GetColumn("FirstName")
        obj.LastName <- row.GetColumn("LastName")
        obj.DateOfBirthDay <- row.GetColumn("DateOfBirthDayCode")
        obj.DateOfBirthMonth <- row.GetColumn("DateOfBirthMonthCode")
        obj.DateOfBirthYear <- row.GetColumn("DateOfBirthYear")
        obj.Relationship <- row.GetColumn("Relationship")
        let occDesc, busDesc =
            if row.GetColumn("EmploymentStatusCode") = "currentlynotworking" then
                row.GetColumn("WhyNotWorking"), "Not In E;mployment" //if row.GetColumn("WhyNotWorking") = "Unemployed" then
                                                                    //    "None - Unemployed"
                                                                    //elif row.GetColumn("WhyNotWorking") = "Houseperson" then
                                                                    //    "None - Household Duties"
                                                                    //elif row.GetColumn("WhyNotWorking") = "Retired" then
                                                                    //    "None - Retired"
                                                                    //else
                                                                    //    "None - Not Employed due to Disability"
            elif row.GetColumn("EmploymentStatusCode") = "F" then
                row.GetColumn("StudentType"), "None - Student"
            else
                row.GetColumn("OccupationTitleDescription"), row.GetColumn("BusinessTypeDescription")
        obj.OccupationTitleDescription <- occDesc
        obj.BusinessTypeDescription <- busDesc
        if row.GetColumn("IsLivingInUkSinceBirth") = "No" then
            obj.IsLivingInUkSinceDay <- (DateTime(toInt(row.GetColumn("IsLivingInUkSinceYear")), toInt(row.GetColumn("IsLivingInUkSinceMonthCode")), 1).AddMonths(1).AddDays(float(-1))).Day.ToString()
            obj.IsLivingInUkSinceMonthCode <- row.GetColumn("IsLivingInUkSinceMonthCode")
            obj.IsLivingInUkSinceYear <- row.GetColumn("IsLivingInUkSinceYear")
        else
            obj.IsLivingInUkSinceDay <- row.GetColumn("DateOfBirthDayCode")
            obj.IsLivingInUkSinceMonthCode <- row.GetColumn("DateOfBirthMonthCode")
            obj.IsLivingInUkSinceYear <- row.GetColumn("DateOfBirthYear")
        obj.LicenceDateDay <- row.GetColumn("LicenceDateDayCode")
        obj.LicenceDateMonth <- row.GetColumn("LicenceDateMonthCode")
        obj.LicenceDateYear <- row.GetColumn("LicenceDateYear")
        obj.LicenceHeldCode <- row.GetColumn("LicenceHeldCode")
        obj.ClaimIds <- row.GetColumn("ClaimIds")
        obj.ConvictionIds <- row.GetColumn("ConvictionIds")
        additionalCollection.Add(obj)
//
// CLAIM LOADER
type Claim() =
    class
        [<DefaultValue>] val mutable  ClaimsId : string
        [<DefaultValue>] val mutable  ClaimDateMonthCode : string
        [<DefaultValue>] val mutable  ClaimDateYear : string
        [<DefaultValue>] val mutable  ClaimCost : string
    end

let mutable claimCollection = new List<Claim>()

let claimLoad (csv_filename : string) = 
    let csv_file = CsvFile.Load(csv_filename).Cache()
    for row in csv_file.Data do
        let obj = new Claim()
        obj.ClaimsId <- row.GetColumn("ClaimsId")
        obj.ClaimDateMonthCode <- row.GetColumn("ClaimDateMonthCode")
        obj.ClaimDateYear <- row.GetColumn("ClaimDateYear")
        obj.ClaimCost <- row.GetColumn("ClaimCost")
        claimCollection.Add(obj)
//
// CONVICTION LOADER
type Conviction() =
    class
        [<DefaultValue>] val mutable  ConvictionId : string
        [<DefaultValue>] val mutable  Conviction : string
        [<DefaultValue>] val mutable  ConvictionDateMonthCode : string
        [<DefaultValue>] val mutable  ConvictionDateYear : string
        [<DefaultValue>] val mutable  NumberOfPoints : string
        [<DefaultValue>] val mutable  FineAmount : string
        [<DefaultValue>] val mutable  BanLength : string
        [<DefaultValue>] val mutable  BreathalyserReading : string
    end

let mutable convictionCollection = new List<Conviction>()

let convictionLoad (csv_filename : string) = 
    let csv_file = CsvFile.Load(csv_filename).Cache()
    for row in csv_file.Data do
        let obj = new Conviction()
        obj.ConvictionId <- row.GetColumn("ConvictionId")
        obj.Conviction <- row.GetColumn("Conviction")
        obj.ConvictionDateMonthCode <- row.GetColumn("ConvictionDateMonthCode")
        obj.ConvictionDateYear <- row.GetColumn("ConvictionDateYear")
        obj.NumberOfPoints <- row.GetColumn("NumberOfPoints")
        obj.FineAmount <- row.GetColumn("FineAmount")
        obj.BanLength <- row.GetColumn("BanLength")
        obj.BreathalyserReading <- row.GetColumn("BreathalyserReading")
        convictionCollection.Add(obj)
//
// VEHICLE LOADER
type Vehicle() =
    class
        [<DefaultValue>] val mutable  VehicleId : string
        [<DefaultValue>] val mutable  RegistrationNumber : string
        [<DefaultValue>] val mutable  CurrentValue : string
        [<DefaultValue>] val mutable  RegYear : string
    end

let mutable vehicleCollection = new List<Vehicle>()

let vehicleLoad (csv_filename : string) = 
    let csv_file = CsvFile.Load(csv_filename).Cache()
    for row in csv_file.Data do
        let obj = new Vehicle()
        obj.VehicleId <- row.GetColumn("VehicleId")
        obj.RegistrationNumber <- row.GetColumn("RegistrationNumber")
        obj.CurrentValue <- row.GetColumn("CurrentValue")
        obj.RegYear <- row.GetColumn("RegYear")
        vehicleCollection.Add(obj)
//
// CAR USAGE LOADER
type CarUsage() =
    class
        [<DefaultValue>] val mutable  CarUsageId : string
        [<DefaultValue>] val mutable  DateOfPurchaseMonthCode : string
        [<DefaultValue>] val mutable  DateOfPurchaseYear : string
        [<DefaultValue>] val mutable  AnnualMileage : string
        [<DefaultValue>] val mutable  AnnualBusinessMileage : string
        [<DefaultValue>] val mutable  VehicleKeptAtNightAddress : string
    end

let mutable carUsageCollection = new List<CarUsage>()

let carUsageLoad (csv_filename : string) = 
    let csv_file = CsvFile.Load(csv_filename).Cache()
    for row in csv_file.Data do
        let obj = new CarUsage()
        obj.CarUsageId <- row.GetColumn("CarUsageId")
        obj.DateOfPurchaseMonthCode <- row.GetColumn("DateOfPurchaseMonthCode")
        obj.DateOfPurchaseYear <- row.GetColumn("DateOfPurchaseYear")
        obj.AnnualMileage <- row.GetColumn("AnnualMileage")
        obj.AnnualBusinessMileage <- row.GetColumn("AnnualBusinessMileage")
        obj.VehicleKeptAtNightAddress <- row.GetColumn("VehicleKeptAtNightAddress")
        carUsageCollection.Add(obj)
//