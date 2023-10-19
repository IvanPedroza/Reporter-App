module EthanolPrecipitations

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
open System.Diagnostics
open PrecipitationHelpers
open CoversheetHelpers

// Function use to fill out Ethanol Precipitation batch reacords
let precipitationStart (inputParams : string list) (precipitationForm : string) (rqstForm : string) (reporter : ExcelWorksheet) (myTools : ExcelWorksheet) = 
    
    // User interface
    let reagentsInput = "precipitations"
    Console.WriteLine "Which bench are you working at?"
    let benchInput = Console.ReadLine ()
    let user = Environment.UserName
    
   
    //Starting point for manipulating sheet info and storing it in "param"
    for param in inputParams do

        // Reads in equipment and lot IDs from LIMS
        let lot = listFunction param reporter 1
        let csname = listFunction param reporter 2
        let geneNumber = listFunction param reporter 5
        let scale = listFunction param reporter 6
        let species = listFunction param reporter 3  
        let customer = listFunction param reporter 4 
        let rxnNumber = listFunction param reporter 8 
        let formulation = listFunction param reporter 9 
        let ship = listFunction param reporter 10 
        let acetateLot = listFunction reagentsInput myTools 2
        let acetateExpiration = listFunction reagentsInput myTools 3
        let ethanolLot = listFunction reagentsInput myTools 4
        let ethanolExpiration = listFunction reagentsInput myTools 5
        let ethanolUBD = listFunction reagentsInput myTools 6
        let calibrationDate = listFunction benchInput myTools 2 
        let p1000 = listFunction benchInput myTools 3
        let p1000Id = listFunction benchInput myTools 4
        let p200 = listFunction benchInput myTools 5
        let p200Id = listFunction benchInput myTools 6
        let p20 = listFunction benchInput myTools 7
        let p20Id = listFunction benchInput myTools 8
        let mc12P200 = listFunction benchInput myTools 19
        let mc12P200Id = listFunction benchInput myTools 20
        let mc12P20 = listFunction benchInput myTools 17
        let mc12P20Id = listFunction benchInput myTools 18


        // Fills out Batch Records for type of CS build
        if param.EndsWith ("N", StringComparison.InvariantCultureIgnoreCase) then 
            // Reads in Batch Record template
            let docArray = File.ReadAllBytes(rqstForm)
            use _copyDoc = new MemoryStream(docArray)
            use myDocument = WordprocessingDocument.Open(_copyDoc, true)
            let body = myDocument.MainDocumentPart.Document.Body
            let docArray = File.ReadAllBytes(rqstForm)
            use copyDoc = new MemoryStream(docArray)
            use myDocument = WordprocessingDocument.Open(copyDoc, true)
            let body = myDocument.MainDocumentPart.Document.Body

            // Finds the text for library concentration checkboxes
            let concentrationCheck = "☒" //used to replace checked box text
            let lessThanSix = (determiningConcentration body 1 1)
            let sixtofourhundred =(determiningConcentration body 1 2)
            let fourhundredplus =  (determiningConcentration body 1 3)

            //Converts gene count string to int and fills in concentration text under conditional statemets
            let geneCount = geneNumber |> int
            match geneCount with
                | _ when geneCount < 6  -> 
                    lessThanSix.Text <- concentrationCheck
                | _ when geneCount < 400 -> 
                    sixtofourhundred.Text <- concentrationCheck
                | _ -> 
                    fourhundredplus.Text <- concentrationCheck


            // Formats the shipping date 
            let formattedShip =
                match ship with 
                    | "TBD"  -> "TBD"
                    | _ -> 
                        let firststring = ship.Substring(0,3)
                        let secondstring = ship.Substring(3,2).ToLower()
                        firststring + secondstring + year

            //Calls funtion to determine codeset type from the name of the codeset
            let codeSetTypes = determine csname formulation

            //Fills each text field of Word Doc with Excel info
            (fillTopCells body 0 ).Text <- codeSetTypes
            (fillTopCells body 1).Text <- rxnNumber
            (fillTopCells body 2).Text <- formulation
            (staticCells body 0 1 0 0).Text <- ""
            (staticCells body 0 1 0 0).Text <- lot
            (staticCells body 0 1 1 0).Text <- csname
            (staticCells body 0 1 2 0).Text <- species
            (dropdownCells body 0).Text <- customer
            (staticCells body 0 1 3 0).Text <- geneNumber.ToString()
            (dropdownCells body 1).Text <- scale.ToString()
            (dropdownCells body 2).Text <- formattedShip

            // Saves filled out cover page 
            myDocument.SaveAs("C:\\Users\\" + user + "\\AppData\\Local\\Temp\\ "+param + " Request Form" + ".docx") |> ignore
        else 
            ignore()
    
        // Reads in Batch Record and starts processing
        let memoryStream = new MemoryStream()
        use fileStream = new FileStream(precipitationForm, FileMode.Open, FileAccess.Read)
        fileStream.CopyTo(memoryStream)
        let myDocument = WordprocessingDocument.Open(memoryStream, true)
        let body = myDocument.MainDocumentPart.Document.Body

        //F ills CS indentifying info for pages one and two
        (getCsInfo body 3 7).Text <- lot + " " + csname
        (getCsInfo body 3 13).Text <- geneNumber
        (getCsInfo body 3 18).Text <- scale
        (getCsInfo body 6 2).Text <- lot + " " + csname
        (getCsInfo body 6 9).Text <- geneNumber
        (getCsInfo body 6 15).Text <- scale  
        (getCsInfo body 10 2).Text <- lot + " " + csname
        (getCsInfo body 10 5).Text <- geneNumber
        (getCsInfo body 10 10).Text <- scale
        (getCsInfo body 15 2).Text <- lot + " " + csname
        (getCsInfo body 15 5).Text <- geneNumber
        (getCsInfo body 15 10).Text <- scale    

        // Fills out reagent lot numbers from LIMS 
        (fillCells body 0 1 2 0 0).Text <- lot
        (fillCells body 0 1 3 0 0).Text <- "N/A"
        (fillCells body 0 2 2 0 0).Text <- acetateLot
        (fillCells body 0 2 3 0 0).Text <- acetateExpiration
        (fillCells body 0 3 2 0 0).Text <- ethanolLot
        (fillCells body 0 3 3 0 0).Text <- ethanolExpiration
        (fillCells body 0 3 3 1 2).Text <- ethanolUBD
        (fillCells body 0 5 0 0 0).Text <- p1000
        (fillCells body 0 5 1 0 0).Text <- p1000Id
        (fillCells body 0 5 2 0 0).Text <- calibrationDate
        (fillCells body 0 6 2 0 0).Text <- calibrationDate
        (fillCells body 0 7 2 0 0).Text <- calibrationDate

        // Gets text in cells which value depends on the scale of the build and the bench where the process will be done
        let acetatePipetteSize = (fillCells body 0 6 0 0 0)
        let acetatePipetteId =  (fillCells body 0 6 1 0 0)
        let poolingPipetteSize = (fillCells body 0 7 0 0 0)
        let poolingPipetteId = (fillCells body 0 7 1 0 0)
        let precipitationVolume = (fillCells body 2 3 1 5 3)
        let acetateVolume = (fillCells body 2 4 1 0 3)
        let alcoholVolume = (fillCells body 2 6 1 0 4)

        //Fills out theoretical volume of pooled reporters
        let theoreticalVolume = thereticalVolumes (scale |> float) (geneNumber |> float)
        (fillCells body 1 4 1 0 2).Text <- theoreticalVolume.ToString()

        //Calculates precipitation and reagent volumes and then assigns eqipment needed for working with those volumes
        if scale = "0.15" then 
            let volume, acetate, alcohol = calculations 5.1 (geneNumber |> float)

            precipitationVolume.Text <- volume.ToString()
            acetateVolume.Text <- acetate.ToString()
            alcoholVolume.Text <- alcohol.ToString()
            poolingPipetteSize.Text <- mc12P20
            poolingPipetteId.Text <- mc12P20Id

            match acetate with 
            | _ when acetate < 20.0 ->
                acetatePipetteSize.Text <- p20
                acetatePipetteId.Text <- p20Id
            | _ when acetate >= 20.0 ->
                acetatePipetteSize.Text <- p200
                acetatePipetteId.Text <- p200Id

        elif scale = "0.25" then 
            let volume, acetate, alcohol = calculations 8.5 (geneNumber |> float)

            precipitationVolume.Text <- volume.ToString()
            acetateVolume.Text <- acetate.ToString()
            alcoholVolume.Text <- alcohol.ToString()
            poolingPipetteSize.Text <- mc12P20
            poolingPipetteId.Text <- mc12P20Id

            match acetate with 
            | _ when acetate < 20.0 ->
                acetatePipetteSize.Text <- p20
                acetatePipetteId.Text <- p200Id
            | _ when acetate >= 20.0 ->
                acetatePipetteSize.Text <- p200
                acetatePipetteId.Text <- p200Id

        else
            let volume = System.Math.Ceiling(theoreticalVolume)
            let acetate = System.Math.Ceiling(volume / 9.0)
            let alcohol = System.Math.Ceiling((volume + acetate) * 2.5)

            (fillCells body 2 2 1 0 2).Text <- volume.ToString()
            acetateVolume.Text <- acetate.ToString()
            alcoholVolume.Text <- alcohol.ToString()
            acetatePipetteSize.Text <- p200
            acetatePipetteId.Text <- p200Id
            poolingPipetteSize.Text <- mc12P200
            poolingPipetteId.Text <- mc12P200Id    

        // Saves documet to temp directory for printing before deleting
        let path = "C:\\users\\" + user + "\\AppData\\Local\\Temp\\ "+param + " Precipitation Form" + ".docx"
        myDocument.SaveAs(path).Close() |> ignore




