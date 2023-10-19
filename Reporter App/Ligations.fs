module Ligations

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq
open System.IO
open LigationHelpers
open CoversheetHelpers

// Function used to fillout ligation Batch Records
let ligationStart (inputParams : string list) (ligationForm : string) (requestForm : string) (reporter : ExcelWorksheet) (oligoStamps : ExcelWorksheet) (myTools : ExcelWorksheet) =
    // Gets user for use in path and error logging
    let user = Environment.UserName
    // User interface
    Console.WriteLine "Will you be digesting all these lots together today?"
    Console.WriteLine ("Yes or No")
    let digestInput = Console.ReadLine ()
    let digest =
        if digestInput.Equals("Y", StringComparison.InvariantCultureIgnoreCase) then 
            "Yes"
        else 
            digestInput
    Console.WriteLine "Which bench space are you using?"
    let benchInput = Console.ReadLine ()
    let reagentsInput = "rpligations"
    
    // Reads in Batch Record template for manipulation
    let backboneSheet = new FileInfo("W:/Production/DV1/code sets/Arrayed Backbones/In process/Array Tracking.xlsx")
    use backbonePackage = new ExcelPackage (backboneSheet)
    let backbones = backbonePackage.Workbook.Worksheets.["Array Tracking"]
    
    let digestReagents = "digest"
    let mutable ligationReactions = []
    let mutable digestReactions = []
    

    //Starting point for manipulating sheet info and storing it in "param"
    for param in inputParams do
            
        // Reads in lot numbers from LIMS
        let indexedParam = param.[1..param.Length] |> string
        let backboneLot = listFunction indexedParam backbones 2 true

        // Retrieves IDs for reagents and equipment from LIMS
        let lot = listFunction param reporter 1 false
        let csName = listFunction param reporter 2 false
        let geneNumber = listFunction param reporter 5 false |> float
        let scale = listFunction param reporter 6 false |> float
        let species = listFunction param reporter 3 false 
        let customer = listFunction param reporter 4 false
        let rxnNumber = listFunction param reporter 8 false
        let formulation = listFunction param reporter 9 false
        let ship = listFunction param reporter 10 false
        let RPLLot = listFunction reagentsInput myTools 2 false
        let RPLExpiration = listFunction reagentsInput myTools 3 false
        let T4BufferLot = listFunction reagentsInput myTools 4 false
        let T4BufferExpiration = listFunction reagentsInput myTools 5 false
        let ATPLot = listFunction reagentsInput myTools 6 false
        let ATPExpiration = listFunction reagentsInput myTools 7 false
        let H2OLot = listFunction reagentsInput myTools 8 false
        let H2OExpiration = listFunction reagentsInput myTools 9 false
        let H2Oubd = listFunction reagentsInput myTools 10 false
        let EnzymeLot = listFunction reagentsInput myTools 11 false
        let EnzymeExp = listFunction reagentsInput myTools 12 false
        let digestBufferLot = listFunction digestReagents myTools 2 false
        let digestBufferExpiration = listFunction digestReagents myTools 3 false
        let cutterLot = listFunction digestReagents myTools 4 false
        let cutterExpiration = listFunction digestReagents myTools 5 false
        let reLot = listFunction digestReagents myTools 6 false
        let reExpiration = listFunction digestReagents myTools 7 false
        let NA = "N/A"
        let pipetteCalibration = listFunction benchInput myTools 2 false
        let p1000Id = listFunction benchInput myTools 4 false
        let p200Id = listFunction benchInput myTools 6 false
        let p20Id = listFunction benchInput myTools 8 false
        let p2Id = listFunction benchInput myTools 10 false
        let p2000Id = listFunction benchInput myTools 12 false
        let mc8P20Id = listFunction benchInput myTools 14 false
        let mc12P20Id = listFunction benchInput myTools 16 false
        let mc8P200Id = listFunction benchInput myTools 18 false
        let mc12P200Id = listFunction benchInput myTools 20 false


        // Fills out Batch Record depending of type and criteria of each build
        if param.EndsWith("RW", StringComparison.InvariantCultureIgnoreCase) then 
            ignore()
        else 
            let docArray = File.ReadAllBytes(requestForm)
            use _copyDoc = new MemoryStream(docArray)
            use myDocument = WordprocessingDocument.Open(_copyDoc, true)
            let body = myDocument.MainDocumentPart.Document.Body
               
            // Reads in template, makes virtual copy and starts processing
            let docArray = File.ReadAllBytes(requestForm)
            use copyDoc = new MemoryStream(docArray)
            use myDocument = WordprocessingDocument.Open(copyDoc, true)
            let body = myDocument.MainDocumentPart.Document.Body

            //next block finds the text for library concentration checkboxes
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


            //Formats the shipping date 
            let formattedShip =
                match ship with 
                    | "TBD"  -> "TBD"
                    | _ -> 
                        let firststring = ship.Substring(0,3)
                        let secondstring = ship.Substring(3,2).ToLower()
                        firststring + secondstring + year

            //Calls funtion to determine codeset type from the name of the codeset
            let codeSetTypes = determine csName formulation

            //Fills each text field of Word Doc with Excel info
            (fillTopCells body 0 ).Text <- codeSetTypes
            (fillTopCells body 1).Text <- rxnNumber
            (fillTopCells body 2).Text <- formulation
            (staticCells body 0 1 0 0).Text <- ""
            (staticCells body 0 1 0 0).Text <- lot
            (staticCells body 0 1 1 0).Text <- csName.Trim()
            (staticCells body 0 1 2 0).Text <- species
            (dropdownCells body 0).Text <- customer
            (staticCells body 0 1 3 0).Text <- geneNumber.ToString()
            (dropdownCells body 1).Text <- scale.ToString()
            (dropdownCells body 2).Text <- formattedShip

            //Saves filled out Doc
            myDocument.SaveAs($"C:\\Users\\{user}\\AppData\\Local\\Temp\\{param} Request Form" + ".docx") |> ignore



        //Reads in Word Doc and starts processing
        let docArray = File.ReadAllBytes(ligationForm)
        use _copyDoc = new MemoryStream(docArray)
        use myDocument = WordprocessingDocument.Open(_copyDoc, true)
        let body = myDocument.MainDocumentPart.Document.Body

        let requested, added, ligator, buffer, atp, ligase, masterMix = oligoStamp(scale |> float)
        
        let lastLot = inputParams.Last()
        let lastLotScale = (listFunction lastLot reporter 6 false) |> float


        // Fills out rework Batch Record if needed
        if param.EndsWith("RW", StringComparison.InvariantCultureIgnoreCase) then
            Console.WriteLine ("How many reworks are you ligationg for " + param)
            let rwNumber = Console.ReadLine ()


            Console.WriteLine ("
            0 = 1X
            1 = 4X
            2 = 5X
            3 = 4X & 5X
            ")
            Console.WriteLine ("Enter the number above that corresponds with the RW scake for " + param)
            let rwScale = Console.ReadLine ()

            let formattedOligo = requested |> float

            let rwOligo = 
                match rwScale with 
                    | _ when rwScale = "0" -> requested
                    | _ when rwScale = "1" -> formattedOligo * 4.0 |> string
                    | _ when rwScale = "2" -> formattedOligo * 5.0 |> string
                    | _ when rwScale = "3" -> (formattedOligo * 4.0 |> string) + "|" + (formattedOligo * 5.0 |> string)

            let rwOligoAddded = 
                match rwScale with 
                    | _ when rwScale = "0" -> (formattedOligo - 3.0)
                    | _ when rwScale = "1" -> (formattedOligo * 4.0) - 3.0
                    | _ when rwScale = "2" -> (formattedOligo * 5.0) - 3.0
                    

            let varOligoToUse = 
                if ((rwOligoAddded % rwOligoAddded) <> 0.0) && (rwScale <> "3") then 
                    let formattedOligo = System.Math.Round(rwOligoAddded, 1)
                    formattedOligo.ToString()
                else 
                    rwOligoAddded.ToString()
                

            let mmStamp = 
                let formatedMm = (masterMix |> float)
                match rwScale with 
                    | _ when rwScale = "0" -> masterMix
                    | _ when rwScale = "1" -> formatedMm * 4.0 |> string
                    | _ when rwScale = "2" -> formatedMm * 5.0 |> string
                    | _ when rwScale = "3" -> (formatedMm * 4.0 |> string) + "|" + (formatedMm * 5.0 |> string)

            let rwGeneNorm = 
                match rwScale with 
                    | _ when rwScale = "0" -> 1.0
                    | _ when rwScale = "1" -> 4.0
                    | _ when rwScale = "2" -> 5.0
                    | _ when rwScale = "3" -> 5.0 



            if not (scale = lastLotScale) then 
                let scaling = ((scale |> float) / (lastLotScale |> float))
                let scaledGenes = (rwNumber |> float) * scaling
                let rwGeneNormalization = scaledGenes * rwGeneNorm
                ligationReactions <- rwGeneNormalization :: ligationReactions
            else 
                let rwGeneNormalization = (rwNumber |> float) * rwGeneNorm
                ligationReactions <- rwGeneNormalization :: ligationReactions

            digestReactions <- (rwNumber |> float) :: digestReactions


            (csId body 3 9).Text <- rwNumber + "/" + (geneNumber |> string)
            (fillCells body 1 2 1 3 2 0).Text <- rwOligo
            (fillCells body 1 9 1 4 3 0).Text <- if rwScale <> "3" then varOligoToUse else ((formattedOligo * 4.0) - 3.0).ToString() + "|" + ((formattedOligo * 5.0) - 3.0).ToString()
            (fillCells body 1 10 1 3 2 0).Text <- mmStamp
        else
            (csId body 3 9).Text <- geneNumber |> string
            (fillCells body 1 2 1 3 2 0).Text <- requested
            (fillCells body 1 9 1 4 3 0).Text <- added
            (fillCells body 1 10 1 3 2 0).Text <- masterMix


        //Creates two lists - a list of total probe number for digest calculations and one of normalized probe number for ligation calculations
        if param.EndsWith ("RW", StringComparison.InvariantCultureIgnoreCase) then
            ignore()
        else
            digestReactions <- geneNumber :: digestReactions


            let calculationsScale =
                if lastLotScale <= 0.5 then 
                    let normalizedScale = 0.5
                    normalizedScale
                else 
                    lastLotScale

            if not (scale = lastLotScale) then
                if scale <= 0.5 then  
                    let scaling = 0.5
                    let newScale = scaling / calculationsScale
                    let reactions = (newScale * geneNumber)
                    ligationReactions <- reactions :: ligationReactions

                else
                    let scaling = scale / calculationsScale
                    let reactions = scaling * geneNumber
                    ligationReactions <- reactions :: ligationReactions
            else 
                ligationReactions <- geneNumber :: ligationReactions

        
        // Determines number of plates that make up a CS
        let plateTotal = System.Math.Ceiling(geneNumber / 96.0) |> int
        let iterator = [1..9]
        let mutable naList = []
        for i in iterator do 
            if i <= plateTotal then 
                let x = ""
                naList <- x :: naList
            else 
                let x = "N/A"
                naList <- x :: naList

        let oligoStampDate = oligoStampDateFinder param oligoStamps 1
       
        //Finds text in the Docx table and fills cells with build specific identifying values
        (csId body 3 3).Text <- lot + " " + (csName |> string)
        (csId body 3 16).Text <- scale |> string

        (fillCells body 0 2 2 0 0 0).Text <- oligoStampDate
        (fillCells body 0 2 3 0 0 0).Text <- NA
        (fillCells body 0 3 2 0 0 0).Text <- backboneLot
        (fillCells body 0 3 3 0 0 0).Text <- NA
        (fillCells body 0 4 2 0 0 0).Text <- RPLLot
        (fillCells body 0 4 3 0 0 0).Text <- RPLExpiration
        (fillCells body 0 5 2 0 0 0).Text <- T4BufferLot
        (fillCells body 0 5 3 0 0 0).Text <- T4BufferExpiration
        (fillCells body 0 6 2 0 0 0).Text <- ATPLot
        (fillCells body 0 6 3 0 0 0).Text <- ATPExpiration
        (fillCells body 0 7 2 0 0 0).Text <- H2OLot
        (fillCells body 0 7 3 0 1 0).Text <- H2OExpiration
        (fillCells body 0 7 3 1 2 0).Text <- H2Oubd //changed
        (fillCells body 0 8 2 0 0 0).Text <- EnzymeLot
        (fillCells body 0 8 3 0 0 0).Text <- EnzymeExp
        (fillCells body 0 10 1 0 0 0).Text <- naList.[8]
        (fillCells body 0 10 2 0 0 0).Text <- naList.[8]
        (fillCells body 0 11 1 0 1 0).Text <- naList.[7]
        (fillCells body 0 11 2 0 0 0).Text <- naList.[7]
        (fillCells body 0 12 1 0 1 0).Text <- naList.[6]
        (fillCells body 0 12 2 0 0 0).Text <- naList.[6]
        (fillCells body 0 13 1 0 1 0).Text <- naList.[5]
        (fillCells body 0 13 2 0 0 0).Text <- naList.[5]
        (fillCells body 0 14 1 0 1 0).Text <- naList.[4]
        (fillCells body 0 14 2 0 0 0).Text <- naList.[4]
        (fillCells body 0 15 1 0 1 0).Text <- naList.[3]
        (fillCells body 0 15 2 0 0 0).Text <- naList.[3]
        (fillCells body 0 16 1 0 1 0).Text <- naList.[2]
        (fillCells body 0 16 2 0 0 0).Text <- naList.[2]
        (fillCells body 0 17 1 0 1 0).Text <- naList.[1]
        (fillCells body 0 17 2 0 0 0).Text <- naList.[1]
        (fillCells body 0 18 1 0 1 0).Text <- naList.[0]
        (fillCells body 0 18 2 0 0 0).Text <- naList.[0]

        // Determine which pipete user should use
        let mcStampPipette = 
                match scale with 
                | _ when scale <= 2.5 && geneNumber <= 64.0 -> mc8P20Id //8 chan 20
                | _ when scale <= 2.5 && geneNumber > 64.0 -> "12 chan 20"
                | _ when scale = 3.0 && geneNumber <= 64.0 -> mc8P200Id //8 chan 20
                | _ when scale = 3.0 && geneNumber > 64.0 -> "12 chan 20"
                | _ when scale > 3.0 && geneNumber <= 64.0 -> mc12P200Id //8 chan 200
                | _ when scale > 3.0 && geneNumber > 64.0 -> "12 chan 200"

        // Fills out equipment used
        (fillCells body 0 20 0 0 0 0).Text <- mcStampPipette
        (fillCells body 0 21 0 0 0 0).Text <- p1000Id
        (fillCells body 0 22 0 0 0 0).Text <- p200Id
        (fillCells body 0 23 0 0 0 0).Text <- if scale = 3.0 then "12-can 200" else NA
        (fillCells body 0 24 0 0 0 0).Text <- NA
        (fillCells body 0 25 0 0 0 0).Text <- NA
        (fillCells body 0 26 0 0 0 0).Text <- NA

        (fillCells body 0 20 1 0 0 0).Text <- pipetteCalibration
        (fillCells body 0 21 1 0 0 0).Text <- pipetteCalibration
        (fillCells body 0 22 1 0 0 0).Text <- pipetteCalibration
        (fillCells body 0 23 1 0 0 0).Text <- if scale = 3.0 then pipetteCalibration else NA
        (fillCells body 0 24 1 0 0 0).Text <- NA
        (fillCells body 0 25 1 0 0 0).Text <- NA
        (fillCells body 0 26 1 0 0 0).Text <- NA

        //Fills out calculations on last document from top to bottom order of document
        if param = lastLot then
            
            let reactionNumber = roundupbyfive ((ligationReactions.Sum()) * 1.1)
           
            let atpNeeded = (atp * reactionNumber) 
            let atpToDilute = roundupbyfive((15.0 * atpNeeded) / 100.0)

            let totalDilutedATP = System.Math.Round(((100.0 * atpToDilute) / 15.0), 1)
            
            let dilutant = System.Math.Round((totalDilutedATP - atpToDilute), 1)

            let rpLigator = (reactionNumber * ligator)
            let ligaseBuffer = (reactionNumber * buffer)
            let ligationatp = System.Math.Round((reactionNumber * atp),2)
            let ligaseEnzyme = (reactionNumber * ligase)
            let aliquots = System.Math.Round(((rpLigator + ligaseBuffer + ligaseEnzyme + atp) / 8.0), 0)

            (fillCells body 1 6 1 0 3 0).Text <- reactionNumber.ToString()
            (fillCells body 1 6 1 2 2 0).Text <- rpLigator.ToString()
            (fillCells body 1 6 1 4 3 0).Text <- ligaseBuffer.ToString()
            (fillCells body 1 6 1 6 3 0).Text <- ligationatp.ToString()
            (fillCells body 1 6 1 8 1 0).Text <- ligaseEnzyme.ToString()
            (fillCells body 1 6 1 11 1 0).Text <- aliquots.ToString()
            (fillCells body 1 5 1 7 2 0).Text <- atpToDilute.ToString()
            (fillCells body 1 5 1 7 8 0).Text <- totalDilutedATP.ToString()
            (fillCells body 1 5 1 10 1 0).Text <- totalDilutedATP.ToString()
            (fillCells body 1 5 1 10 6 0).Text <- atpToDilute.ToString()
            (fillCells body 1 5 1 10 10 0).Text <- dilutant.ToString()
            ignore()
        
        //Adds footnotes and a reference note in comments if condition is met 
        calNote body inputParams param

        //Starts processing digest section of document if condition is met
        if digest.Equals( "Yes", StringComparison.InvariantCultureIgnoreCase) then 
            
            //Calculates reaction number from list and calls function to calculate reagent quantities
            let totalDigestRxns = roundupbyfive ((digestReactions.Sum()) * 1.1)
            let DEPC, neBuffer, cutterMix, pviiEnyme = digesting totalDigestRxns

            //Fills in build specific identifying values
            (fillCells body 2 2 2 0 0 0).Text <- H2OLot
            (fillCells body 2 2 3 0 1 0).Text <- H2OExpiration
            (fillCells body 2 2 3 1 1 0).Text <- H2Oubd
            (fillCells body 2 3 2 0 0 0).Text <- digestBufferLot
            (fillCells body 2 3 3 0 0 0).Text <- digestBufferExpiration
            (fillCells body 2 4 2 0 0 0).Text <- cutterLot
            (fillCells body 2 4 3 0 0 0).Text <- cutterExpiration
            (fillCells body 2 5 2 0 0 0).Text <- reLot
            (fillCells body 2 5 3 0 0 0).Text <- reExpiration

            (fillCells body 2 8 1 0 1 0).Text <- naList.[7]
            (fillCells body 2 8 2 0 0 0).Text <- naList.[7]
            (fillCells body 2 9 1 0 1 0).Text <- naList.[6]
            (fillCells body 2 9 2 0 0 0).Text <- naList.[6]
            (fillCells body 2 10 1 0 1 0).Text <- naList.[5]
            (fillCells body 2 10 2 0 0 0).Text <- naList.[5]
            (fillCells body 2 11 1 0 1 0).Text <- naList.[4]
            (fillCells body 2 11 2 0 0 0).Text <- naList.[4]
            (fillCells body 2 12 1 0 1 0).Text <- naList.[3]
            (fillCells body 2 12 2 0 0 0).Text <- naList.[3]
            (fillCells body 2 13 1 0 1 0).Text <- naList.[2]
            (fillCells body 2 13 2 0 0 0).Text <- naList.[2]
            (fillCells body 2 14 1 0 1 0).Text <- naList.[1]
            (fillCells body 2 14 2 0 0 0).Text <- naList.[1]
            (fillCells body 2 15 1 0 1 0).Text <- naList.[0]
            (fillCells body 2 15 2 0 0 0).Text <- naList.[0]

            (fillCells body 2 17 0 0 0 0).Text <- mc12P20Id
            (fillCells body 2 18 0 0 0 0).Text <- p1000Id
            (fillCells body 2 19 0 0 0 0).Text <- p200Id
            (fillCells body 2 20 0 0 0 0).Text <- NA
            (fillCells body 2 21 0 0 0 0).Text <- NA
            (fillCells body 2 22 0 0 0 0).Text <- NA
            (fillCells body 2 23 0 0 0 0).Text <- NA
            (fillCells body 2 24 0 0 0 0).Text <- NA

            (fillCells body 2 17 1 0 0 0).Text <- pipetteCalibration
            (fillCells body 2 18 1 0 0 0).Text <- pipetteCalibration
            (fillCells body 2 19 1 0 0 0).Text <- pipetteCalibration
            (fillCells body 2 20 1 0 0 0).Text <- NA
            (fillCells body 2 21 1 0 0 0).Text <- NA
            (fillCells body 2 22 1 0 0 0).Text <- NA
            (fillCells body 2 23 1 0 0 0).Text <- NA
            (fillCells body 2 24 1 0 0 0).Text <- NA

            //Assigns reagent values to the last of the requested builds
            if param = lastLot then 
                (fillCells body 3 5 1 2 1 0).Text <- totalDigestRxns.ToString()
                (fillCells body 3 5 1 4 2 0).Text <- DEPC
                (fillCells body 3 5 1 6 3 0).Text <- neBuffer
                (fillCells body 3 5 1 8 2 0).Text <- cutterMix
                (fillCells body 3 5 1 10 3 0).Text <- pviiEnyme
                (fillCells body 3 5 1 13 1 0).Text <- totalDigestRxns.ToString()
                digestFootnote body inputParams
            else 
                digestFootnote body inputParams

            //Assigns footnotes and writes reference note if condition is met
            if inputParams.Length > 1 then 
                let lastLot = inputParams.Last()
                if param = lastLot then 
                    let restOfList = inputParams.[0..inputParams.Length - 2] |> String.concat ", "
                    let firstnote = "① Calculations include " + restOfList + "."
                    (fillCells body 3 11 0 2 0 0).Text <- firstnote
                   
                else
                    let note = "① Calculations are on " + lastLot + "."
                    (fillCells body 3 11 0 2 0 0).Text <- note

        
        // Saves document to temp folder for printing and then deletion
        let User = Environment.UserName
        let path = "C:/Users/" + User + "//AppData/Local/Temp/ "+param + " Reporter Ligation Batch Record" + ".docx"
        myDocument.SaveAs(path).Close()

