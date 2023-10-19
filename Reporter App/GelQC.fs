module GelQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open GelHelpers

// Function for filling out Gel QC Batch Record
let gelQcStart(inputParams : string list) (gelForm : string) (reporter : ExcelWorksheet)(myTools : ExcelWorksheet) =
    // Gets logged user for path and error logging
    let user =  Environment.UserName 

    // Cycles through lots being QC-ed and fills out a Batch Record for each of them
    for param in inputParams do 

        // Gets reagent information from LIMS
        let negative = gelsListFunction "rpgelqc" myTools 2
        let negativeExp = gelsListFunction "rpgelqc" myTools 3
        let hyperLadder = gelsListFunction "rpgelqc" myTools 4
        let hyperLadderExp = gelsListFunction "rpgelqc" myTools 5
        
        // Reads in Batch Record templete
        let docArray = File.ReadAllBytes(gelForm)
        use _copyDoc = new MemoryStream(docArray)
        use gelDocument = WordprocessingDocument.Open(_copyDoc, true)
        let gelBody = gelDocument.MainDocumentPart.Document.Body

        // Reads in CS identifying info from LIMS
        let lot, csName, species, customer, geneNumber, scale = (codesetIdentifiers param reporter)

        // Determines how many plates the CS is made up of
        let plateCount = System.Math.Floor((geneNumber|>float) / 96.0)

        // Calculates number of wells that contain samples
        let totalGenesToGel =
            if param.EndsWith("RW", StringComparison.InvariantCultureIgnoreCase) then 
                Console.WriteLine ("How many probes are you geling for " + param + "?")
                let genesToGel = Console.ReadLine ()
                genesToGel + "/" + geneNumber
                   
            else 
                if plateCount > 1.0 then 
                    let unGeledGenes = 96.0 * plateCount
                    let genesToGel = (geneNumber|>float) - unGeledGenes
                    genesToGel.ToString() + "/" + geneNumber
                else
                    geneNumber
                    
        // Fills out cs and reagent lot information
        (gelsCsInfoHeader gelBody 2 3).Text <- lot + " " + csName
        (gelsCsInfoHeader gelBody 2 11).Text <- totalGenesToGel
        (gelsCsInfoHeader gelBody 2 16).Text <- scale.ToString()
        (gelsTableFiller gelBody 0 1 3 0).Text <- hyperLadder
        (gelsTableFiller gelBody 0 1 4 0).Text <- hyperLadderExp
        (gelsTableFiller gelBody 0 2 3 0).Text <- negative
        (gelsTableFiller gelBody 0 2 4 0).Text <- negativeExp

        // Saves filled out Batch Record to temp directory for printing and deleting 
        let gelBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP Gel Batch Record" + ".docx"
        gelDocument.SaveAs(gelBatchRecordPath).Close() |> ignore


