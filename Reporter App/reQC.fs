module ReQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open reQCHelpers

// Function used to fill out re-QC Batch Record
let reQcStart(inputParams : string list) (reQcForm : string) (reporter : ExcelWorksheet)(myTools : ExcelWorksheet) =
    let user =  Environment.UserName 
    for param in inputParams do 

        // User interface
        Console.WriteLine ("How many probes are being reQC-ed for " + param)
        let reQcGeneNumber = Console.ReadLine()

        // Reads in reagent lot info from LIMS
        let negative = gelsListFunction "rpgelqc" myTools 2
        let negativeExp = gelsListFunction "rpgelqc" myTools 3
        let hyperLadder = gelsListFunction "rpgelqc" myTools 4
        let hyperLadderExp = gelsListFunction "rpgelqc" myTools 5

        // Reads in blank Batch Record form
        let docArray = File.ReadAllBytes(reQcForm)
        use _copyDoc = new MemoryStream(docArray)
        use reQcDocument = WordprocessingDocument.Open(_copyDoc, true)
        let reQCBody = reQcDocument.MainDocumentPart.Document.Body

        // Assigns CS identifiers from LIMS
        let lot, csName, geneNumber, scale = (codesetIdentifiers param reporter)

        // Fills out fields in Batch Record
        (gelsCsInfoHeader reQCBody 2 6).Text <- lot + " " + csName
        (gelsCsInfoHeader reQCBody 2 13).Text <- reQcGeneNumber
        (gelsCsInfoHeader reQCBody 2 17).Text <- geneNumber.ToString()
        (gelsCsInfoHeader reQCBody 2 22).Text <- scale.ToString()
        (gelsTableFiller reQCBody 0 1 3 0).Text <- hyperLadder
        (gelsTableFiller reQCBody 0 1 4 0).Text <- hyperLadderExp
        (gelsTableFiller reQCBody 0 2 3 0).Text <- negative
        (gelsTableFiller reQCBody 0 2 4 0).Text <- negativeExp

        // Saves form in temp folder for printing and deletion
        let reQcBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP reQC Batch Record" + ".docx"
        reQcDocument.SaveAs(reQcBatchRecordPath).Close() |> ignore

