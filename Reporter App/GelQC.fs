module GelQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open GelHelpers


let gelQcStart(inputParams : string list) (gelForm : string) (reporter : ExcelWorksheet)(myTools : ExcelWorksheet) =

    let user =  Environment.UserName 

    for param in inputParams do 

        let negative = gelsListFunction "rpgelqc" myTools 2
        let negativeExp = gelsListFunction "rpgelqc" myTools 3
        let hyperLadder = gelsListFunction "rpgelqc" myTools 4
        let hyperLadderExp = gelsListFunction "rpgelqc" myTools 5

        let docArray = File.ReadAllBytes(gelForm)
        use _copyDoc = new MemoryStream(docArray)
        use gelDocument = WordprocessingDocument.Open(_copyDoc, true)
        let gelBody = gelDocument.MainDocumentPart.Document.Body

        let lot, csName, species, customer, geneNumber, scale = (codesetIdentifiers param reporter)

        let plateCount = System.Math.Floor((geneNumber|>float) / 96.0)
        
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

        (gelsCsInfoHeader gelBody 2 3).Text <- lot + " " + csName
        (gelsCsInfoHeader gelBody 2 11).Text <- totalGenesToGel
        (gelsCsInfoHeader gelBody 2 16).Text <- scale.ToString()
        (gelsTableFiller gelBody 0 1 3 0).Text <- hyperLadder
        (gelsTableFiller gelBody 0 1 4 0).Text <- hyperLadderExp
        (gelsTableFiller gelBody 0 2 3 0).Text <- negative
        (gelsTableFiller gelBody 0 2 4 0).Text <- negativeExp

        let gelBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP Gel Batch Record" + ".docx"
        gelDocument.SaveAs(gelBatchRecordPath).Close() |> ignore


