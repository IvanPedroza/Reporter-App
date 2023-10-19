module Anneals

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
open System.Diagnostics
open AnnealHelpers

// Function used to fill out annealing Batch Records
let annealStart (inputParams : string list) (annealsForm : string) (reporter : ExcelWorksheet) (myTools : ExcelWorksheet) =
    // Gets logged user for use in path and error logging
    let user = Environment.UserName
    // User interface
    Console.WriteLine "Which bench space are you using?"
    let benchInput = Console.ReadLine ()
    let reagentsInput = "anneals"
    let heatBlockInput = "heatblocks"



    // Starting point for manipulating sheet info and storing it in "param"
    for param in inputParams do

        // Reads in equipments IDs for bench specified from LIMS
        let pipetteCalibration = listFunction benchInput myTools 2
        let p1000 = listFunction benchInput myTools 3
        let p1000Id = listFunction benchInput myTools 4
        let p200 = listFunction benchInput myTools 5
        let p200Id = listFunction benchInput myTools 6
        let p100 = listFunction benchInput myTools 7
        let p100Id = listFunction benchInput myTools 8
        let p20 = listFunction benchInput myTools 9
        let p20Id = listFunction benchInput myTools 10
        let p10 = listFunction benchInput myTools 11
        let p10Id = listFunction benchInput myTools 12
        let p2 = listFunction benchInput myTools 13
        let p2Id = listFunction benchInput myTools 14
        let p2000 = listFunction benchInput myTools 15
        let p2000Id = listFunction benchInput myTools 16
        let mc8P20Id = listFunction benchInput myTools 18
        let mc12P20Id = listFunction benchInput myTools 22
        let mc8P200Id = listFunction benchInput myTools 20
        let mc12P200Id = listFunction benchInput myTools 24

        // Reads in reagent lot info from LIMS
        let trisLot = listFunction reagentsInput myTools 2
        let trisExp = listFunction reagentsInput myTools 3
        let trisUbd = listFunction reagentsInput myTools 4
        let waterLot = listFunction reagentsInput myTools 5
        let waterExp = listFunction reagentsInput myTools 6
        let waterUbd = listFunction reagentsInput myTools 7
        let sspeLot = listFunction reagentsInput myTools 8
        let sspeExp = listFunction reagentsInput myTools 9
        let sspeUbd = listFunction reagentsInput myTools 10
        let ethanolLot = listFunction reagentsInput myTools 11
        let ethanolUbd = listFunction reagentsInput myTools 12
        let dv1Lot = listFunction reagentsInput myTools 13
        let dv1Exp = listFunction reagentsInput myTools 14
        let h37Id = listFunction heatBlockInput myTools 2
        let h37Cal = listFunction heatBlockInput myTools 3
        let h45Id = listFunction heatBlockInput myTools 4
        let h45Cal = listFunction heatBlockInput myTools 5
        let h65Id = listFunction heatBlockInput myTools 6
        let h65Cal = listFunction heatBlockInput myTools 7
        let h75Id = listFunction heatBlockInput myTools 8
        let h75Cal = listFunction heatBlockInput myTools 9

      
        // Reads in blank Batch Record template 
        let readExcelBytes = File.ReadAllBytes(annealsForm)
        let excelStream = new MemoryStream(readExcelBytes)
        let formPackage = new ExcelPackage(excelStream)
        let annealSheet = formPackage.Workbook.Worksheets.["FRM-M0055"]
        

        // Takes value of each cell of the row in which the input lies and stores it for use in filling out Word Doc
        let lot = listFunction param reporter 1 
        let csName = listFunction param reporter 2 
        let geneNumber = listFunction param reporter 5  |> float
        let scale = listFunction param reporter 6  |> float
           

        // Fills out Batch Record at specified cells
        annealSheet.Cells.[5, 2].Value <- lot + " " + csName
        annealSheet.Cells.[5, 5].Value <- geneNumber
        annealSheet.Cells.[5, 7].Value <- scale
        annealSheet.Cells.[8, 4].Value <- trisLot
        annealSheet.Cells.[8, 6].Value <- trisExp
        annealSheet.Cells.[9, 6].Value <- "Use by date: " + trisUbd
        annealSheet.Cells.[10, 4].Value <- waterLot
        annealSheet.Cells.[10, 6].Value <- waterExp
        annealSheet.Cells.[11, 6].Value <- "Use by date: " + waterUbd
        annealSheet.Cells.[12, 4].Value <- sspeLot
        annealSheet.Cells.[12, 6].Value <- sspeExp
        annealSheet.Cells.[13, 6].Value <- "Use by date: " + sspeUbd
        annealSheet.Cells.[14, 4].Value <- ethanolLot
        annealSheet.Cells.[14, 6].Value <- "Use by date: " + ethanolUbd
        annealSheet.Cells.[15, 4].Value <- dv1Lot
        annealSheet.Cells.[15, 6].Value <- dv1Exp
        annealSheet.Cells.[17, 3].Value <- h37Id
        annealSheet.Cells.[17, 6].Value <- h37Cal
        annealSheet.Cells.[18, 3].Value <- h45Id
        annealSheet.Cells.[18, 6].Value <- h45Cal
        annealSheet.Cells.[19, 3].Value <- h65Id
        annealSheet.Cells.[19, 6].Value <- h65Cal
        annealSheet.Cells.[20, 3].Value <- "CHP- " + h75Id 
        annealSheet.Cells.[20, 6].Value <- h75Cal
        annealSheet.Cells.[22, 3].Value <- p10Id
        annealSheet.Cells.[22, 6].Value <- pipetteCalibration
        annealSheet.Cells.[23, 3].Value <- p20Id
        annealSheet.Cells.[23, 6].Value <- pipetteCalibration
        annealSheet.Cells.[24, 3].Value <- p100Id
        annealSheet.Cells.[24, 6].Value <- pipetteCalibration
        annealSheet.Cells.[25, 3].Value <- p200Id
        annealSheet.Cells.[25, 6].Value <- pipetteCalibration
        annealSheet.Cells.[26, 3].Value <- p1000Id
        annealSheet.Cells.[26, 6].Value <- pipetteCalibration
        annealSheet.Cells.[27, 3].Value <- p2000Id
        annealSheet.Cells.[27, 6].Value <- pipetteCalibration


        // Saves filled out Batch Record for printing and deletion
        let user = Environment.UserName
        let path = "C:/Users/" + user + "//AppData/Local/Temp/ "+param + " RP Anneal Batch Record" + ".xlsx"
