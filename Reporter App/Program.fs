// Learn more about F# at http://fsharp.org

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
open System.Diagnostics
open Sentry

// Paths where filled out Batch Records will be temporarily saved for printing
let pathsList (user : string) (param : string) = 
    [
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP Request Form" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Rp Ligation Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP Gel Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP Zag Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP reQC Batch Record" + ".docx"
    "C:/users/" + user + "/AppData/Local/Temp/ "+param + " Precipitation Form" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " RP Anneal Batch Record" + ".xlsx"
    "C:/users/" + user + "/AppData/Local/Temp/ "+param + " Normalization Precipitation Form" + ".docx"
    ]

// Function used to print Batch Records and then deleting the temp file
let printDocuments (path : string) =
    let printing = new Process()
    printing.StartInfo.FileName <- path
    printing.StartInfo.Verb <- "Print"
    printing.StartInfo.CreateNoWindow <- true
    printing.StartInfo.UseShellExecute <- true
    printing.EnableRaisingEvents <- true
    printing.Start() |> ignore
    printing.WaitForExit(10000)
    File.Delete(path)

[<EntryPoint>]
let main argv =

    // Reading in exel doc
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial  
    
    //User interface - input will take string as input
    Console.WriteLine "What RP lot will you work on?"
    let input = Console.ReadLine ()
    let inputSplit = input.Split(' ')
    let inputParams = [for i in inputSplit do i.ToUpper()]
    Console.WriteLine "Enter the number of the process you are conducting?"
    Console.WriteLine ( 
   "
    1 - Ligate 
    2 - Gel 
    3 - Zag
    4 - Re-Qc 
    5 - Precipitate 
    6 - Anneal
    7 - Normalize
    ")

    let processInput = Console.ReadLine ()

    // Error logging useing Sentry IO
    use __ = SentrySdk.Init ( fun o ->
           o.Dsn <-  "https://376aacc0025449a0bccf667abb1a4e39@o811036.ingest.sentry.io/5805230"
           o.SendDefaultPii <- true
           o.TracesSampleRate <- 1.0
           o.AttachStacktrace <- true
           o.ShutdownTimeout <- TimeSpan.FromSeconds 10.0 
           o.MaxBreadcrumbs <- 50 
           )

    SentrySdk.ConfigureScope(fun scope -> scope.SetTag("User Input", input) )
    SentrySdk.AddBreadcrumb(input)
    SentrySdk.ConfigureScope(fun newTag -> newTag.SetTag("Manufacturing_Process", processInput))

    
    
    // Reading in CS indentifying info
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial
    let fileInfo = new FileInfo("W:/Production/Reporter requests/CODESET TRACKERS/CodeSet Status.xlsx")
    use package = new ExcelPackage (fileInfo)
    let reporter = package.Workbook.Worksheets.["Upstream - RP"]

    // Reading in reagents from LIMS
    let reagentsInfo = new FileInfo("S:/ip/reagentsandtools.xlsx")
    use reagentsPackage = new ExcelPackage (reagentsInfo)
    let myTools = reagentsPackage.Workbook.Worksheets.["tools"]

    //Reading in excel file from path on oligo sheet
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial
    let oligoInfo = new FileInfo("W:/Production/Probe Oligos/REMP Files/_Re-Rack Files/Rerack Status.xlsx")
    use oligoPackage = new ExcelPackage (oligoInfo)
    let oligoStamps = oligoPackage.Workbook.Worksheets.["CodeSet Archive"]

    // Blank Batch Record paths
    let rqstform = "W:/program_files/FRM-M0207 Reporter Probe Synthesis Request.docx"
    let ligationForm = "W:/program_files/FRM-M0053-12_reporter probe ligation on dv1 backbone.docx"
    let gelForm = "W:/program_files/frm-m0218-04_ruo gel electrophoresis qc for reporter probe ligations.docx"
    let zagForm = "W:/program_files/FRM-10498-02_RP Ligation QC Using CE.docx"
    let reQcForm = "W:/program_files/frm-m0201-06_gel qc of rp ligations using samples prepared for analysis by ce.docx"
    let precipitationsForm = "W:/program_files/FRM-M0054-09 Ethanol Precipitation of Ligated and Pooled Reporters.docx"
    let annealsForm = "W:/program_files/FRM-M0055-09 Reporter Library Anneal Calculator.xlsx"
    let normalizationsForm = "W:/program_files/FRM-M0062-08 Normalization and Pre-incubation Calculator.xlsx"
    let normPrecipitationForm = "W:/program_files/FRM-M0057-11 Ethanol Precipitation of Annealed Reporters.docx"
    

    let user = Environment.UserName


    //Starts reading values of excel and stores it in "param"
    try
        
        if processInput.Equals("1", StringComparison.InvariantCultureIgnoreCase) then 
            Ligations.ligationStart inputParams rqstform ligationForm reporter oligoStamps myTools
               

        elif processInput.Equals("2", StringComparison.InvariantCultureIgnoreCase) then 
            GelQC.gelQcStart inputParams gelForm reporter myTools
         

        elif processInput.Equals("3", StringComparison.InvariantCultureIgnoreCase) then
            ZagQC.zagStart inputParams zagForm reporter myTools

        elif processInput.Equals("4", StringComparison.InvariantCultureIgnoreCase) then
            ReQC.reQcStart inputParams reQcForm reporter myTools


        elif processInput.Equals("5", StringComparison.InvariantCultureIgnoreCase) then
            EthanolPrecipitations.precipitationStart inputParams precipitationsForm rqstform reporter myTools

        elif processInput.Equals("6", StringComparison.InvariantCultureIgnoreCase) then 
            Anneals.annealStart inputParams annealsForm reporter myTools

        elif processInput.Equals("7", StringComparison.InvariantCultureIgnoreCase) then 
            Normalizations.normalizationStart inputParams normalizationsForm normPrecipitationForm reporter myTools

        else 
            Console.WriteLine "Invalid Process Entry..."


    
    with 
        | ex ->
            ex |> SentrySdk.CaptureException |> ignore

            for param in inputParams do 
                let docs = pathsList user param

                for each in docs do 
                    if File.Exists(each) then 
                        File.Delete(each)
                        Console.WriteLine "User Error..."
            
    try
        for param in inputParams do 
            let docs = pathsList user param
            for each in docs do 
                if File.Exists(each) then 
                    printDocuments each
    with 
        | _ -> 
            Console.WriteLine "Unable to print documents"

    0 // return an integer exit code
    
