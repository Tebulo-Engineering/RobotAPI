Imports System.Data
Imports System.IO
Imports System.Net.Security
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO
Imports RobotOM
Imports RobotOM.IRobotSelectionType

Module Program

    Sub Main(args As String())
        DumpTables()
    End Sub


    Sub DumpTables()

        Dim RobApp As New RobotApplication
        Dim RobModel As IRobotStructure = RobApp.Project.Structure
        Dim projPref As RobotProjectPreferences
        Dim RDmDetRes As IRDimDetailedRes
        Dim RDMAllRes As IRDimAllRes
        Dim BarCol_i As RobotBar
        projPref = RobApp.Project.Preferences
        Dim t As RobotTable
        Dim tf As RobotTableFrame
        Dim pathcwd As String
        Dim csvPath As String
        Dim oldFolderPath As String
        Dim rtdfolder As Object
        Dim rtdFiles As Array
        Dim strFiles As New List(Of Object)
        Dim Fullpath As String
        Dim FName As String
        Dim Filename As String
        Dim oFSB As Object
        Dim oFolder As Object
        Dim oFile As Object
        Dim i As Integer
        Dim Tags As Object
        Dim Tag As String
        Dim ManualInput As String
        Dim Automatic As Boolean
        Dim Valid As Boolean
        Dim Tag1 As Object
        Dim Tag2 As Object
        Dim Tag3 As Object
        Dim MemberVer As Boolean
        Dim Question As Boolean
        Dim BatchMove As Boolean
        Dim Archive As Boolean
        Dim SimpleCasesYN As Boolean
        Dim SimpleCases As String
        Dim RSelection As RobotSelection
        RSelection = RobApp.Project.Structure.Selections.Create(IRobotObjectType.I_OT_BAR)
        Dim BarCol As RobotBarCollection
        Dim fso : fso = CreateObject("Scripting.FileSystemObject")
        Dim RatioFailMessage As String
        Dim RatioFailLog() As String = Array.Empty(Of String)()


        MemberVer = UserYesNo("Do you want member verification results? (will take longer)") ' this bool to false if you do not want member verification

        Automatic = UserYesNo("Do you want to use the filename if no tag can be found? If you select No, a prompt will be given (Y/n)") 'Check if manual or automatic mode

        BatchMove = UserYesNo("Do you want to batch process? No will prompt for tags (Y/n)") 'Checks weather to do all files in the dir, or specific tags

        Archive = UserYesNo("Do you want to move all existing csv files to a archive directory? If plan to process any tags that have a csv file in the main dir, this is recommended")
        'Checks to move all the csv files. If they're already there, they get written into. however, if the program crashes during operation, saying no to this will process the files to #files
        SimpleCasesYN = UserYesNo("Are the simple cases anything other than '1to7 30to33' -> basic 7 and notional")
        If SimpleCasesYN Then
            Console.Write("Enter the simple cases in the style expected by RSA, e.g. 1to7 30to33" & vbCrLf)
            Console.Write("Messing this up will probably crash the programme, and there's no check written against it" & vbCrLf)
            SimpleCases = Console.ReadLine()
        Else
            SimpleCases = "1to7 30to33"
        End If


        'Move to correct folder
        get_proj()
            'Get the absolute path
            pathcwd = FileSystem.CurrentDirectory
            csvPath = pathcwd + "/csv/"


            Dim rtdPath
            rtdPath = pathcwd
            rtdFiles = Directory.GetFiles(rtdPath, "*.rtd")

            'Deleting all the .rtx files
            'rtx files are modification protection. If they're present the script crashes.
            'If someone is working in a file, do not delete the rtx file
            'This deletes ALL rtx files, even if you only want to process a few tags, maybe those ones are not even locked
            Dim filesToDelete As String() = Directory.GetFiles(pathcwd, "*.rtx")
            Dim DeleteYesNo As Boolean
            If filesToDelete.Length > 0 Then
                DeleteYesNo = UserYesNo("Found .rtx files, delete protections? ONLY DO THIS IF NO FILES ARE IN USE")
            End If
            If DeleteYesNo Then
                For Each item In filesToDelete
                    Try
                        System.IO.File.Delete(item)
                        Console.WriteLine($"Deleted file: {item}")
                    Catch ex As Exception
                        Console.WriteLine($"Failed to delete file: {item}. Reason: {ex.Message}")
                    End Try
                Next
            End If

        'Create a seperate folder for csv files if it does not yet exist
        If Not Directory.Exists(csvPath) Then
                'doesn't exist, so create the folder
                Directory.CreateDirectory(csvPath)
            Else
                ' Create a folder with a timestamp
                Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim oldFolderPathWithTimestamp As String = Path.Combine(csvPath, "archived_on_" & timestamp)
                Directory.CreateDirectory(oldFolderPathWithTimestamp)

                ' Get the files in the csvPath folder
                Dim files As String() = Directory.GetFiles(csvPath)

                ' Try to move each file to the timestamped folder
                If Archive Then
                    Try
                        For Each filePath As String In files
                            Dim currentFileName As String = Path.GetFileName(filePath)
                            Dim destinationPath As String = Path.Combine(oldFolderPathWithTimestamp, currentFileName)
                            File.Move(filePath, destinationPath)
                        Next
                    Catch
                        Console.WriteLine("Failed to archive old csv files, one might be open. If you contine it might not produce correct results")
                        Console.WriteLine("Press Enter to continue, press ctrl + C to abort")
                        Console.ReadLine()
                    End Try

                End If
            End If


            Dim TagsList As String() = Array.Empty(Of String)()
            Dim Tagx As New Regex("\d{4}\-\d{2}") 'regex of 4 numbers "-" and two numbers
            'Loop through each file

            If Not BatchMove Then
                'If only certain tags need to be processed, this user input requests and seperates them
                Dim TagsInput As String
                Console.WriteLine("Please input the tags you wish to process, seperated by a comma (xxxx-xx)")
                TagsInput = Console.ReadLine()
                TagsList = TagsInput.Split(","c)
                Console.WriteLine("Tags: " & TagsInput.Trim())
            End If


            If Not BatchMove Then
                'If only certain tags need to be processed, this user input requests and seperates them
                Dim TagsInput As String
                Console.WriteLine("Please input the tags you wish to process, seperated by a comma (xxxx-xx)")
                TagsInput = Console.ReadLine()
                TagsList = TagsInput.Split(","c)
                Console.WriteLine("Tags: " & TagsInput.Trim())
            End If

            For Each File In rtdFiles 'Where the magic happens, open every robot file and do the whole loop
                If LCase(Right(File, 4)) = ".rtd" Then
                    If LCase(Left(File, 5)) <> "robot" Then
                        Filename = rtdPath + File
                        If Tagx.IsMatch(File) Then 'Check if a tag can automatically be found
                            Tags = Tagx.Match(File)
                            If Tags.Captures.Count = 1 Then
                                Tag1 = Tags.Captures(0)
                                Tag = Tag1.Value
                            ElseIf Tags.Count = 2 Then
                                Tag1 = Tags(0)
                                Tag2 = Tags(1)
                                Tag = Tag1.Value + "+" + Tag2.Value
                            Else
                                Tag1 = Tags(0)
                                Tag2 = Tags(1)
                                Tag3 = Tags(2)
                                Tag = Tag1.Value + "+" + Tag2.Value + "+" + Tag3.Value 'Assign a maximum of 3 tags to the file, important to use a plus instead of an underscore, those break latex
                            End If
                        Else
                            If Automatic = True Then
                                Tag = File
                            Else
                                Console.Write("No tag found! Please input tag for file " + File) 'Give a prompt if no tags can be found
                                Tag = Console.ReadLine()
                            End If
                        End If

                        'The easiest way to select only certain tags.
                        'The use of GoTo is frowned upon and considered bad programming, sue me
                        'It skips to the end of the "for each", and therefore goes to the next tag, if it doesn't trigger it continues as normal
                        If Not BatchMove Then
                            If Not TagsList.Contains(Tag) Then
                                GoTo SkipTag
                            End If
                        End If

                        Console.Write("Opening " & Tag & vbCrLf)
                        RobApp.Project.Open(File)

                        RobApp.Project.ViewMngr.CurrentLayout = IRobotLayoutId.I_LI_MODEL_GEOMETRY
                        'Run the calculation if neccesary
                        If (RobApp.Project.Structure.Results.Available = False) Then
                            Try
                                Console.WriteLine($"Calculating {Tag}")
                                RobApp.Project.CalcEngine.Calculate()
                                RobApp.Project.Save()
                            Catch
                                RatioFailMessage = $"Failed to calculate results for {Tag}, please run calculation and try again..." & vbCrLf
                                Console.WriteLine(RatioFailMessage)
                                Console.WriteLine($"Skipping {Tag}...")
                                ReDim Preserve RatioFailLog(RatioFailLog.Length) 'Increase size of array by 1
                                RatioFailLog(RatioFailLog.Length - 1) = RatioFailMessage 'Add latest fail message to the list
                                GoTo SkipTag
                            End Try

                        End If

                        ' the forces to kN
                        Dim FU As RobotOM.RobotUnitData
                        FU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_FORCE)
                        Console.Write($"Before changing units, Force is in {FU.Name}" & vbCrLf)
                        FU.E = False
                        FU.Name = "kN"
                        FU.Precision = 2

                        ' the stresses to MPa
                        Dim SU As RobotOM.RobotUnitComplexData
                        SU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_STRESS)
                        Console.Write($"Before changing units, Stress is {SU.Name} and {SU.Name2}" & vbCrLf)
                        SU.E = False
                        SU.Name = "MPa"
                        'SU.Name2 = "mm2"
                        SU.Precision = 2

                        ' the moments to kNm
                        Dim MU As RobotOM.RobotUnitComplexData
                        MU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_MOMENT)
                        Console.Write($"Before changing units, Moment is {MU.Name} and {MU.Name2}" & vbCrLf)
                        MU.E = False
                        MU.Name = "kN"
                        MU.Name2 = "m"
                        MU.Precision = 2

                        ' the dimensions to mm
                        Dim DU As RobotOM.RobotUnitData
                        DU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_STRUCTURE_DIMENSION)
                        DU.E = False
                        DU.Name = "mm"
                        DU.Precision = 2

                        'Actually  the variables and refresh units
                        projPref.Units.Set(IRobotUnitType.I_UT_STRUCTURE_DIMENSION, DU)
                        projPref.Units.Set(IRobotUnitType.I_UT_FORCE, FU)
                        projPref.Units.Set(IRobotUnitType.I_UT_MOMENT, MU)
                        projPref.Units.Set(IRobotUnitType.I_UT_STRESS, SU)
                        projPref.Units.Refresh()
                        Console.Write($"The units are: {DU.Name}, {FU.Name}, {MU.Name} {MU.Name2} and {SU.Name}/{SU.Name2}" & vbCrLf)
                        Console.Write("Set units for " & Tag & vbCrLf)



                        Console.Write($"Before creating tables, there are {RobApp.Project.ViewMngr.TableCount} tables" & vbCrLf)
                        'Make sure the required tables are present, if these tables are already open, then duplicates are made, this doesn't matter
                        t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_BARS, IRobotTableDataType.I_TDT_VALUES)
                        t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_NODES, IRobotTableDataType.I_TDT_VALUES)
                        t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_LOADS, IRobotTableDataType.I_TDT_VALUES)
                        t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_REACTIONS, IRobotTableDataType.I_TDT_VALUES)
                        t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_STRESSES, IRobotTableDataType.I_TDT_VALUES)
                        t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_PROPERTIES, IRobotTableDataType.I_TDT_MEMBERS)
                        t.AddColumn(179)
                        t.AddColumn(180)
                        t.AddColumn(182)
                        Console.Write($"After creating tables, there are {RobApp.Project.ViewMngr.TableCount} tables" & vbCrLf)
                    'the columns added above are for the stresses, no idea how to find out which are which, good luck

                    'Dim ViewMngr As IRobotViewMngr
                    'Dim ActiveView As RobotView
                    'Dim CaseSel As RobotSelection
                    'Dim ActiveViewNumber As Long
                    'ViewMngr = RobApp.Project.ViewMngr
                    'For k = 1 To ViewMngr.ViewCount
                    '    If ViewMngr.GetView(k).Window.IsActive = -1 Then ActiveViewNumber = k : Exit For
                    'Next k
                    'ActiveView = ViewMngr.GetView(ActiveViewNumber)
                    'CaseSel = ActiveView.Selection.Get(IRobotObjectType.I_OT_CASE)
                    'CaseSel.FromText("Simple Cases")
                    'ViewMngr.Refresh()
                    'Console.WriteLine("Selected Simple Cases" & vbCrLf)

                    If MemberVer = True Then
                        BarCol = RobApp.Project.Structure.Bars.GetAll
                        Dim CaseCol = RobApp.Project.Structure.Cases.GetAll

                        Dim ratioName = csvPath & "Ratio" & Tag & ".csv"
                        Dim ratioNum As IO.StreamWriter = FileSystem.OpenTextFileWriter(ratioName, True)
                        ratioNum.Write("Member" & ";" & "Case" & ";" & "Ratio" & "Results" & vbCrLf)
                        Debug.Print("Member verification")
                        Dim RDMServer As IRDimServer
                        RDMServer = RobApp.Kernel.GetExtension("RDimServer")
                        RDMServer.Mode = IRDimServerMode.I_DSM_STEEL
                        Dim RDmEngine As IRDimCalcEngine
                        RDmEngine = RDMServer.CalculEngine

                        'the part below is optional, use it if you want to  calculation parameters by the code

                        Dim RDmCalPar As IRDimCalcParam
                        Dim RDmCalCnf As IRDimCalcConf

                        RDmCalPar = RDmEngine.GetCalcParam
                        RDmCalCnf = RDmEngine.GetCalcConf

                        Dim RdmStream As IRDimStream 'Data stream for thing parameters
                        RdmStream = RDMServer.Connection.GetStream
                        RdmStream.Clear()

                        'Calculate results for all sections

                        RdmStream.WriteText("all") ' member(s) selection
                        'Dim v = RDmCalPar.GetObjsList(IRDimCalcParamVerifType.I_DCPVT_MEMBERS_VERIF) 'members verification




                        Try
                            For i = 1 To BarCol.Count
                                RDmEngine.Solve(Nothing)
                                BarCol_i = BarCol.Get(i)
                                RDMAllRes = RDmEngine.Results
                                RDmDetRes = RDMAllRes.Get(BarCol_i.Number)
                                ratioNum.Write(BarCol_i.Number & ";" & RDmDetRes.GovernCaseName & ";" & RDmDetRes.Ratio & vbCrLf)
                                RDMAllRes = RDmEngine.Results
                                Debug.Print("Member verification finished")
                            Next i
                        Catch
                            RatioFailMessage = $"Failed to calculate member ver for {Tag}, please run calculation and try again..." & vbCrLf
                            Console.WriteLine(RatioFailMessage)
                            ReDim Preserve RatioFailLog(RatioFailLog.Length) 'Increase size of array by 1
                            RatioFailLog(RatioFailLog.Length - 1) = RatioFailMessage 'Add latest fail message to the list
                        End Try
                        ratioNum.Close()

                    End If

                    Dim nTables As Long
                    'Count the tables
                    nTables = RobApp.Project.ViewMngr.TableCount
                    'Console.Write($"{Tag} has {nTables} tables" & vbCrLf)
                    Dim tFilter As New Regex(":([0-9]+)")
                    For i = 1 To nTables

                        'Console.WriteLine($"Project is {RobApp.Project}, ViewMngr is {RobApp.Project.ViewMngr} and the Table is {RobApp.Project.ViewMngr.GetTable(i)} and the Recylce is {RobApp.Project.ViewMngr.CurrentLayout}" & vbCrLf)
                        tf = RobApp.Project.ViewMngr.GetTable(i) 'Read out the tables
                        FName = tf.Window.Caption
                        'Console.Write($"Table number {i} is called {FName}" & vbCrLf)
                        Dim spacepos = InStr(1, FName, " ")
                        If spacepos <> 0 Then
                            FName = Left(FName, spacepos) 'remove leading spaces
                        End If

                        Dim ntabs = tf.Count

                        For j = 1 To ntabs
                            tf.Get(j).Window.Activate()
                            t = tf.Get(j)
                            tf.Current = j
                            Dim tabname = tf.GetName(j)
                            Dim match As Match = tFilter.Match(FName)
                            If match.Success Then
                                'Console.WriteLine($"Table duplicate ({FName}) not printed")
                            Else
                                If Trim(FName) = "Reactions" Then 'Or Trim(FName) = "Stresses" Then 'We want the reactions and stresses envelope, it's more compact
                                    'Make sure that the correct simple cases are selected here!!!
                                    t.Select(I_ST_CASE, SimpleCases)
                                    If tabname = "Values" Then
                                        'DoEvents
                                        Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                        t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                        Console.WriteLine($"Writing tab {tabname} of table {FName} for tag {Tag}")
                                    End If
                                ElseIf Trim(FName) = "Loads" Then 'For loads, the table or text edition needs to be used, otherwise it'll likely be empty
                                    If tabname = "Text edition" Then
                                        'DoEvents
                                        Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                        t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                        Console.WriteLine($"Writing tab {tabname} of table {FName} for tag {Tag}")
                                    End If
                                ElseIf Trim(FName) = "Properties" Then 'Get the properties for each file, so they can be combined later
                                    If tabname = "Members" Then
                                        'DoEvents
                                        Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                        t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                        Console.WriteLine($"Writing tab {tabname} of table {FName} for tag {Tag}")
                                    End If
                                Else
                                    If tabname = "Values" Then 'Everything else just values
                                        'DoEvents
                                        Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                        t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                        Console.WriteLine($"Writing tab {tabname} of table {FName} for tag {Tag}")
                                    End If
                                End If
                            End If
                        Next j
                    Next i
                    RobApp.Project.Close()
                End If
            End If
SkipTag:
        Next

        RobApp = Nothing

        'get current folder
        Dim strPath
        strPath = csvPath
        Dim strDir = New DirectoryInfo(strPath)
        For Each fi In strDir.EnumerateFileSystemInfos()
            strFiles.Add(fi)
        Next
        'Loop through each file
        For Each Item In strFiles
            'Process file if it's .csv
            If LCase(Right(Item.Name, 4)) = ".csv" Then
                ProcessFile(strPath, Item.Name)
            End If
        Next

        Console.WriteLine("Dumping finished")
        For Each message In RatioFailLog
            Console.WriteLine(message)
        Next
        Console.WriteLine("Press Enter to continue")
        Console.ReadLine()


    End Sub

    Function ChangeLine(strLine, strFileName) As String
        If Trim(strLine) <> "" Then
            If Left(strFileName, 9) = "Reactions" Then
                Dim Tag1 = Right(strFileName, Len(strFileName) - 9)
                Dim Tag2 = Left(Tag1, Len(Tag1) - 4)
                strLine = Replace(LCase(Tag2) + ".", "csv", "") & strLine
            End If
            'edit line content, replacing the decimal and column seperators, so latex can understand it
            strLine = Replace(strLine, ",", ".")
            strLine = Replace(strLine, ";", ",")
        End If
        If Left(Trim(strLine), 5) = "Point" Then
            strLine = Right(Trim(strLine), Len(Trim(strLine)) - 5)
        End If
        If Left(Trim(strLine), 1) = "," Then
            strLine = Right(Trim(strLine), Len(Trim(strLine)) - 1)
        End If
        ChangeLine = strLine
    End Function

    Function ProcessFile(strPath, strFileName)
        Const ForReading = 1, ForWriting = 2
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        'open file
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim F_r = FileSystem.OpenTextFileReader(strPath & "\" & strFileName)
        'Dim F = fso.OpenTextFile(strPath & "\" & strFileName, ForReading, True, TristateUseDefault)
        'read first line
        Dim strLine = F_r.ReadLine()
        'check for delimiter
        If InStr(strLine, ";") <> 0 Then
            'read file content line by line
            strLine = ChangeLine(strLine, "")
            Dim StrFile = strLine
            While Not F_r.EndOfStream ' while we are not finished reading through the file
                strLine = F_r.ReadLine
                strLine = ChangeLine(strLine, strFileName)
                StrFile = StrFile & vbCrLf & strLine
            End While
            F_r.Close()
            'save file
            'F = fso.OpenTextFile(strPath & "\" & "#" & strFileName, ForWriting, True)
            Dim F_w = FileSystem.OpenTextFileWriter(strPath & "\" & "#" & strFileName, True)
            F_w.WriteLine(StrFile)
            F_w.Close()
        Else
            F_r.Close()
        End If
    End Function

    Function UserYesNo(Question As String) As Boolean
        Dim input As String = String.Empty
        Dim stringYN As String = "[y|n]"
        Dim r As New Regex(stringYN)
        Do While input = String.Empty
            Console.WriteLine(Question)
            input = Console.ReadLine().ToLower()
            If Not r.IsMatch(input) Then input = String.Empty
        Loop
        Return input.StartsWith("y")
    End Function

    Function get_proj()
        Dim J As Object = Directory.GetDirectories("J:\\")
        Dim proj_nr As String = String.Empty
        Dim Question As Boolean = True
        Do While Question
            While proj_nr.Length <> 8
                Console.WriteLine("Please input project number (xx-xxxxx, 7 digits)")
                proj_nr = Console.ReadLine()
            End While
            Dim ye_nr As String = "20" & proj_nr.Substring(0, 2)
            If Not J Is Nothing Then
                For Each Dir_J In J
                    If Dir_J.Contains(ye_nr) Then
                        Console.WriteLine("year found: " & Dir_J)
                        ChDir(Dir_J)
                        Exit For
                    End If
                Next
            End If

            Dim CurrDir As String = Directory.GetCurrentDirectory()
            Dim Proj As Object = Directory.GetDirectories(CurrDir)
            If Not Proj Is Nothing Then
                For Each Dir_proj In Proj
                    If Dir_proj.Contains(proj_nr) Then
                        Dim Dir_calc = Dir_proj & "\\500 In bewerking\\540 Berekening"
                        Console.WriteLine("Directory found: " & Dir_calc)
                        ChDir(Dir_calc)
                        Question = False
                        Exit For
                    End If
                Next
            End If
        Loop
    End Function

End Module











