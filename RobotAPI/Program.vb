Imports System.IO
Imports System.Net.Security
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports RobotOM

Module Program

    Sub Main(args As String())
        DumpTables()
    End Sub


    Sub DumpTables()

        Dim RobApp As New RobotApplication
        Dim projPref As RobotProjectPreferences
        projPref = RobApp.Project.Preferences
        Dim t As RobotTable
        Dim tf As RobotTableFrame
        Dim path As String
        Dim csvPath As String
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
        Dim RSelection As RobotSelection
        RSelection = RobApp.Project.Structure.Selections.Create(IRobotObjectType.I_OT_BAR)
        Dim BarCol As RobotBarCollection
        Dim fso : fso = CreateObject("Scripting.FileSystemObject")

        ' this bool to false if you do not want member verification
        MemberVer = True

        Automatic = True 'UserYesNo("Do you want to use the filename if no tag can be found? If you select No, a prompt will be given (Y/n)") 'Check if manual or automatic mode


        'Move to correct folder
        get_proj()
        'Get the absolute path
        path = FileSystem.CurrentDirectory
        csvPath = path + "/csv/"


        Dim rtdPath
        rtdPath = path
        rtdFiles = Directory.GetFiles(rtdPath)

        'Create a seperate folder for csv files if it does not yet exist
        If Not Directory.Exists(csvPath) Then
            'doesn't exist, so create the folder
            Directory.CreateDirectory(csvPath)
        End If

        Dim Tagx As New Regex("\d{4}\-\d{2}") 'regex of 4 numbers "-" and two numbers
        'RegExTag.Global = True 'Check for more than one instance
        'RegExTag.IgnoreCase = True
        'Loop through each file
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
                    Console.Write(Tag & vbCrLf)
                    RobApp.Project.Open(File)


                    'Run the calculation if neccesary
                    If (RobApp.Project.Structure.Results.Available = False) Then
                        RobApp.Project.CalcEngine.Calculate()
                        'RobApp.Project.CalcEngine.SaveResultsInExternalFile
                    End If



                    'Make sure the required tables are present, if these tables are already open, then duplicates are made, this doesn't matter
                    t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_BARS, IRobotTableDataType.I_TDT_VALUES)
                    t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_NODES, IRobotTableDataType.I_TDT_VALUES)
                    t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_LOADS, IRobotTableDataType.I_TDT_VALUES)
                    t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_REACTIONS, IRobotTableDataType.I_TDT_VALUES)
                    t = RobApp.Project.ViewMngr.CreateTable(IRobotTableType.I_TT_STRESSES, IRobotTableDataType.I_TDT_VALUES)
                    t.AddColumn(179)
                    t.AddColumn(180)
                    t.AddColumn(182)

                    ' the forces to kN
                    Dim FU As RobotOM.RobotUnitData
                    FU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_FORCE)
                    FU.E = False
                    FU.Name = "kN"
                    FU.Precision = 2

                    ' the stresses to N/mm2
                    Dim SU As RobotOM.RobotUnitData
                    SU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_STRESS)
                    SU.E = False
                    SU.Name = "MPa"
                    SU.Precision = 2

                    ' the moments to kNm
                    Dim MU As RobotOM.RobotUnitData
                    MU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_MOMENT)
                    MU.E = False
                    MU.Name = "kNm"
                    MU.Precision = 2

                    ' the dimensions to mm
                    Dim DU As RobotOM.RobotUnitData
                    DU = projPref.Units.Get(RobotOM.IRobotUnitType.I_UT_STRUCTURE_DIMENSION)
                    DU.E = False
                    DU.Name = "mm"
                    DU.Precision = 0

                    'Actually  the variables and refresh units
                    projPref.Units.Set(IRobotUnitType.I_UT_FORCE, FU)
                    projPref.Units.Set(IRobotUnitType.I_UT_MOMENT, MU)
                    projPref.Units.Set(IRobotUnitType.I_UT_STRUCTURE_DIMENSION, DU)
                    projPref.Units.Set(IRobotUnitType.I_UT_STRESS, SU)
                    projPref.Units.Refresh()

                    Dim nTables As Long
                    'Count the tables
                    nTables = RobApp.Project.ViewMngr.TableCount



                    If MemberVer = True Then
                        BarCol = RobApp.Project.Structure.Bars.GetAll
                        Dim CaseCol = RobApp.Project.Structure.Cases.GetAll

                        Dim ratioName = csvPath & "Ratio" & Tag & ".csv"
                        Dim ratioNum As IO.StreamWriter = FileSystem.OpenTextFileWriter(ratioName, True)
                        ratioNum.Write("Bar" & ";" & "Case" & ";" & "Ratio")
                        Debug.Print("Member verification")
                        For i = 1 To BarCol.Count
                            For j = 1 To CaseCol.Count

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

                                Dim RdmStream As IRDimStream 'Data stream for ting parameters
                                RdmStream = RDMServer.Connection.GetStream
                                RdmStream.Clear()

                                'Calculate results for all sections
                                Dim aaa = BarCol.Get(i) 'This is the start of the problems, here I need to get BarCol.Get(i).Number, but it cannot find Number, in any way shape or form, same for CaseCol & other properties
                                Dim bbb = CaseCol.Get(j)
                                Console.WriteLine(aaa)
                                RdmStream.WriteText(aaa) ' member(s) selection
                                'Dim v = RDmCalPar.GetObjsList(IRDimCalcParamVerifType.I_DCPVT_MEMBERS_VERIF) 'members verification
                                RDmCalPar.SetObjsList(IRDimCalcParamVerifType.I_DCPVT_MEMBERS_VERIF, RdmStream)
                                RDmCalPar.SetLimitState(IRDimCalcParamLimitStateType.I_DCPLST_ULTIMATE, 1) ' Limit State
                                RdmStream.Clear()
                                RdmStream.WriteText(bbb.ToString) ' Load Case(s)
                                RDmCalPar.GetLoadsList(RdmStream)
                                RDmEngine.GetCalcConf()
                                RDmEngine.GetCalcParam()

                                'end of calclulation parameter tings

                                RDmEngine.Solve(Nothing)

                                Dim RDmDetRes As IRDimDetailedRes
                                Dim RDMAllRes As IRDimAllRes
                                If InStr(1, LCase(bbb.ToString), "sls") = 0 Then 'We do not want SLS, that does not work
                                    'Debug.Print "About to write the results of bar: " & BarCol.Get(i).Number & " case: " & CaseCol.Get(j).Name
                                    RDMAllRes = RDmEngine.Results
                                    RDmDetRes = RDMAllRes.Get(aaa) 'Hier gaat het nu fout: System.InvalidCastException: 'Conversion from type 'IRobotBar' to type 'Integer' is not valid.'
                                    ratioNum.Write(aaa.ToString & ";" & RDmDetRes.GovernCaseName & ";" & RDmDetRes.Ratio)

                                End If
                                'printing the results to csv
                            Next j
                        Next i
                        ratioNum.Close()
                        Debug.Print("Member verification finished")
                    End If

                    For i = 1 To nTables

                        tf = RobApp.Project.ViewMngr.GetTable(i) 'Read out the tables
                        FName = tf.Window.Caption
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
                            If Trim(FName) = "Reactions" Then 'Or Trim(FName) = "Stresses" Then 'We want the reactions and stresses envelope, it's more compact
                                If tabname = "Envelope" Then
                                    'DoEvents
                                    Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                    t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                End If
                            ElseIf Trim(FName) = "Loads" Then 'For loads, the table or text edition needs to be used, otherwise it'll likely be empty
                                If tabname = "Text edition" Then
                                    'DoEvents
                                    Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                    t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                End If
                            Else
                                If tabname = "Values" Then 'Everything else just values
                                    'DoEvents
                                    Fullpath = csvPath + Trim(FName) + Tag + ".csv"
                                    t.Printable.SaveToFile(Fullpath, IRobotOutputFileFormat.I_OFF_TEXT)
                                End If
                            End If
                        Next j
                    Next i
                    RobApp.Project.Close()
                End If
            End If
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











