using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using LogAnalyzer.Helper;

namespace FirstProject.Helper
{
    //public class CommandSummary
    //{
    //    public string CommandName { get; set; }
    //    public string Timestamp { get; set; }
    //    public string TimeTaken { get; set; }
    //    public string Status { get; set; }
    //}

    //public class DeviceCommandSummary : CommandSummary
    //{
    //    public string BuildVersion { get; set; }
    //    public string LocalRadioID {  get; set; }
    //    public string ModuleID { get; set; }
    //    public string DcwVersion { get; set; }
    //    public string ModuleFirmwareVersion { get; set; }
    //    public string MeterFirmwareVersion { get; set; }
    //    public string MajorMinorSoftID { get; set; }
    //    public string Comments { get; set; }
    //    public string TotalTimeTakenPerRun { get; set; }
    //}

    internal class LogConverter :CommandSummary
    {
        public static void ConvertLogToExcel(string logPath, string excelPath)
        {
            var logLines = File.ReadAllLines(logPath);
            string currentCommand = null;
            string currentTimestamp = null;
            string currentBuildVersion = null;
            string currentLocalRadioId= null;
            string currentModuleID = null;
            string totalTime = null;
            string currentStatus = "F";
            string prevAddedCommand = null;
            string lastLocalRadioID = null;
            string currentComment = null;
            string currentTimePerRun = null;
            int FailCount = 0;
            int retry = 0;
            string retryCommandName=null;
            string Error1 = "80 52 failed to complete.";
            Boolean E1 = false;
            int E1Count = 0,TotalRun=0;
            string Error2 = "TunerData Command failed to complete.";
            Boolean E2 = false;
            int E2Count = 0;
            string Error3 = "Connection with module failed - EB Timeout Error";
            string Error6 = "Failed to connect the endpoint";



            Boolean E3 = false;
            int E3Count = 0;
            int E4Count = 0;
            int E5Count = 0;
            int E6Count = 0;

            List<DeviceCommandSummary> commandSummaries = new List<DeviceCommandSummary>();

            foreach (var line in logLines)
            {

                if (line.Contains("Notify to UI for EndPoint"))
                {

                    Match lanMatch = Regex.Match(line, @"(?<=Notify to UI for EndPoint -\s*)\w+");
                    string foundID = lanMatch.Value;
                    if (lanMatch.Length != 8)
                    {
                        var decodingTimestamp = line.Substring(11, 12);
                        var commandEntry = commandSummaries.Find(c => c.CommandName == "ConnectLocalRadio" && string.IsNullOrEmpty(c.LocalRadioID));
                        commandEntry.Comments = "Invalid LocalRadioID";
                        E4Count++;
                        FailCount++;
                        commandEntry.LocalRadioID = foundID;
                        commandEntry.TimeTaken = CalculateTotalTime(commandEntry.Timestamp.Substring(11, 12), decodingTimestamp);
                        commandEntry.Status = "F";
                        continue;
                    

                    }

                    if (lanMatch.Success)
                    {
                        var commandEntry = commandSummaries.Find(c => c.CommandName == "ConnectLocalRadio" && string.IsNullOrEmpty(c.LocalRadioID));

                        if (commandEntry != null)
                        {
                            commandEntry.LocalRadioID = foundID;
                            lastLocalRadioID = foundID; // Update for future entries
                        }
                        else if (foundID != lastLocalRadioID)
                        {
                            currentModuleID = foundID;
                            var firstCmd = commandSummaries.FindLast(c => c.CommandName == "ConnectLocalRadio");
                            var lastCmd = commandSummaries.LastOrDefault();

                            foreach (var cmd in commandSummaries.Where(c => string.IsNullOrEmpty(c.LocalRadioID)))
                            {
                             
                                // Set LocalRadioID from last known, if not already set
                                if (string.IsNullOrEmpty(cmd.LocalRadioID))
                                    cmd.LocalRadioID = lastLocalRadioID;
                            }
                            if (firstCmd != null && lastCmd != null)
                            {
                                // Get the indexes of firstCmd and lastCmd
                                int startIndex = commandSummaries.IndexOf(firstCmd);
                                int endIndex = commandSummaries.IndexOf(lastCmd);

                                // If both indexes are valid and the start is before the end
                                if (startIndex != -1 && endIndex != -1 && endIndex >= startIndex)
                                {
                                    // Loop through the range from firstCmd to lastCmd
                                    for (int k = startIndex; k <= endIndex; k++)
                                    {
                                        var cmd = commandSummaries[k];

                                        if (string.IsNullOrEmpty(cmd.ModuleID))
                                            cmd.ModuleID = currentModuleID;
                                    }
                                }
                            }
                        }
                    }


                }


                if (line.Contains("Added command:"))
                {
                    
                    currentCommand = ExtractAddedCommand(line);
                    if(currentCommand == "DisconnectLocalRadio")
                    {
                        var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");
                        var lastCmd = commandSummaries.LastOrDefault();
                        string startTime = firstCmd.Timestamp.Substring(11, 12);
                        string EndTime = line.Substring(11, 12);
                        currentTimePerRun = CalculateTotalTime(startTime, EndTime);
                        
                        

                        if (firstCmd != null && lastCmd != null )
                        {

                            // Find index range
                            int startIndex = commandSummaries.LastIndexOf(firstCmd);
                            int endIndex = commandSummaries.IndexOf(lastCmd);

                            // Apply comment only in that range
                            for (int j = startIndex; j <= endIndex; j++)
                            {
                                if (string.IsNullOrEmpty(commandSummaries[j].TotalTimeTakenPerRun))
                                {
                                    commandSummaries[j].TotalTimeTakenPerRun = currentTimePerRun;
                                }
                            }
                        }
                        continue;

                    }
                    
                    currentTimestamp = line.Substring(0, 23);
                    if(currentCommand== "EndPointDisconnect" |
                        currentCommand == "PanaCommandHandler" | currentCommand=="Device8052" | currentCommand == "TunerData" )
                    {
                        continue;
                    }
                    
                    // Extract build version
                    Match match = Regex.Match(line, @"-\s*([\d]+\.[\d]+\.[\d]+\.[\d]+)\s*-");
                    currentBuildVersion = match.Success ? match.Groups[1].Value : null;

                    commandSummaries.Add(new DeviceCommandSummary
                    {
                        BuildVersion = currentBuildVersion,
            
                        CommandName = currentCommand,
                        Timestamp = currentTimestamp,
                        TimeTaken = "NA",  // Empty for now
                        Status = "F"     // Default status Failed
                    });
                }

                if(line.Contains("Building SetTimeCommand"))
                {
                    // Extract build version
                    Match match = Regex.Match(line, @"-\s*([\d]+\.[\d]+\.[\d]+\.[\d]+)\s*-");
                    currentBuildVersion = match.Success ? match.Groups[1].Value : null;
                    currentTimestamp = line.Substring(0, 23);
                    commandSummaries.Add(new DeviceCommandSummary
                    {
                        BuildVersion = currentBuildVersion,
                        CommandName = "SetTimeCommand",
                        Timestamp = currentTimestamp,
                        TimeTaken = "NA",  // Empty for now
                        Status = "F"     // Default status Failed
                    });

                }
                if (line.Contains("Command Executed:"))
                {
                    currentCommand = ExtractExecutedCommand(line);
                    var lastCmd = commandSummaries.LastOrDefault();
                   if(lastCmd!=null && lastCmd.CommandName == "ManufacturingData")
                    {
                        commandSummaries.Remove(lastCmd);
                    }
                   
                    if (currentCommand == "ManufacturingData" )
                    {
                        currentTimestamp = line.Substring(0, 23);
                        Match match = Regex.Match(line, @"-\s*(\d+\.\d+\.\d+\.\d+)\s*-");
                        currentBuildVersion = match.Success ? match.Groups[1].Value : null;

                        commandSummaries.Add(new DeviceCommandSummary
                        {
                            BuildVersion = currentBuildVersion,

                            CommandName = currentCommand,
                            Timestamp = currentTimestamp,
                            TimeTaken = "NA",  // Empty for now
                            Status = "F"     // Default status Failed
                        });


                    }
                }
                if(line.Contains("Message Payload:  (Hex: 00,00)") )
                {
                    var lastCmd = commandSummaries.LastOrDefault();
                    var decodingTimestamp = line.Substring(11, 12);
                    if (lastCmd != null && lastCmd.CommandName == "AcquireSingleChannelTargetWanSearch")
                    {
                        lastCmd.Status = "S";
                        lastCmd.TimeTaken = CalculateTotalTime(lastCmd.Timestamp.Substring(11,12), decodingTimestamp);
                    }

                }
                if (line.Contains("Message Payload: (Hex: 10, 00)"))
                {
                    var lastCmd = commandSummaries.LastOrDefault();
                    var decodingTimestamp = line.Substring(11, 12);
                    if (lastCmd != null && lastCmd.CommandName == "AcquireSingleChannelTargetWanSearch" )
                    {
                        lastCmd.Status = "F";
                        lastCmd.TimeTaken = CalculateTotalTime(lastCmd.Timestamp.Substring(11, 12), decodingTimestamp);
                    }

                }
                if (line.Contains("Building Get8052Command"))
                {
                    currentTimestamp = line.Substring(0, 23);
                   // Match match = Regex.Match(line, @"-\\s*([\\d]+\\.[\\d]+\\.[\\d]+\\.[\\d]+)\\s*-");
                    Match match = Regex.Match(line, @"-\s*(\d+\.\d+\.\d+\.\d+)\s*-");
                    currentBuildVersion = match.Success ? match.Groups[1].Value : null;
                    
                    var lastCmd = commandSummaries.LastOrDefault();
                    if(lastCmd != null && lastCmd.CommandName == "ConnectLocalRadio" && lastCmd.Status == "F")
                    {
                        continue;
                    }
                    if (lastCmd != null && lastCmd.CommandName == "AcquireSingleChannelTargetWanSearch" && lastCmd.Status=="F")
                    {
                        lastCmd.Status = "S";
                    }
                    commandSummaries.Add(new DeviceCommandSummary
                    {
                        BuildVersion = currentBuildVersion,
                        CommandName = "Device8052",
                        Timestamp = currentTimestamp,
                        TimeTaken = "NA",  // Empty for now
                        Status = "F"     // Default status Failed
                    });

                }
                if(line.Contains("Tunerdatacommand Command Executed"))
                {
                    var lastCmd = commandSummaries.LastOrDefault();
                    if (lastCmd != null && lastCmd.CommandName == "ConnectLocalRadio" && lastCmd.Status == "F")
                    {
                        continue;
                    }
                    if (lastCmd!=null && lastCmd.CommandName== "LocalRadioStatus")
                    {
                        continue;
                    }
                    currentTimestamp = line.Substring(0, 23);
                    Match match = Regex.Match(line, @"-\\s*([\\d]+\\.[\\d]+\\.[\\d]+\\.[\\d]+)\\s*-");
                    currentBuildVersion = match.Success ? match.Groups[1].Value : null;
                    commandSummaries.Add(new DeviceCommandSummary
                    {
                        BuildVersion = currentBuildVersion,

                        CommandName = "TunerData",
                        Timestamp = currentTimestamp,
                        TimeTaken = "NA",  // Empty for now
                        Status = "F"     // Default status Failed
                    });

                }


                if (line.Contains("Decoding the response for:"))
                {
                    var decodingCommand = ExtractDecodingCommand(line);
                    var decodingTimestamp = line.Substring(11, 12);

                    // Find the matching command
                    var commandEntry = commandSummaries.FindLast(c => c.CommandName == decodingCommand && c.Status == "F");

                    if (commandEntry != null)
                    {
                        commandEntry.TimeTaken = CalculateTotalTime(commandEntry.Timestamp.Substring(11,12), decodingTimestamp);
                        commandEntry.Status = "S"; // Success
                    }
                }
                //MESH IP CASE
                if(line.Contains("Local radio connected successfully"))
                {

                    var decodingTimestamp = line.Substring(11, 12);

                    // Find the matching command
                    var commandEntry = commandSummaries.FindLast(c => c.CommandName == "ConnectLocalRadio" );

                    if (commandEntry != null && commandEntry.Status=="F")
                    {
                        commandEntry.TimeTaken = CalculateTotalTime(commandEntry.Timestamp.Substring(11, 12), decodingTimestamp);
                        commandEntry.Status = "S"; // Success
                    }
                   

                }
                if (line.Contains("DCW Version"))
                {
                    var currentDcwVersion = ExtractDCWVersion(line);
                    var lastCmd = commandSummaries.LastOrDefault();
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");

                    if (firstCmd != null && lastCmd != null && (lastCmd.CommandName == "TunerData" 
                        || lastCmd.CommandName== "ManufacturingData"))
                    {

                        // Find index range
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);

                        // Apply comment only in that range
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].DcwVersion))
                            { 
                                commandSummaries[j].DcwVersion = currentDcwVersion;
                            }
                        }
                    }
                }
                if(line.Contains("Module Firmware Version"))
                {
                    var currentModuleFirmwareVersion = ExtractModuleFirmwareVersion(line);
                    var lastCmd = commandSummaries.LastOrDefault();
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");
                    if (lastCmd != null && (lastCmd.CommandName == "TunerData"
                        || lastCmd.CommandName == "ManufacturingData") && firstCmd != null)
                    {
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].ModuleFirmwareVersion))
                            {
                                commandSummaries[j].ModuleFirmwareVersion = currentModuleFirmwareVersion;
                            }
                        }
                    }

                }
                if(line.Contains("Meter Firmware Version"))
                {
                     var currentMeterFirmwareVersion = ExtractMeterFirmwareVersion(line); 
                    var lastCmd = commandSummaries.LastOrDefault();
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");
                    if (lastCmd != null && (lastCmd.CommandName == "TunerData" || lastCmd.CommandName == "ManufacturingData") && firstCmd !=null)
                    {
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].MeterFirmwareVersion))
                            {
                                commandSummaries[j].MeterFirmwareVersion = currentMeterFirmwareVersion;
                            }
                        }
                    }

                }
                if (line.Contains("InitializeForDevice"))
                {
                    var lastCmd = commandSummaries.LastOrDefault();
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");
                    if (lastCmd != null && lastCmd.CommandName == "TunerData" && lastCmd.Status == "S" && string.IsNullOrEmpty(lastCmd.MajorMinorSoftID) &&
                        firstCmd!= null)
                    {
                        (string majorId, string minorId, string softId) = ExtractIds(line);
                        if (majorId != null && minorId != null && softId != null)
                        {
                            var currentMajorMinorSoftID = "0x" + majorId + "/0x" + minorId + "/0x" + softId;
                            int startIndex = commandSummaries.LastIndexOf(firstCmd);
                            int endIndex = commandSummaries.IndexOf(lastCmd);
                            for (int j = startIndex; j <= endIndex; j++)
                            {
                                if (string.IsNullOrEmpty(commandSummaries[j].MajorMinorSoftID))
                                {
                                    commandSummaries[j].MajorMinorSoftID = currentMajorMinorSoftID;
                                }
                            }
                        }
                    }

                }
                if (line.Contains("Failed to connect the endpoint"))
                {
                    string currComment = Error6;
                    E6Count++;
                    var lastCmd = commandSummaries.LastOrDefault();
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");

                    if (firstCmd != null && lastCmd != null)
                    {
                        lastCmd.Status = "F";

                        // Find index range
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);

                        // Apply comment only in that range
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].Comments))
                            {
                                commandSummaries[j].Comments = currComment;
                            }
                        }
                    }

                }
                if(line.Contains("failed to complete"))
                {
                    
                    var lastCmd = commandSummaries.LastOrDefault();
                    (string failureTime, string failureCommand) = ExtractFailureDetails(line);
                    if (failureTime!=null && failureCommand!=null && lastCmd != null && lastCmd.Status=="F")
                    {
                        lastCmd.TimeTaken = CalculateTotalTime(lastCmd.Timestamp.Substring(11, 12), failureTime);
                    }
                    if (line.Contains(Error1))
                    {
                        E1 = true;
                    }
                    if (line.Contains(Error2))
                    {
                        E2 = true;
                    }
                    
                }

                //errors
                if (line.Contains("Retry attempt for 8052 command is exhausted and device did not connect")
                      || line.Contains("Processing of 80 52 response failed"))
                {
                    currentComment = ExtractTextAfterHyphen(line);
                    if (E1==true)
                    {
                        currentComment = Error1;
                        
                    }
                    E1Count++;
                    FailCount++;

                    var lastCmd = commandSummaries.LastOrDefault(c => c.CommandName == "Device8052");
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");

                    if (firstCmd != null && lastCmd != null)
                    {
                        lastCmd.Status = "F";

                        // Find index range
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);

                        // Apply comment only in that range
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].Comments))
                            {
                                commandSummaries[j].Comments = currentComment;
                            }
                        }
                    }
                    E1 = false;
                }

                if(line.Contains("Retry attempt for TunerData command is exhausted and device did not connect, hence notify the view"))
                {
                    currentComment = ExtractTextAfterHyphen(line);
                    if(E2 == true)
                    {
                        currentComment = Error2;
                    }
                    E2Count++;
                    FailCount++;

                    var lastCmd = commandSummaries.LastOrDefault(c => c.CommandName == "TunerData");
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");

                    if (firstCmd != null && lastCmd != null)
                    {
                        lastCmd.Status = "F";

                        // Find index range
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);

                        // Apply comment only in that range
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].Comments))
                            {
                                commandSummaries[j].Comments = currentComment;
                            }
                        }
                    }
                    E2 = false;

                }

                if(line.Contains("Maximum retry attempts reached for AcquireSingleChannelWanSearchFailed"))
                {
                    currentComment = ExtractTextAfterHyphen(line);
                    E5Count++;
                    FailCount++;
                    var decodingTimestamp = line.Substring(11, 12);

                    // Find the matching command

                    var lastCmd = commandSummaries.LastOrDefault(c => c.CommandName == "AcquireSingleChannelTargetWanSearch");
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");

                    if (firstCmd != null && lastCmd != null)
                    {
                        lastCmd.Status = "F";
                        lastCmd.TimeTaken = CalculateTotalTime(lastCmd.Timestamp.Substring(11, 12), decodingTimestamp);

                        // Find index range
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);

                        // Apply comment only in that range
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].Comments))
                            {
                                commandSummaries[j].Comments = currentComment;
                            }
                        }
                    }

                }
                if (line.Contains(Error3))
                {
                    E3 = true;
                }
                if(line.Contains("Retry attempt for Electric Pana command is exhausted and device did not connect, hence notify the view"))
                {
                    currentComment = ExtractTextAfterHyphen(line);
                    if (E3 == true)
                    {
                        currentComment = Error3;
                    }
                    E3Count++;
                    FailCount++;

                    var lastCmd = commandSummaries.LastOrDefault(c => c.CommandName == "PanaCommand");
                    var firstCmd = commandSummaries.LastOrDefault(c => c.CommandName == "ConnectLocalRadio");

                    if (firstCmd != null && lastCmd != null)
                    {
                        lastCmd.Status = "F";

                        // Find index range
                        int startIndex = commandSummaries.LastIndexOf(firstCmd);
                        int endIndex = commandSummaries.IndexOf(lastCmd);

                        // Apply comment only in that range
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            if (string.IsNullOrEmpty(commandSummaries[j].Comments))
                            {
                                commandSummaries[j].Comments = currentComment;
                            }
                        }
                    }
                    E3 = false;

                }



            }
            List<List<DeviceCommandSummary>> lanGroups = new List<List<DeviceCommandSummary>>();
            List<DeviceCommandSummary> currentGroup = new List<DeviceCommandSummary>();
            string previousModuleID = null;

            //ModuleID 
            int mx = 0;

            foreach (var command in commandSummaries)
            {
                if (command.CommandName == "ConnectLocalRadio")
                {
                    if (currentGroup.Count > 0)
                    {
                        lanGroups.Add(currentGroup);
                        mx = Math.Max(mx, currentGroup.Select(c => c.CommandName).Distinct().Count());
                    }
                    currentGroup = new List<DeviceCommandSummary> { command };
                }
                else if (command.ModuleID == previousModuleID)
                {
                    currentGroup.Add(command);
                }
                else
                {
                    if (currentGroup.Count > 0)
                    {
                        lanGroups.Add(currentGroup);
                        mx = Math.Max(mx, currentGroup.Select(c => c.CommandName).Distinct().Count());
                    }
                    currentGroup = new List<DeviceCommandSummary> { command };
                }

                previousModuleID = command.ModuleID;
            }

            // Final group check
            if (currentGroup.Count > 0)
            {
                lanGroups.Add(currentGroup);
                mx = Math.Max(mx, currentGroup.Select(c => c.CommandName).Distinct().Count());
            }

            //EXCEL
            using (var workbook = new XLWorkbook())
            {
                var ws1 = workbook.Worksheets.Add("Device Info");
                ws1.SheetView.FreezeRows(1);
                // Headers
                ws1.Cell(1, 1).Value = "Build Version";
                ws1.Cell(1, 2).Value = "LocalRadio ID";
                ws1.Cell(1, 3).Value = "Module ID";

                int i = 1, col = 4;  //col=4
                while (i <= mx)
                {

                    ws1.Cell(1, col).Value = "Command" + i;
                    ws1.Cell(1, col + 1).Value = "Timestamp";
                    ws1.Cell(1, col + 2).Value = "Time Taken (s)";
                    ws1.Cell(1, col + 3).Value = "Status";
                    col = col + 4;
                    i++;
                }
                col = col - 1;
                ws1.Cell(1, col + 1).Value = "Comments";//col - 39 mx - 9
                ws1.Cell(1, col + 2).Value = "DCW Version";  //col+1
                ws1.Cell(1, col + 3).Value = "Module Firmware Version";//col+2
                ws1.Cell(1, col + 4).Value = "Meter Firmware Version";//col+3
                ws1.Cell(1, col + 5).Value = "Major/Minor/Soft ID";
                ws1.Cell(1, col + 6).Value = "TotalTimeTakenPerRun";
                int totalCol = (mx * 4) + 9;
                for (int c = 1; c <= totalCol; c++)
                {
                    var cell = ws1.Cell(1, c);
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                int row = 2;

                foreach (var group in lanGroups)
                {
                    var firstCommand = group.FirstOrDefault();

                    if (firstCommand != null)
                    {
                        int currentRow = row;

                        ws1.Cell(currentRow, 1).Value = firstCommand.BuildVersion;
                        TotalRun++;
                        ws1.Cell(currentRow, 2).Value = firstCommand.LocalRadioID;
                        ws1.Cell(currentRow, 3).Value = firstCommand.ModuleID;

                        string previousCommand = null;
                        int commandColumn = 4;

                        foreach (var command in group)
                        {
                            
                            if (command.CommandName == previousCommand)
                            {
                                currentRow++;        // go to next row for retry
                                commandColumn -= 4;   // reset column to original command column
                            }

                            ws1.Cell(currentRow, commandColumn).Value = command.CommandName;
                            ws1.Cell(currentRow, commandColumn + 1).Value = command.Timestamp;
                            ws1.Cell(currentRow, commandColumn + 2).Value = command.TimeTaken;

                            var statusCell = ws1.Cell(currentRow, commandColumn + 3);
                            statusCell.Value = command.Status;

                            if (command.Status == "S")
                            {
                                statusCell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                            }
                            else if (command.Status == "F")
                            {
                                statusCell.Style.Fill.BackgroundColor = XLColor.LightCoral;
                            }

                            statusCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            previousCommand = command.CommandName;
                            commandColumn += 4;
                        }

                        int baseColumn = 3 + (mx * 4);
                        ws1.Cell(currentRow, baseColumn + 1).Value = firstCommand.Comments;
                        ws1.Cell(currentRow, baseColumn + 2).Value = firstCommand.DcwVersion;
                        ws1.Cell(currentRow, baseColumn + 3).Value = firstCommand.ModuleFirmwareVersion;
                        ws1.Cell(currentRow, baseColumn + 4).Value = firstCommand.MeterFirmwareVersion;
                        ws1.Cell(currentRow, baseColumn + 5).Value = firstCommand.MajorMinorSoftID;
                        ws1.Cell(currentRow, baseColumn + 6).Value = firstCommand.TotalTimeTakenPerRun;

                        row = currentRow + 1; // move to the next starting row for next group
                    }
                }

                var ws2 = workbook.Worksheets.Add("Connection Summary");

                //error percentage 
                void ApplyCellStyling(IXLCell cell)
                {
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                void AddLabeledValueRow(IXLWorksheet ws, int row, string label, double? numericValue = null, string textValue = null)
                {
                    var labelCell = ws.Cell(row, 1);
                    var valueCell = ws.Cell(row, 2);

                    labelCell.Value = label;

                    if (numericValue.HasValue)
                        valueCell.Value = numericValue.Value;
                    else if (!string.IsNullOrEmpty(textValue))
                        valueCell.Value = textValue;
                    else
                        valueCell.Value = string.Empty;

                    ApplyCellStyling(labelCell);
                    ApplyCellStyling(valueCell);
                }

                AddLabeledValueRow(ws2, 1, "Total Runs", numericValue: TotalRun);
                AddLabeledValueRow(ws2, 2, "Total fail", numericValue: FailCount);
                AddLabeledValueRow(ws2, 3, "Failure %", numericValue: ((double)FailCount / TotalRun) * 100);
                AddLabeledValueRow(ws2, 4, "Success%", numericValue: ((double)(TotalRun - FailCount) / TotalRun) * 100);
                AddLabeledValueRow(ws2, 5, Error1, numericValue: E1Count);
                AddLabeledValueRow(ws2, 6, Error2, numericValue: E2Count);
                AddLabeledValueRow(ws2, 7, Error3, numericValue: E3Count);
                AddLabeledValueRow(ws2, 8, "Invalid LocalRadioID", numericValue: E4Count);
                AddLabeledValueRow(ws2, 9, "AcquireSingleChannelWanSearchFailed Max Retry",numericValue: E5Count);
                AddLabeledValueRow(ws2, 10, Error6, numericValue: E6Count);



                ws2.Columns().AdjustToContents();
                workbook.SaveAs(excelPath);
            }
           
        }

        private static string ExtractAddedCommand(string logLine)
        {
            var match = Regex.Match(logLine, @"Added command:\s+(\w+)");
            return match.Success ? match.Groups[1].Value : null;
        }
        private static string ExtractExecutedCommand(string logLine)
        {
            var match = Regex.Match(logLine, @"Command Executed:\s+(\w+)");
            return match.Success ? match.Groups[1].Value : null;
        }

        private static string CalculateTotalTime(string startTimestamp, string endTimestamp)
        {
            try
            {
                DateTime startTime = DateTime.ParseExact(startTimestamp, "HH:mm:ss,fff", null);
                DateTime endTime = DateTime.ParseExact(endTimestamp, "HH:mm:ss,fff", null);

                TimeSpan duration = endTime - startTime;
                return duration.TotalSeconds.ToString("F2"); // Format to 2 decimal places
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error calculating total time: " + ex.Message);
                return "Error";
            }
        }
        private static string ExtractDecodingCommand(string logLine)
        {
            var match = Regex.Match(logLine, @"Decoding the response for:\s+(\w+)");
            return match.Success ? match.Groups[1].Value : null;
        }
        private static string ExtractDCWVersion(string logLine)
        {

            var regex = new Regex(@"\b([0-9A-F]{2,4}(?:\.[0-9A-F]{2,4}){2,})\b", RegexOptions.IgnoreCase);
            var match = regex.Match(logLine);

            
            return match.Success ? match.Value : null;
        }
        private static string ExtractModuleFirmwareVersion(string logLine)
        {
            var regex = new Regex(@"\b([A-Z0-9]{5,}-\d{2}\.\d{2}\.[A-Z0-9]{3})\b");
            var match = regex.Match(logLine);
            return match.Success ? match.Value : null;
        }
        private static string ExtractMeterFirmwareVersion(string logLine)
        {

            var regex = new Regex(@"\b([A-Z0-9]{5,}-[A-Z0-9]{5,})\b", RegexOptions.IgnoreCase);
            var match = regex.Match(logLine);

            
            return match.Success ? match.Value : null;
        }

        private static (string majorId, string minorId, string softId) ExtractIds(string logLine)
        {
            var regex = new Regex(@"majorId (\w{2}) minorId (\w{2}) softId (\w{4})");
            var match = regex.Match(logLine);

            if (match.Success)
            {
                string majorId = match.Groups[1].Value;
                string minorId = match.Groups[2].Value;
                string softId = match.Groups[3].Value;
                return (majorId, minorId, softId);
            }

            return (null, null, null);
        }

        private static (string failureTime, string failureCommand) ExtractFailureDetails(string logLine)
        {
            // Extract the failure time using Substring
            string failureTime = logLine.Substring(11, 12);

            // Use regex to capture the failure command
            var regex = new Regex(@"-\s([\w\d\s]+)\sfailed to complete\.");
            var match = regex.Match(logLine);

            if (match.Success)
            {
                string failureCommand = match.Groups[1].Value.Trim();  // Remove any extra spaces
                return (failureTime, failureCommand);
            }

            return (failureTime, null);
        }

        private static string ExtractTextAfterHyphen(string logLine)
        {
            // Find the last hyphen and extract the remaining text
            int lastHyphenIndex = logLine.LastIndexOf(" - ");

            if (lastHyphenIndex != -1 && lastHyphenIndex + 3 < logLine.Length)
            {
                return logLine.Substring(lastHyphenIndex + 3).Trim();
            }

            return null;
        }

    }
}
