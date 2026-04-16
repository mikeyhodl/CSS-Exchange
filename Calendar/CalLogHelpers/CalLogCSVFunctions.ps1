# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# ===================================================================================================
# Constants to support the script
# ===================================================================================================
# Non-US date formats for TryParseExact fallback (constant, created once)
[string[]]$script:DateFormats = @(
    "d/M/yyyy h:mm:ss tt",
    "d/M/yyyy H:mm:ss",
    "d/M/yyyy H:mm",
    "d/M/yyyy h:mm tt",
    "dd/MM/yyyy h:mm:ss tt",
    "dd/MM/yyyy H:mm:ss",
    "dd/MM/yyyy HH:mm:ss",
    "dd/MM/yyyy H:mm",
    "dd/MM/yyyy h:mm tt",
    "d/M/yyyy",
    "dd/MM/yyyy"
)

$script:CalendarItemTypes = @{
    'IPM.Schedule.Meeting.Request.AttendeeListReplication' = "AttendeeList"
    'IPM.Schedule.Meeting.Canceled'                        = "Cancellation"
    'IPM.OLE.CLASS.{00061055-0000-0000-C000-000000000046}' = "Exception"
    'IPM.Schedule.Meeting.Notification.Forward'            = "Forward.Notification"
    'IPM.Appointment'                                      = "Ipm.Appointment"
    'IPM.Schedule.Meeting.Request'                         = "Meeting.Request"
    'IPM.CalendarSharing.EventUpdate'                      = "SharingCFM"
    'IPM.CalendarSharing.EventDelete'                      = "SharingDelete"
    'IPM.Schedule.Meeting.Resp'                            = "Resp.Any"
    'IPM.Schedule.Meeting.Resp.Neg'                        = "Resp.Neg"
    'IPM.Schedule.Meeting.Resp.Tent'                       = "Resp.Tent"
    'IPM.Schedule.Meeting.Resp.Pos'                        = "Resp.Pos"
    '(Occurrence Deleted)'                                 = "Exception.Deleted"
}

<#
.SYNOPSIS
Resolves the ItemClass to a friendly type name using the CalendarItemTypes table.
If no exact match is found, progressively strips the last dot-segment and retries.
e.g. IPM.Schedule.Meeting.Resp.Tent.Follow → IPM.Schedule.Meeting.Resp.Tent → "Resp.Tent"
#>
function GetItemType {
    param([string]$ItemClass)
    $lookup = $ItemClass
    while (-not [string]::IsNullOrEmpty($lookup)) {
        $type = $script:CalendarItemTypes[$lookup]
        if (-not [string]::IsNullOrEmpty($type)) {
            return $type
        }
        # Strip the last segment and try the parent
        $lastDot = $lookup.LastIndexOf('.')
        if ($lastDot -le 0) { break }
        $lookup = $lookup.Substring(0, $lastDot)
    }
    return $null
}

# ===================================================================================================
# Functions to support the script
# ===================================================================================================

<#
.SYNOPSIS
Looks to see if there is a Mapping of ExternalMasterID to FolderName
#>
function MapSharedFolder {
    param(
        $ExternalMasterID
    )
    if ($ExternalMasterID -eq "NotFound") {
        return "Not Shared"
    } else {
        $script:SharedFolders[$ExternalMasterID]
    }
}

<#
.SYNOPSIS
Replaces a value of NotFound with a blank string.
#>
function ReplaceNotFound {
    param (
        $Value
    )
    if ($Value -eq "NotFound") {
        return ""
    } else {
        return $Value
    }
}

<#
.SYNOPSIS
Creates a Mapping of ExternalMasterID to FolderName
#>
function CreateExternalMasterIDMap {
    # This function will create a Map of the log folder to ExternalMasterID
    $script:SharedFolders = [System.Collections.SortedList]::new()
    $unknownCount = 0
    Write-Verbose "Starting CreateExternalMasterIDMap"

    foreach ($ExternalID in $script:GCDO.ExternalSharingMasterId | Select-Object -Unique) {
        if ($ExternalID -eq "NotFound") {
            continue
        }

        $AllFolderNames = @($script:GCDO | Where-Object { $_.ExternalSharingMasterId -eq $ExternalID } | Select-Object -ExpandProperty OriginalParentDisplayName | Select-Object -Unique)

        # Remove empty/null and 'Calendar' (the default folder name) entries
        $AllFolderNames = @($AllFolderNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        if ($AllFolderNames.count -gt 1) {
            # We have 2+ FolderNames, remove only exact match to 'Calendar' (the default folder name, not folder names containing 'Calendar')
            $Filtered = $AllFolderNames | Where-Object { $_ -ne 'Calendar' }
            if ($Filtered.Count -gt 0) {
                $AllFolderNames = @($Filtered)
            }
            # else keep the original list — all entries were 'Calendar'
        }

        if ($AllFolderNames.Count -eq 0) {
            $unknownCount++
            $script:SharedFolders[$ExternalID] = "UnknownSharedCalendarCopy$unknownCount"
            Write-Host -ForegroundColor red "Found Zero folder names to map to for ExternalID [$ExternalID]."
        } elseif ($AllFolderNames.Count -eq 1) {
            $script:SharedFolders[$ExternalID] = $AllFolderNames[0]
            Write-Verbose "Found map: [$($AllFolderNames[0])] is for $ExternalID"
        } else {
            # we still have multiple possible Folder Names, need to chose one or combine
            Write-Host -ForegroundColor Red "Unable to Get Exact Folder for $ExternalID"
            Write-Host -ForegroundColor Red "Found $($AllFolderNames.count) possible folders: $($AllFolderNames -join ', ')"

            if ($AllFolderNames.Count -eq 2) {
                $script:SharedFolders[$ExternalID] = $AllFolderNames[0] + " + " + $AllFolderNames[1]
            } else {
                $unknownCount++
                $script:SharedFolders[$ExternalID] = "UnknownSharedCalendarCopy$unknownCount"
            }
        }
    }

    Write-Host -ForegroundColor Green "Created the following Shared Calendar Mapping:"
    foreach ($Key in $script:SharedFolders.Keys) {
        Write-Host -ForegroundColor Green "$Key : $($script:SharedFolders[$Key])"
    }
    Write-Verbose "Created the following Mapping :"
    Write-Verbose $script:SharedFolders
}

<#
.SYNOPSIS
Convert a csv value to multiLine.
#>
function MultiLineFormat {
    param(
        $PassedString
    )
    $PassedString = $PassedString -replace "},", "},`n"
    return $PassedString.Trim()
}

# ===================================================================================================
# Build CSV to output
# ===================================================================================================

<#
.SYNOPSIS
Builds the CSV output from the Calendar Diagnostic Objects
#>
function BuildCSV {

    Write-Host "Starting to Process Calendar Logs..."
    $script:MailboxList = @{}
    # Initialize lookup caches to avoid redundant CN resolution across hundreds of log entries
    $script:SMTPAddressCache = @{}
    $script:DisplayNameCache = @{}
    Write-Host "Creating Map of Mailboxes to CNs..."
    CreateExternalMasterIDMap
    ConvertCNtoSMTP
    FixCalendarItemType($script:GCDO)

    Write-Host "Making Calendar Logs more readable..."
    $Index = 0
    $GCDOResults = foreach ($CalLog in $script:GCDO) {
        $Index++
        $ItemType = GetItemType $CalLog.ItemClass

        # CleanNotFounds
        $PropsToClean = "FreeBusyStatus", "ClientIntent", "AppointmentSequenceNumber", "AppointmentLastSequenceNumber", "RecurrencePattern", "AppointmentAuxiliaryFlags", "EventEmailReminderTimer", "IsSeriesCancelled", "AppointmentCounterProposal", "MeetingRequestType", "SendMeetingMessagesDiagnostics", "AttendeeCollection"
        foreach ($Prop in $PropsToClean) {
            # Exception objects, etc. don't have these properties.
            if ($null -ne $CalLog.$Prop) {
                $CalLog.$Prop = ReplaceNotFound($CalLog.$Prop)
            }
        }

        # Output one row (collected by the foreach assignment)
        [PSCustomObject]@{
            #'LogRow'                         = $Index
            'LogTimestamp'                   = ConvertDateTime($CalLog.LogTimestamp)
            'LogRowType'                     = $CalLog.LogRowType.ToString()
            'SubjectProperty'                = $CalLog.SubjectProperty
            'Client'                         = $CalLog.ShortClientInfoString
            'LogClientInfoString'            = $CalLog.LogClientInfoString
            'TriggerAction'                  = $CalLog.CalendarLogTriggerAction
            'ItemClass'                      = $ItemType
            'Seq:Exp:ItemVersion'            = CompressVersionInfo($CalLog)
            'Organizer'                      = GetDisplayName($CalLog.From)
            'From'                           = GetSMTPAddress($CalLog.From)
            'FreeBusy'                       = $CalLog.FreeBusyStatus.ToString()
            'ResponsibleUser'                = GetSMTPAddress($CalLog.ResponsibleUserName)
            'Sender'                         = GetSMTPAddress($CalLog.Sender)
            'LogFolder'                      = $CalLog.ParentDisplayName
            'OriginalLogFolder'              = $CalLog.OriginalParentDisplayName
            'SharedFolderName'               = MapSharedFolder($CalLog.ExternalSharingMasterId)
            'ReceivedRepresenting'           = GetSMTPAddress($CalLog.ReceivedRepresenting)
            'MeetingRequestType'             = $CalLog.MeetingRequestType.ToString()
            'StartTime'                      = ConvertDateTime($CalLog.StartTime)
            'EndTime'                        = ConvertDateTime($CalLog.EndTime)
            'OriginalStartDate'              = ConvertDateTime($CalLog.OriginalStartDate)
            'Location'                       = $CalLog.Location
            'CalendarItemType'               = $CalLog.CalendarItemType.ToString()
            'RecurrencePattern'              = $CalLog.RecurrencePattern
            'AppointmentAuxiliaryFlags'      = $CalLog.AppointmentAuxiliaryFlags.ToString()
            'DisplayAttendeesAll'            = $(if ($CalLog.DisplayAttendeesAll -eq "NotFound") { "-" } else { $CalLog.DisplayAttendeesAll })
            'AttendeeCount'                  = GetAttendeeCount($CalLog.DisplayAttendeesAll)
            'AppointmentState'               = $CalLog.AppointmentState.ToString()
            'ResponseType'                   = $CalLog.ResponseType.ToString()
            'ClientIntent'                   = $CalLog.ClientIntent.ToString()
            'AppointmentRecurring'           = $CalLog.AppointmentRecurring
            'HasAttachment'                  = $CalLog.HasAttachment
            'IsCancelled'                    = $CalLog.IsCancelled
            'IsAllDayEvent'                  = $CalLog.IsAllDayEvent
            'Sensitivity'                    = $CalLog.Sensitivity
            'IsSeriesCancelled'              = $CalLog.IsSeriesCancelled
            'SendMeetingMessagesDiagnostics' = $CalLog.SendMeetingMessagesDiagnostics
            'AttendeeCollection'             = MultiLineFormat($CalLog.AttendeeCollection)
            'CalendarLogRequestId'           = $CalLog.CalendarLogRequestId.ToString()    # Move to front.../ Format in groups???
            'CleanGlobalObjectId'            = $CalLog.CleanGlobalObjectId
        }
    }
    $script:EnhancedCalLogs = $GCDOResults

    Write-Host -ForegroundColor Green "Calendar Logs have been processed, Exporting logs to file..."
    Export-CalLog
}

function ConvertDateTime {
    param(
        [string] $DateTime
    )
    if ([string]::IsNullOrEmpty($DateTime) -or
        $DateTime -eq "N/A" -or
        $DateTime -eq "NotFound") {
        return ""
    }

    $InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    $DateStyles = [System.Globalization.DateTimeStyles]::None
    $Parsed = [DateTime]::MinValue

    # Priority 1 & 2: InvariantCulture TryParse
    # Handles ISO 8601 and M/d/yyyy (the default EXO output format)
    if ([DateTime]::TryParse($DateTime, $InvariantCulture, $DateStyles, [ref]$Parsed)) {
        return $Parsed
    }

    # Priority 3: Explicit non-US formats via TryParseExact
    # Reached when day > 12 makes InvariantCulture fail (e.g., "22/07/2024")
    if ([DateTime]::TryParseExact($DateTime, $script:DateFormats, $InvariantCulture, $DateStyles, [ref]$Parsed)) {
        return $Parsed
    }

    # Priority 4: Current culture (last resort for unexpected formats)
    if ([DateTime]::TryParse($DateTime, [ref]$Parsed)) {
        return $Parsed
    }

    # Priority 5: Return MinValue so sorting is not broken by mixed types
    Write-Warning "Unable to parse date: [$DateTime]"
    return [DateTime]::MinValue
}

function GetAttendeeCount {
    param(
        [string] $AttendeesAll
    )
    if ($AttendeesAll -ne "NotFound") {
        return ($AttendeesAll -split ';').Count
    } else {
        return "-"
    }
}

<#
.SYNOPSIS
Corrects the CalenderItemType column
#>
function FixCalendarItemType {
    param(
        $CalLogs
    )
    foreach ($CalLog in $CalLogs) {
        if ($CalLog.OriginalStartDate -ne "NotFound" -and ![string]::IsNullOrEmpty($CalLog.OriginalStartDate)) {
            $CalLog.CalendarItemType = "Exception"
            $CalLog.isException = $true
        }
    }
}

function CompressVersionInfo {
    param(
        $CalLog
    )
    [string] $CompressedString = ""
    if ($CalLog.AppointmentSequenceNumber -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.AppointmentSequenceNumber)) {
        $CompressedString = "-:"
    } else {
        $CompressedString = $CalLog.AppointmentSequenceNumber.ToString() + ":"
    }
    if ($CalLog.AppointmentLastSequenceNumber -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.AppointmentLastSequenceNumber)) {
        $CompressedString += "-:"
    } else {
        $CompressedString += $CalLog.AppointmentLastSequenceNumber.ToString() + ":"
    }
    if ($CalLog.ItemVersion -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.ItemVersion)) {
        $CompressedString += "-"
    } else {
        $CompressedString += $CalLog.ItemVersion.ToString()
    }

    return $CompressedString
}
