#########################
# FilePropertiesToExcel #
# By Jacob Bean         #
#                       #
#############################
# Syncro Custom Assets Used #
#                           #
#                           #
#########################################################
# Description                                           #
#                                                       #
# retrives properties of any file type                  #
# specifically the date and the name                    #
# pump that shit into excel                             #
# then export it to local desktop or specified path     #
#                                                       #
#########################################################
# NOTES
# 10-6-22
# so I first need to give a file
# once the file has been given
# find the name 
# find the date
# Export to CSV
#########################################################
# CHANGE_LOG
# 
# planned changes / implemented changes
# done - feed a .txt of all file names
# get-content on a folder
# create a gui
# DRAG AND DROP UWU!!!!!!!
#
#
# If am set cell color if
##########################################################

##########################
# BEGIN GLOBAL VARIABLES #
##########################
$VerbosePreference = "continue"
$pcName = hostname
$files = @()
$outArray = @()
##########################
# END   GLOBAL VARIABLES #
##########################

################
#    -MAIN-    #
#              #
################

$file = Read-Host -Prompt "Enter the directory path`n`nExample: C:\Users\jbean\Desktop\paths`nPlease paste the exact file name"
$paths = Get-ChildItem -Path $file



$paths | ForEach-Object{
    $temp = $_
    $fileName = $temp.Name
    $fileDate = $temp.CreationTime
    $month = (Get-Culture).DateTimeFormat.GetMonthName($fileDate.Month) 
    $day = $fileDate.Day
    $year = $fileDate.Year
    $pmFlag = $false
    $amPM = ""

    $hour = $fileDate.Hour 
    $minute = $fileDate.Minute
    if($hour -gt 12){$pmFlag = $true}
    if($hour -lt 9){ $hour = "0"+$hour;}
    if($minute -le 9){$minute = "0"+$minute}

    $date = "$month" + "/" + "$day" + "/" + "$year"
    $time = "$hour" + ":" + "$minute"

    if($pmFlag -eq $true){$amPM = "PM"}
    else{$amPM = "AM"}

    #Write-Verbose "File Name: $fileName`n"
    #Write-Verbose "CreationDate: $date`n"
    #Write-Verbose "CreationTime: $time`n`n"

    $files = [pscustomobject]@{
        FileNumber = "$fileName"
        Month  = "$month"
        Day = "$day"
        Time = "$time"
        am_pm = "$amPM"
    }
    foreach($object in $files){
        Write-Verbose "" -verbose

        $tempFileNumber = $object.FileNumber
        $tempMonth = $object.Month
        $tempDay = $object.Day
        $tempTime = $object.Time
        $tempAM_PM = $object.am_pm

        Write-Verbose "$tempFileNumber" -verbose
        Write-Verbose "$tempMonth" -verbose
        Write-Verbose "$tempDay" -verbose
        Write-Verbose "$tempTime" -verbose
        Write-Verbose "$tempAM_PM" -verbose

    $outArray += New-Object PsObject -Property @{


        'FileNumber' = $object.FileNumber
        'Month' = $object.Month
        'Day' = $object.Day
        'Time' = $object.Time
        'AM/PM' = $object.am_pm
    
    }


}

}



#Write-Output $files
$userDesktop = "C:\Exports"
#input file full of the files

$outArray | Export-Csv -NoTypeInformation -literalPath "$userDesktop\temp.csv"
