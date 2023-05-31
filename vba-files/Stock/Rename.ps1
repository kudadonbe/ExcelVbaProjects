# Author: Kudadonbe
# Date: 2023-05-30 
# Version: 1.0.0

# set suffix="GRF-2023-"
# rename all files in the folder one by one
# get the file 
# get the request number from the file name
# change request number format to "000"
# make new file name with the suffix and the request number
# rename the file with the new file name

$suffix="GRF-2023-"
$files = Get-ChildItem -Path .\ -Filter *.xlsx

foreach ($file in $files) {
    $requestNumber = $file.Name.Substring(0, $file.Name.IndexOf("_"))
    $newFileName = $suffix + $requestNumber.PadLeft(3, "0") + ".xlsx"
    Rename-Item -Path $file.Name -NewName $newFileName
}

    
