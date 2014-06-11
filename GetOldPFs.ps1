#Get all the PFs
$publicfolders=Get-PublicFolder -Recurse -ResultSize Unlimited
#Set up some variables
$newfolders=@()
$oldfolders=@()
$newfolderstats=@()
$oldfolderstats=@()
#loop through each of the PFs looking for "new" content
foreach ($publicfolder in $publicfolders) {
    #Ignore the root
    If(($publicfolder.identity -eq '\') -or ($publicfolder.identity -eq "IPM_SUBTREE")) {
    }else{
        #grab today's date
        $date=Get-Date
        #set $old to null
        $old=""
        #if there's an item within the time frame then push it into $old
        $old=Get-PublicFolderItemStatistics $publicfolder | ?{$_.LastModificationTime -gt $date.AddYears(-3)}
        if ($old -ne $null) {
            #if $old is not empty we have a PF that is "in use"
            $newfolders+=$publicfolder
        } else {
            #If it's empty then the PF is not "in use"
            $oldfolders+=$publicfolder
        }
    }
}
#get the PF stats of each "in use" PF
foreach ($newfolder in $newfolders){
    $newfolderstats+=Get-PublicFolderStatistics $newfolder.identity
}
#get the PF stats for each "old" PF
foreach ($oldfolder in $oldfolders){
    $oldfolderstats+=Get-PublicFolderStatistics $oldfolder.identity
}
#send your output to a CSV file
$newfolderstats | Export-Csv newfolders.csv
$oldfolderstats | Export-Csv oldfolders.csv