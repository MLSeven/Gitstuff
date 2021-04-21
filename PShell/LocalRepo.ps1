function list-repo
{
    $path = "/var/opt/morpheus/morpheus-ui/repo/git"

    $rootfolders = get-childitem $path | Select-object -prop name,fullname,lastwritetime,lastaccesstime
    $childfolders = $rootfolders.fullname | 
        foreach-object {$n=$_; Get-ChildItem $_ -recurse} | 
            select-object -prop @{name="name";expression={$n}},fullname,lastwritetime,lastaccesstime
    
    $all = $rootfolders+$childfolders
    $all
}

