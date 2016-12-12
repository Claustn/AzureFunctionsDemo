$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$name = "HTTPBody $($requestBody.MyName)"

if ($req_query_MyName) 
{
    $name = "HTTP Req $req_query_MyName" 
}

if (-not (Get-Module -Name "PScribo"))
{
    Write-Output "PScribo not installed";
}
else
{
    Write-Output "PScribo installed";
}

Get-Module -Listavailable | Out-String
Write-output "$($env:MySecretUser)"
Out-File -Encoding Ascii -FilePath $res -inputObject "Hello $name"

