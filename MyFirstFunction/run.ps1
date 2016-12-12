$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$name = "HTTPBody $($requestBody.MyName)"

if ($req_query_MyName) 
{
    $name = "HTTP Req $req_query_MyName" 
}
Write-output "$($env:MySecretUser)"
Out-File -Encoding Ascii -FilePath $res -inputObject "Hello $name"