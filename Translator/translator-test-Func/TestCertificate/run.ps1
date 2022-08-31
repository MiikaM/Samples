using namespace System.Net
<#
.SYNOPSIS
    Tests if a certificate is valid for Sharepoint use
.DESCRIPTION
    Tests if the Certificate is valid and calls the Sharepoint base url via PnP Powershell with the certificate information.
.LINK
    https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
.EXAMPLE
    clientId - The client id is for the app registration for which uses the certificate
    CertificatePath - For azure function the test url is usually something like "C:\home\site\wwwroot\Certificates\<Certificate-name>.pfx" for local testing use local directory path.
    Tenant - Tenant e.g. <Tenant-name>.onmicrosoft.com
    Url - You can use any site in the Sharepoint page library but usually the base path is used e.g. "https://<Tenant-name>.sharepoint.com"
    Thumbprint - The certificate thumbprint
    CertPassword - The password for the certificate

    More help for creating and setting up a TLS/SSL certificate for your application can be found here =>  https://docs.microsoft.com/en-us/azure/app-service/configure-ssl-certificate-in-code

    Also for local testing make sure that you have PnP shell installed on your system. Here you can find information about installing the PnP powershell module => https://pnp.github.io/powershell/articles/installation.html 
#>


{0}

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)


# Write to the Azure Functions log stream.
Write-Host $Request
# POST method: $req
$requestBody = $Request.Body
$clientId = "<Client-id>"
$CertificatePath = "C:\home\site\wwwroot\Certificates\<Certificate-name>.pfx"
$Tenant = "<Tenant>.onmicrosoft.com"
$Url = "https://<Tenant>.sharepoint.com"
$Thumbprint = "<Thumbprint>"
$CertPassword = (ConvertTo-SecureString -AsPlainText '<Password>' -Force)
# Connect-PnPOnline -ClientId $clientId -Tenant $Tenant -Url $Url -Thumbprint $Thumbprint
Connect-PnPOnline -CertificatePath $CertificatePath -CertificatePassword $CertPassword -Tenant $Tenant -ClientId $clientId -Url $Url 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = "Yay"
})
