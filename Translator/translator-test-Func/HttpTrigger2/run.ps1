using namespace System.Net
using namespace System.Net

<#
.SYNOPSIS
    Translates a Sharepoint page to the intended language using Azure Translator service
.DESCRIPTION
    Translates a Sharepoint page to the intended language using Azure Translator service. Script takes 
    - pageTitle (e.g. <Page-name>.aspx)
    - siteUrl (e.g. "https://<Tenant>.sharepoint.com/sites/Translator-test")
    - language (e.g. "en").

    Language is gotten from the translated page e.g. for English it's "en" and is based on the folder names in the Sharepoint pages page library.
    The script first get's the Text content from the Sharepoint page with the use of PnP powershell. After getting the text areas it starts the translation via Azure Translator Service for each text area separately. When the
    translation is ready it set's the new text to the same text area with the InstanceId property.
    
    Some helpful links:
      - Installing PnP powershell => https://pnp.github.io/powershell/articles/installation.html
      - Cmdlets for PnP powershell => https://pnp.github.io/powershell/cmdlets/Add-PnPAlert.html
      - Azure Translator documentation => https://docs.microsoft.com/en-us/azure/cognitive-services/translator/
      - Configure TLS/SSL certificate for Azure function => https://docs.microsoft.com/en-us/azure/app-service/configure-ssl-certificate-in-code
.NOTES
    Remember to use ([System.Text.Encoding]::UTF8.GetBytes("[$textJson]")) when getting text from the sharepoint page since powershell doesn't use UTF8-encoding by default 
    More information => https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_character_encoding?view=powershell-7.2
.EXAMPLE
    requestBody - Should include following properties 
      - siteUrl
      - pageUrl
      - language
    clientId - The client id is for the app registration for which uses the certificate
    clientSecret - A valid client secret for the app registration.
    CertPath - For azure function the test url is usually something like "C:\home\site\wwwroot\Certificates\<Certificate-name>.pfx" for local testing use local directory path.
    translatorKey - Key for the translator instance
    Tenant - Tenant e.g. <Tenant-name>.onmicrosoft.com
    Url - You can use any site in the Sharepoint page library but usually the base path is used e.g. "https://<Tenant-name>.sharepoint.com"
    Thumbprint - The certificate thumbprint
    CertPass - The password for the certificate

    For local testing you can also connect to the Sharepoint page via Admin account credentials

    Script can be tested by 
      1. Create a resource group where you will also create an Azure Translator service instance (Help => https://docs.microsoft.com/en-us/azure/cognitive-services/translator/how-to-create-translator-resource)
      2. After creating the resource you have to update the HttpTrigger2 function Translator key environment variable matching your Translator instance
      3. Then run the azure function locally with ""func start"" cmdlet
      4. After the function is running go create a Sharepoint page with text blocks in some language. Then create a Sharepoint page translation for the page (HELP => https://support.microsoft.com/en-us/office/create-multilingual-sharepoint-sites-pages-and-news-2bb7d610-5453-41c6-a0e8-6f40b3ed750c) NOTE! Do not modify the translated page after creating the translation, the script will do that for you.
      5. After creating the translated page send the function the siteUrl, pageTitle and language property in the body of a REST request (recommended to use some type of REST api program e.g. Nightingale, PostMan, insomnia etc.)
      6. After the script has run the translated page should have been translated.
      
#>

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Write-Host $Request
# POST method: $req
$requestBody = $Request.Body
$clientId = $env:CLIENT_ID
$clientSecret = $env:CLIENT_SECRET
$translatorKey= $env:TRANSLATOR_KEY
$CertPath = $env:LocalCertificate
$Tenant = $env:TENANT
$CertPass = (ConvertTo-SecureString -AsPlainText $env:CERTIFICATION_PASSWORD -Force)

# Interact with body of the request
$siteURL = $requestBody.siteUrl
$targetLanguage = $requestBody.language
$pageTitle = $requestBody.pageTitle

Write-Host "$siteURL, $targetLanguage, $pageTitle, $translatorKey "


# Translate function
function Start-Translation{
    param(
    [Parameter(Mandatory=$true)]
    [string]$text,
    [Parameter(Mandatory=$true)]
    [string]$language
    )
 
    $baseUri = "https://api.cognitive.microsofttranslator.com"
   $headers= @{}

    $headers.Add("Ocp-Apim-Subscription-Key", $translatorKey)
    $headers.Add("Ocp-Apim-Subscription-Region", "northeurope")
    $headers.Add("Content-Type", "application/json")
    $headers.Add("charset", "UTF-8")

    # Create JSON array with 1 object for request body
    $textJson = @{
        "Text" = "$text"
        } | ConvertTo-Json
 
    $body = ([System.Text.Encoding]::UTF8.GetBytes("[$textJson]"))
    
  # Uri for the request includes language code and text type, which is always html for SharePoint text web parts
    $uri = "$baseUri/translate?api-version=3.0&to=$language&textType=html"
    
  #   Write-Output $uri
    
    # Write-Output $body

    # Send request for translation and extract translated text
    $results = Invoke-RestMethod -Method POST -Uri $uri -Body $body -Headers $headers
    $translatedText = $results[0].translations[0].text

    return $translatedText
}

Connect-PnPOnline -Url $siteURL -CertificatePath $CertPath -CertificatePassword $CertPass -ClientId $clientId -Tenant $Tenant

# $newPage = Get-PnPClientSidePage "$targetLanguage/$pageTitle.aspx"
$newPage = Get-PnPPage "$targetLanguage/$pageTitle.aspx"

$textControls = $newPage.Controls | Where-Object {$_.Type.Name -eq "PageText"}

Write-Output $textControls

 foreach ($textControl in $textControls) {
        $translatedControlText = Start-Translation -text $textControl.Text -language $targetLanguage
        # Set-PnPClientSideText -Page $newPage -InstanceId $textControl.InstanceId -Text $translatedControlText
        Set-PnPPageTextPart -Page $newPage -InstanceId $textControl.InstanceId -Text $translatedControlText
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = "{'message':'Page $pageTitle has been translated to $targetLanguage'}"
    })