param([string]$Prefix = "http://localhost:8765/")

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$dataRoot = if ($env:NEXTSTEP_DATA_DIR) { $env:NEXTSTEP_DATA_DIR } else { Join-Path $root ".nextstep-data" }
$uploadRoot = Join-Path $dataRoot "uploads"
$dbPath = Join-Path $dataRoot "state.json"
$script:mpesaCredentialPath = Join-Path $dataRoot "mpesa-credential.xml"
$script:premiumPriceKes = 300
$script:mailCredentialPath = Join-Path $dataRoot "mail-credential.xml"
$script:notificationEmail = "dn7229457@gmail.com"
$script:feedbackAttempts = @{}
$materialsRoot = if ($env:NEXTSTEP_MATERIALS_DIR) { $env:NEXTSTEP_MATERIALS_DIR } else { Join-Path $env:USERPROFILE "OneDrive\Desktop\Davesave" }
$script:sessions = @{}
$script:leases = @{}
$script:passwordResets = @{}
$script:localMaterials = @{
    local1 = "Week5_Chemiluminescence_and_Dissociation.pdf"
    local2 = "3-Hydrogen and hydrides.pptx"
    local3 = "4 -Group I.pptx"
    local4 = "6-Group II.ppt"
    local5 = "7-The p-Block Elements.pptx"
    local6 = "BBA 300 ORGANIZATIONAL THEORY  MODULE.doc"
    local7 = "Calculus 2 Part 2.pdf"
    local8 = "Chemistry_Group_2_Allied_Industries.pptx"
    local9 = "Lecture 1_2026.pdf"
    local10 = "Lecture2_2026.pdf"
    local11 = "Lecture3_2026.pdf"
    local12 = "Lecture 4_2026.pdf"
    local13 = "Numerical Integration for Calculus II.pdf"
}
$script:allowedExtensions = @(".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".csv", ".txt", ".md", ".zip")
$script:contentTypes = @{
    ".pdf" = "application/pdf"
    ".txt" = "text/plain; charset=utf-8"
    ".csv" = "text/csv; charset=utf-8"
    ".md" = "text/markdown; charset=utf-8"
    ".doc" = "application/msword"
    ".docx" = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ".ppt" = "application/vnd.ms-powerpoint"
    ".pptx" = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ".xls" = "application/vnd.ms-excel"
    ".xlsx" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ".zip" = "application/zip"
}

New-Item -ItemType Directory -Path $uploadRoot -Force | Out-Null
if (Test-Path $dbPath) {
    $script:db = Get-Content -Raw -Path $dbPath | ConvertFrom-Json
    if (-not $script:db.users) { $script:db.users = @() }
    if (-not $script:db.documents) { $script:db.documents = @() }
    if ($script:db.PSObject.Properties.Name -notcontains "payments") {
        $script:db | Add-Member -NotePropertyName payments -NotePropertyValue @()
    }
} else {
    $script:db = [pscustomobject]@{ users = @(); documents = @(); payments = @() }
}
foreach ($storedUser in @($script:db.users)) {
    if ($storedUser.PSObject.Properties.Name -notcontains "premiumUntilUtc") {
        $storedUser | Add-Member -NotePropertyName premiumUntilUtc -NotePropertyValue ""
    }
}

function Save-Database {
    $json = ConvertTo-Json -InputObject $script:db -Depth 20
    $tempPath = "$dbPath.tmp"
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tempPath, $json, $encoding)
    Move-Item -LiteralPath $tempPath -Destination $dbPath -Force
}

function Send-Json($Context, [int]$Status, $Value) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Value -Depth 20 -Compress))
    $Context.Response.StatusCode = $Status
    $Context.Response.ContentType = "application/json; charset=utf-8"
    $Context.Response.Headers["Cache-Control"] = "no-store"
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Send-Error($Context, [int]$Status, [string]$Message) {
    Send-Json $Context $Status @{ error = $Message }
}

function ConvertFrom-SecureText([System.Security.SecureString]$Value) {
    $pointer = [IntPtr]::Zero
    try {
        $pointer = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Value)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    }
    finally {
        if ($pointer -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
        }
    }
}

function Get-MpesaConfiguration {
    if (-not (Test-Path -LiteralPath $script:mpesaCredentialPath -PathType Leaf)) {
        throw "M-Pesa checkout is not configured. Run Configure-Nextstep-Mpesa.ps1 on this computer."
    }
    $configuration = Import-Clixml -LiteralPath $script:mpesaCredentialPath
    if (-not $configuration.ConsumerKey -or
        $configuration.ConsumerSecret -isnot [System.Security.SecureString] -or
        -not $configuration.BusinessShortCode -or
        $configuration.PassKey -isnot [System.Security.SecureString] -or
        -not $configuration.CallbackUrl -or
        $configuration.Environment -notin @("sandbox", "production") -or
        $configuration.TransactionType -notin @("CustomerPayBillOnline", "CustomerBuyGoodsOnline")) {
        throw "The saved M-Pesa configuration is invalid. Run Configure-Nextstep-Mpesa.ps1 again."
    }
    return $configuration
}

function Get-MpesaApiBase($Configuration) {
    if ($Configuration.Environment -eq "production") { return "https://api.safaricom.co.ke" }
    return "https://sandbox.safaricom.co.ke"
}

function Get-MpesaAccessToken($Configuration) {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $secret = ConvertFrom-SecureText $Configuration.ConsumerSecret
    $basic = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Configuration.ConsumerKey):$secret"))
    $headers = @{ Authorization = "Basic $basic" }
    $url = "$(Get-MpesaApiBase $Configuration)/oauth/v1/generate?grant_type=client_credentials"
    $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -TimeoutSec 20
    if (-not $response.access_token) { throw "Safaricom did not return an M-Pesa access token." }
    return [string]$response.access_token
}

function Get-MpesaPassword($Configuration, [string]$Timestamp) {
    $passKey = ConvertFrom-SecureText $Configuration.PassKey
    $source = "$($Configuration.BusinessShortCode)$passKey$Timestamp"
    return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($source))
}

function Normalize-KenyanPhone([string]$Phone) {
    $digits = $Phone.Trim() -replace "[\s()-]", ""
    if ($digits -match "^\+?254([17]\d{8})$") { return "254$($Matches[1])" }
    if ($digits -match "^0([17]\d{8})$") { return "254$($Matches[1])" }
    return $null
}

function Find-PaymentByCheckoutId([string]$CheckoutId) {
    return @($script:db.payments | Where-Object { $_.checkoutRequestId -eq $CheckoutId } | Select-Object -First 1)
}

function Invoke-MpesaStkQuery($Configuration, [string]$CheckoutRequestId) {
    $timestamp = (Get-Date).ToString("yyyyMMddHHmmss")
    $payload = @{
        BusinessShortCode = [string]$Configuration.BusinessShortCode
        Password = Get-MpesaPassword $Configuration $timestamp
        Timestamp = $timestamp
        CheckoutRequestID = $CheckoutRequestId
    }
    $headers = @{ Authorization = "Bearer $(Get-MpesaAccessToken $Configuration)" }
    return Invoke-RestMethod -Method Post `
        -Uri "$(Get-MpesaApiBase $Configuration)/mpesa/stkpushquery/v1/query" `
        -Headers $headers -ContentType "application/json" `
        -Body (ConvertTo-Json -InputObject $payload -Compress) -TimeoutSec 20
}

function Read-JsonBody($Request) {
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    try {
        $json = $reader.ReadToEnd()
        if (-not $json) { return $null }
        try { return ConvertFrom-Json -InputObject $json -ErrorAction Stop }
        catch { throw [System.FormatException]::new("Request body must contain valid JSON.") }
    }
    finally { $reader.Dispose() }
}

function Read-RequestBytes($Request) {
    $stream = New-Object System.IO.MemoryStream
    try {
        $Request.InputStream.CopyTo($stream)
        return ,$stream.ToArray()
    } finally { $stream.Dispose() }
}

function Get-UserByEmail([string]$Email) {
    return @($script:db.users | Where-Object { $_.email -eq $Email } | Select-Object -First 1)
}

function Send-NextstepEmail([string]$To, [string]$Subject, [string]$Body, [string]$ReplyTo) {
    if (-not (Test-Path -LiteralPath $script:mailCredentialPath -PathType Leaf)) {
        throw "Email is not configured. Run Configure-Nextstep-Mail.ps1 as this Windows user."
    }
    $credential = Import-Clixml -LiteralPath $script:mailCredentialPath
    if ($credential -isnot [System.Management.Automation.PSCredential] -or $credential.UserName -ne $script:notificationEmail) {
        throw "The saved email credential is invalid. Run Configure-Nextstep-Mail.ps1 again."
    }

    $message = [System.Net.Mail.MailMessage]::new()
    $client = [System.Net.Mail.SmtpClient]::new("smtp.gmail.com", 587)
    try {
        $message.From = [System.Net.Mail.MailAddress]::new($script:notificationEmail, "Nextstep Academy")
        $message.To.Add($To)
        if ($ReplyTo) { $message.ReplyToList.Add([System.Net.Mail.MailAddress]::new($ReplyTo)) }
        $message.Subject = $Subject
        $message.Body = $Body
        $message.IsBodyHtml = $false
        $client.EnableSsl = $true
        $client.Timeout = 6000
        $client.Credentials = $credential.GetNetworkCredential()
        $client.Send($message)
    }
    finally {
        $client.Dispose()
        $message.Dispose()
    }
}

function Send-AuthEventEmails($User, [string]$Event) {
    $eventName = if ($Event -eq "signup") { "New account created" } else { "Successful login" }
    $occurredAt = [DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'")
    $adminSent = $false
    $adminBody = @"
$eventName
Name: $($User.name)
Email: $($User.email)
Role: $($User.role)
Time: $occurredAt
"@
    try {
        Send-NextstepEmail -To $script:notificationEmail -Subject "Nextstep Academy: $eventName" -Body $adminBody -ReplyTo ""
        $adminSent = $true
    }
    catch {
        Write-Warning "Could not send the owner $Event notification for $($User.email): $($_.Exception.Message)"
    }

    $welcomeSent = $null
    if ($Event -eq "signup") {
        $welcomeSent = $false
        $welcomeBody = @"
Hello $($User.name),

Welcome to Nextstep Academy! Your account has been created successfully. You can now log in to explore learning resources and study materials.

We are glad to have you with us.

Nextstep Academy
"@
        try {
            Send-NextstepEmail -To $User.email -Subject "Welcome to Nextstep Academy" -Body $welcomeBody -ReplyTo ""
            $welcomeSent = $true
        }
        catch {
            Write-Warning "Could not send the welcome email to $($User.email): $($_.Exception.Message)"
        }
    }
    return @{ adminSent = $adminSent; welcomeSent = $welcomeSent }
}

function New-RandomToken([int]$ByteCount) {
    $bytes = New-Object byte[] $ByteCount
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) }
    finally { $rng.Dispose() }
    return [Convert]::ToBase64String($bytes).TrimEnd("=").Replace("+", "-").Replace("/", "_")
}

function Set-NewPassword($User, [string]$Password) {
    $salt = New-Object byte[] 16
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($salt) }
    finally { $rng.Dispose() }
    $derive = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($Password, $salt, 210000)
    try { $passwordHash = [Convert]::ToBase64String($derive.GetBytes(32)) }
    finally { $derive.Dispose() }
    $User.salt = [Convert]::ToBase64String($salt)
    $User.passwordHash = $passwordHash
}

function Get-DocumentById([string]$Id) {
    return @($script:db.documents | Where-Object { $_.id -eq $Id } | Select-Object -First 1)
}

function Get-SessionUser($Request) {
    $cookieHeader = $Request.Headers["Cookie"]
    if ($cookieHeader -notmatch "(?:^|;\s*)nsa_session=([^;]+)") { return $null }
    $token = [System.Uri]::UnescapeDataString($Matches[1])
    $session = $script:sessions[$token]
    if (-not $session -or $session.expires -lt [DateTime]::UtcNow) {
        if ($token) { $script:sessions.Remove($token) }
        return $null
    }
    return Get-UserByEmail $session.email
}

function Set-Session($Context, $User) {
    $random = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($random) }
    finally { $rng.Dispose() }
    $token = [Convert]::ToBase64String($random).TrimEnd("=").Replace("+", "-").Replace("/", "_")
    $script:sessions[$token] = @{ email = $User.email; expires = [DateTime]::UtcNow.AddHours(12) }
    $Context.Response.Headers.Add("Set-Cookie", "nsa_session=$token; Path=/; HttpOnly; SameSite=Strict; Max-Age=43200")
}

function Clear-Session($Context) {
    $cookieHeader = $Context.Request.Headers["Cookie"]
    if ($cookieHeader -match "(?:^|;\s*)nsa_session=([^;]+)") {
        $script:sessions.Remove([System.Uri]::UnescapeDataString($Matches[1]))
    }
    $Context.Response.Headers.Add("Set-Cookie", "nsa_session=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0")
}

function Get-AccountView($User) {
    $activeLeases = @($script:leases.Values | Where-Object { $_.owner -eq $User.email -and $_.expires -gt [DateTime]::UtcNow }).Count
    $uploads = [int]$User.uploads
    $downloads = [int]$User.downloads
    $premiumUntil = $null
    if ($User.premiumUntilUtc) {
        try {
            $premiumUntil = [DateTimeOffset]::Parse([string]$User.premiumUntilUtc).UtcDateTime
        } catch {
            Write-Warning "Ignoring invalid premium expiry for $($User.email)."
        }
    }
    $isPremium = $premiumUntil -and $premiumUntil -gt [DateTime]::UtcNow
    $pendingPayment = @($script:db.payments |
        Where-Object { $_.email -eq $User.email -and $_.status -eq "pending" } |
        Select-Object -First 1)
    return @{
        email = $User.email
        name = $User.name
        role = $User.role
        country = $User.country
        system = $User.system
        level = $User.level
        path = $User.path
        uploadCount = $uploads
        downloadCount = $downloads
        availableDownloads = [Math]::Max(0, $uploads - $downloads - $activeLeases)
        isPremium = [bool]$isPremium
        premiumUntilUtc = if ($isPremium) { $premiumUntil.ToString("o") } else { $null }
        premiumCheckoutConfigured = [bool](Test-Path -LiteralPath $script:mpesaCredentialPath -PathType Leaf)
        pendingPremiumPaymentId = if ($pendingPayment.Count) { [string]$pendingPayment[0].id } else { $null }
    }
}

function Get-DocumentInfo([string]$Id) {
    if ($script:localMaterials.ContainsKey($Id)) {
        $name = $script:localMaterials[$Id]
        $path = Join-Path $materialsRoot $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
        return @{ id = $Id; name = $name; title = [System.IO.Path]::GetFileNameWithoutExtension($name); subject = "Study material"; kind = "Notes"; local = $true; path = $path }
    }
    $document = Get-DocumentById $Id
    if (-not $document) { return $null }
    $path = Join-Path $uploadRoot $document.file
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    return @{ id = $document.id; name = $document.name; title = $document.title; subject = $document.subject; kind = $document.kind; local = $false; path = $path; mime = $document.mime }
}

function Send-Document($Context, $Info, [bool]$Inline) {
    $bytes = [System.IO.File]::ReadAllBytes($Info.path)
    $extension = [System.IO.Path]::GetExtension($Info.name).ToLowerInvariant()
    $mime = if ($script:contentTypes.ContainsKey($extension)) { $script:contentTypes[$extension] } else { "application/octet-stream" }
    $safeName = [System.Uri]::EscapeDataString($Info.name)
    $disposition = if ($Inline -and $extension -in @(".pdf", ".txt", ".csv", ".md")) { "inline" } else { "attachment" }
    $Context.Response.StatusCode = 200
    $Context.Response.ContentType = $mime
    $Context.Response.Headers["Content-Disposition"] = "$disposition; filename*=UTF-8''$safeName"
    $Context.Response.Headers["X-Content-Type-Options"] = "nosniff"
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Get-DocumentList($User) {
    $documents = if ($User.role -eq "tutor") {
        @($script:db.documents | Where-Object { $_.by -eq $User.email })
    } else {
        @($script:db.documents)
    }
    return ,@($documents | ForEach-Object {
        [pscustomobject]@{
            id = $_.id; name = $_.name; title = $_.title; subject = $_.subject; kind = $_.kind
            country = $_.country; system = $_.system; level = $_.level; path = $_.path
            by = $_.by; byName = $_.byName; date = $_.date; size = $_.size
        }
    })
}

function Remove-ExpiredLeases {
    foreach ($key in @($script:leases.Keys)) {
        if ($script:leases[$key].expires -lt [DateTime]::UtcNow) { $script:leases.Remove($key) }
    }
}

function Handle-MpesaCallback($Context) {
    $body = Read-JsonBody $Context.Request
    $callback = $body.Body.stkCallback
    $checkoutId = [string]$callback.CheckoutRequestID
    if (-not $checkoutId) { Send-Error $Context 400 "M-Pesa callback is missing its checkout request ID."; return }
    $payment = Find-PaymentByCheckoutId $checkoutId
    if (-not $payment) { Send-Error $Context 404 "No matching payment request was found."; return }
    if ($payment.status -eq "completed") {
        Send-Json $Context 200 @{ ResultCode = 0; ResultDesc = "Payment already processed." }
        return
    }

    if ([string]$callback.ResultCode -ne "0") {
        $payment.status = "failed"
        $payment.resultDescription = [string]$callback.ResultDesc
        Save-Database
        Send-Json $Context 200 @{ ResultCode = 0; ResultDesc = "Payment result recorded." }
        return
    }

    $configuration = Get-MpesaConfiguration
    $verification = Invoke-MpesaStkQuery $configuration $checkoutId
    if ([string]$verification.ResultCode -ne "0") {
        throw "Safaricom has not confirmed this payment as successful."
    }

    $metadata = @{}
    foreach ($item in @($callback.CallbackMetadata.Item)) {
        if ($item.Name) { $metadata[[string]$item.Name] = $item.Value }
    }
    $paidAmount = 0
    if ($null -ne $metadata.Amount) { $paidAmount = [decimal]$metadata.Amount }
    $paidPhone = Normalize-KenyanPhone ([string]$metadata.PhoneNumber)
    if ($paidAmount -ne $script:premiumPriceKes -or
        $paidPhone -ne $payment.phone -or
        -not $metadata.MpesaReceiptNumber) {
        $payment.status = "review"
        $payment.resultDescription = "Safaricom callback did not match the expected amount, phone, and receipt."
        Save-Database
        Send-Error $Context 400 "The payment details could not be verified. Contact support with your M-Pesa receipt."
        return
    }

    $user = Get-UserByEmail $payment.email
    if (-not $user) { Send-Error $Context 404 "The account for this payment no longer exists."; return }
    $previousExpiry = [string]$user.premiumUntilUtc
    $previousStatus = [string]$payment.status
    $previousReceipt = [string]$payment.mpesaReceiptNumber
    $now = [DateTime]::UtcNow
    $expiry = $now
    if ($previousExpiry) {
        try {
            $existingExpiry = [DateTimeOffset]::Parse($previousExpiry).UtcDateTime
            if ($existingExpiry -gt $now) { $expiry = $existingExpiry }
        } catch {
            Write-Warning "Ignoring invalid premium expiry while renewing $($user.email)."
        }
    }
    $expiry = $expiry.AddDays(30)
    $user.premiumUntilUtc = $expiry.ToString("o")
    $payment.status = "completed"
    $payment.mpesaReceiptNumber = [string]$metadata.MpesaReceiptNumber
    $payment.completedUtc = $now.ToString("o")
    $payment.premiumUntilUtc = $expiry.ToString("o")
    try { Save-Database }
    catch {
        $user.premiumUntilUtc = $previousExpiry
        $payment.status = $previousStatus
        $payment.mpesaReceiptNumber = $previousReceipt
        throw
    }
    Send-Json $Context 200 @{ ResultCode = 0; ResultDesc = "Payment verified and Premium activated." }
}

function Handle-Request($Context) {
    $request = $Context.Request
    $response = $Context.Response
    $method = $request.HttpMethod
    $path = [System.Uri]::UnescapeDataString($request.Url.AbsolutePath)
    $response.Headers["X-Content-Type-Options"] = "nosniff"
    $response.Headers["Referrer-Policy"] = "same-origin"

    if ($method -eq "POST" -and $path -eq "/api/payments/mpesa/callback") {
        Handle-MpesaCallback $Context
        return
    }

    if ($method -eq "GET" -and $path -eq "/") {
        $bytes = [System.IO.File]::ReadAllBytes((Join-Path $root "nextstep-academy.html"))
        $response.ContentType = "text/html; charset=utf-8"
        $response.Headers["Cache-Control"] = "no-store"
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        return
    }
    if ($method -eq "GET" -and $path -match "^/ku-campus-images/([a-z0-9-]+\.jpg)$") {
        $imagePath = Join-Path (Join-Path $root "ku-campus-images") $Matches[1]
        if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) { Send-Error $Context 404 "Image not found."; return }
        $bytes = [System.IO.File]::ReadAllBytes($imagePath)
        $response.ContentType = "image/jpeg"
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        return
    }
    if ($method -eq "GET" -and $path -in @("/manifest.webmanifest", "/service-worker.js", "/nextstep-icon.svg", "/nextstep-icon-192.png", "/nextstep-icon-512.png")) {
        $assetName = $path.TrimStart("/")
        $assetPath = Join-Path $root $assetName
        if (-not (Test-Path -LiteralPath $assetPath -PathType Leaf)) { Send-Error $Context 404 "App asset not found."; return }
        $response.ContentType = switch ($assetName) {
            "manifest.webmanifest" { "application/manifest+json; charset=utf-8" }
            "service-worker.js" { "application/javascript; charset=utf-8" }
            "nextstep-icon.svg" { "image/svg+xml; charset=utf-8" }
            { $_ -in @("nextstep-icon-192.png", "nextstep-icon-512.png") } { "image/png" }
        }
        $response.Headers["Cache-Control"] = if ($assetName -in @("service-worker.js", "manifest.webmanifest")) { "no-cache" } else { "public, max-age=3600" }
        if ($assetName -eq "service-worker.js") { $response.Headers["Service-Worker-Allowed"] = "/" }
        $bytes = [System.IO.File]::ReadAllBytes($assetPath)
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        return
    }

    if ($path -notlike "/api/*") { Send-Error $Context 404 "Page not found."; return }
    Remove-ExpiredLeases

    if ($method -eq "GET" -and $path -eq "/api/mail/status") {
        Send-Json $Context 200 @{ configured = [bool](Test-Path -LiteralPath $script:mailCredentialPath -PathType Leaf) }
        return
    }

    if ($method -eq "POST" -and $path -eq "/api/feedback") {
        $clientAddress = $request.Headers["CF-Connecting-IP"]
        if ($clientAddress -notmatch "^[0-9A-Fa-f:.]{3,45}$") { $clientAddress = $request.RemoteEndPoint.Address.ToString() }
        $now = [DateTime]::UtcNow
        $windowStart = $now.AddMinutes(-10)
        $attempts = @()
        if ($script:feedbackAttempts.ContainsKey($clientAddress)) {
            $attempts = @($script:feedbackAttempts[$clientAddress] | Where-Object { $_ -gt $windowStart })
        }
        if ($attempts.Count -ge 3) { Send-Error $Context 429 "Please wait a few minutes before sending more feedback."; return }
        $script:feedbackAttempts[$clientAddress] = @($attempts) + @($now)

        $body = Read-JsonBody $request
        $name = ([string]$body.name).Trim()
        $email = ([string]$body.email).Trim()
        $message = ([string]$body.message).Trim()
        if (-not $name -or $name.Length -gt 100 -or $email -notmatch "^\S+@\S+\.\S+$" -or $email.Length -gt 254 -or -not $message -or $message.Length -gt 5000) {
            Send-Error $Context 400 "Provide your name, a valid email address, and a message up to 5000 characters."; return
        }
        $feedbackBody = "Name: $name`r`nEmail: $email`r`n`r`n$message"
        try {
            Send-NextstepEmail -To $script:notificationEmail -Subject "Nextstep Academy feedback" -Body $feedbackBody -ReplyTo $email
        }
        catch {
            Write-Warning "Could not send feedback from $($email): $($_.Exception.Message)"
            $message = if (-not (Test-Path -LiteralPath $script:mailCredentialPath -PathType Leaf)) {
                "Email is not configured. Run Configure-Nextstep-Mail.ps1 as this Windows user, then retry sending feedback."
            } else {
                "Email delivery failed. Check the server window and verify the Gmail App Password with Configure-Nextstep-Mail.ps1, then retry."
            }
            Send-Error $Context 503 $message; return
        }
        Send-Json $Context 200 @{ ok = $true }
        return
    }

    if ($method -eq "POST" -and $path -eq "/api/auth/register") {
        $body = Read-JsonBody $request
        $email = ([string]$body.email).Trim().ToLowerInvariant()
        $name = ([string]$body.name).Trim()
        $password = [string]$body.password
        $role = [string]$body.role
        if ($email -notmatch "^\S+@\S+\.\S+$" -or $email.Length -gt 254 -or -not $name -or $name.Length -gt 100 -or $password.Length -lt 8 -or $password.Length -gt 1024 -or $role -notin @("student", "tutor")) {
            Send-Error $Context 400 "Provide a valid name, email, role, and password with at least 8 characters."; return
        }
        if (Get-UserByEmail $email) { Send-Error $Context 409 "That email is already registered."; return }
        $salt = New-Object byte[] 16
        [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($salt)
        $derive = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($password, $salt, 210000)
        try { $passwordHash = [Convert]::ToBase64String($derive.GetBytes(32)) }
        finally { $derive.Dispose() }
        $user = [pscustomobject]@{
            email = $email; name = $name; role = $role
            salt = [Convert]::ToBase64String($salt); passwordHash = $passwordHash
            country = ""; system = ""; level = ""; path = ""; uploads = 0; downloads = 0; premiumUntilUtc = ""
        }
        $script:db.users = @($script:db.users) + @($user)
        try { Save-Database }
        catch {
            $script:db.users = @($script:db.users | Where-Object { $_.email -ne $email })
            throw
        }
        Set-Session $Context $user
        $emailStatus = Send-AuthEventEmails $user "signup"
        Send-Json $Context 201 @{ user = (Get-AccountView $user); adminNotificationSent = $emailStatus.adminSent; welcomeEmailSent = $emailStatus.welcomeSent }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/auth/login") {
        $body = Read-JsonBody $request
        $email = ([string]$body.email).Trim().ToLowerInvariant()
        $user = Get-UserByEmail $email
        if (-not $user) { Send-Error $Context 401 "Incorrect email or password."; return }
        $salt = [Convert]::FromBase64String($user.salt)
        $derive = [System.Security.Cryptography.Rfc2898DeriveBytes]::new([string]$body.password, $salt, 210000)
        try { $candidate = $derive.GetBytes(32) }
        finally { $derive.Dispose() }
        $expected = [Convert]::FromBase64String($user.passwordHash)
        $different = 0
        for ($i = 0; $i -lt $expected.Length; $i++) { $different = $different -bor ($candidate[$i] -bxor $expected[$i]) }
        if ($different -ne 0) { Send-Error $Context 401 "Incorrect email or password."; return }
        Set-Session $Context $user
        Send-AuthEventEmails $user "login"
        Send-Json $Context 200 @{ user = (Get-AccountView $user) }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/auth/forgot-password") {
        $body = Read-JsonBody $request
        $email = ([string]$body.email).Trim().ToLowerInvariant()
        if ($email -notmatch "^\S+@\S+\.\S+$" -or $email.Length -gt 254) {
            Send-Error $Context 400 "Enter a valid email address."; return
        }

        $user = Get-UserByEmail $email
        if ($user) {
            foreach ($key in @($script:passwordResets.Keys)) {
                if ($script:passwordResets[$key].expires -lt [DateTime]::UtcNow) {
                    $script:passwordResets.Remove($key)
                }
            }
            $token = New-RandomToken 32
            $script:passwordResets[$token] = @{ email = $email; expires = [DateTime]::UtcNow.AddMinutes(30) }

            $baseUrl = [string]$request.Headers["Origin"]
            if (-not $baseUrl -and $request.Headers["Referer"]) {
                try {
                    $referer = [System.Uri]::new([string]$request.Headers["Referer"])
                    $baseUrl = "$($referer.Scheme)://$($referer.Authority)"
                } catch { $baseUrl = "" }
            }
            if (-not $baseUrl) { $baseUrl = $Prefix.TrimEnd("/") }
            $resetUrl = "$($baseUrl.TrimEnd('/'))/?reset=$([System.Uri]::EscapeDataString($token))"
            $emailBody = @"
Hello $($user.name),

We received a request to reset your Nextstep Academy password.

Reset your password here:
$resetUrl

This link expires in 30 minutes. If you did not request it, you can ignore this email.

Nextstep Academy
"@
            try {
                Send-NextstepEmail -To $email -Subject "Reset your Nextstep Academy password" -Body $emailBody -ReplyTo ""
            }
            catch {
                $script:passwordResets.Remove($token)
                Write-Warning "Could not send password reset email to $email`: $($_.Exception.Message)"
                $message = if (-not (Test-Path -LiteralPath $script:mailCredentialPath -PathType Leaf)) {
                    "Password reset email is not configured. Run Configure-Nextstep-Mail.ps1 on this computer."
                } else {
                    "The password reset email could not be sent. Check the server email configuration and try again."
                }
                Send-Error $Context 503 $message
                return
            }
        }
        Send-Json $Context 200 @{ ok = $true; message = "If an account uses that email, a password reset link has been sent." }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/auth/reset-password") {
        $body = Read-JsonBody $request
        $token = [string]$body.token
        $password = [string]$body.password
        if ($password.Length -lt 8 -or $password.Length -gt 1024) {
            Send-Error $Context 400 "Password needs 8 to 1024 characters."; return
        }
        $reset = $script:passwordResets[$token]
        if (-not $reset -or $reset.expires -lt [DateTime]::UtcNow) {
            if ($token) { $script:passwordResets.Remove($token) }
            Send-Error $Context 400 "This password reset link is invalid or has expired."; return
        }
        $user = Get-UserByEmail ([string]$reset.email)
        if (-not $user) {
            $script:passwordResets.Remove($token)
            Send-Error $Context 400 "This password reset link is no longer valid."; return
        }
        Set-NewPassword $user $password
        try { Save-Database }
        catch {
            Send-Error $Context 500 "The password could not be saved. Please try again."; return
        }
        $script:passwordResets.Remove($token)
        foreach ($sessionKey in @($script:sessions.Keys)) {
            if ($script:sessions[$sessionKey].email -eq $user.email) { $script:sessions.Remove($sessionKey) }
        }
        Send-Json $Context 200 @{ ok = $true; message = "Your password has been reset. You can now log in." }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/auth/logout") {
        Clear-Session $Context
        Send-Json $Context 200 @{ ok = $true }
        return
    }
    if ($method -eq "GET" -and $path -eq "/api/auth/me") {
        $user = Get-SessionUser $request
        if ($user) { Send-Json $Context 200 @{ user = (Get-AccountView $user) } }
        else { Send-Json $Context 200 @{ user = $null } }
        return
    }

    $user = Get-SessionUser $request
    if (-not $user) { Send-Error $Context 401 "Please log in to continue."; return }
    if ($method -eq "GET" -and $path -eq "/api/account") {
        Send-Json $Context 200 @{ account = (Get-AccountView $user) }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/premium/checkout") {
        $configuration = $null
        try { $configuration = Get-MpesaConfiguration }
        catch {
            Send-Error $Context 503 $_.Exception.Message
            return
        }
        $body = Read-JsonBody $request
        $phone = Normalize-KenyanPhone ([string]$body.phone)
        if (-not $phone) {
            Send-Error $Context 400 "Enter a valid Kenyan mobile number, such as 0712345678."
            return
        }
        $recentPending = $false
        foreach ($pendingPayment in @($script:db.payments | Where-Object { $_.email -eq $user.email -and $_.status -eq "pending" })) {
            try {
                if ([DateTimeOffset]::Parse([string]$pendingPayment.createdUtc).UtcDateTime -gt [DateTime]::UtcNow.AddMinutes(-5)) {
                    $recentPending = $true
                    break
                }
            } catch { }
        }
        if ($recentPending) {
            Send-Error $Context 409 "A recent M-Pesa request is still pending. Complete it on your phone or wait a few minutes before trying again."
            return
        }

        $payment = [pscustomobject]@{
            id = [Guid]::NewGuid().ToString("N")
            email = $user.email
            phone = $phone
            amount = $script:premiumPriceKes
            status = "pending"
            createdUtc = [DateTime]::UtcNow.ToString("o")
            checkoutRequestId = ""
            merchantRequestId = ""
            mpesaReceiptNumber = ""
        }
        $script:db.payments = @($script:db.payments) + @($payment)
        try { Save-Database }
        catch {
            $script:db.payments = @($script:db.payments | Where-Object { $_.id -ne $payment.id })
            Send-Error $Context 500 "Could not save your payment request. No M-Pesa prompt was sent."
            return
        }

        try {
            $timestamp = (Get-Date).ToString("yyyyMMddHHmmss")
            $payload = @{
                BusinessShortCode = [string]$configuration.BusinessShortCode
                Password = Get-MpesaPassword $configuration $timestamp
                Timestamp = $timestamp
                TransactionType = [string]$configuration.TransactionType
                Amount = $script:premiumPriceKes
                PartyA = $phone
                PartyB = [string]$configuration.BusinessShortCode
                PhoneNumber = $phone
                CallBackURL = [string]$configuration.CallbackUrl
                AccountReference = "NEXTSTEP"
                TransactionDesc = "Nextstep Premium 30-day access"
            }
            $headers = @{ Authorization = "Bearer $(Get-MpesaAccessToken $configuration)" }
            $push = Invoke-RestMethod -Method Post `
                -Uri "$(Get-MpesaApiBase $configuration)/mpesa/stkpush/v1/processrequest" `
                -Headers $headers -ContentType "application/json" `
                -Body (ConvertTo-Json -InputObject $payload -Compress) -TimeoutSec 30
            if ([string]$push.ResponseCode -ne "0" -or -not $push.CheckoutRequestID) {
                $payment.status = "failed"
                $payment.resultDescription = [string]$push.ResponseDescription
                Save-Database
                Send-Error $Context 502 "Safaricom could not start the M-Pesa prompt. $([string]$push.ResponseDescription)"
                return
            }
            $payment.checkoutRequestId = [string]$push.CheckoutRequestID
            $payment.merchantRequestId = [string]$push.MerchantRequestID
            Save-Database
            Send-Json $Context 202 @{ checkoutId = $payment.id; customerMessage = [string]$push.CustomerMessage }
        }
        catch {
            $payment.status = "failed"
            $payment.resultDescription = "M-Pesa request could not be started."
            try { Save-Database } catch { Write-Warning "Could not record the failed M-Pesa request: $($_.Exception.Message)" }
            Write-Warning "Could not start Premium checkout for $($user.email): $($_.Exception.Message)"
            Send-Error $Context 502 "M-Pesa checkout could not be started. Check the Daraja configuration and try again."
        }
        return
    }
    if ($method -eq "GET" -and $path -match "^/api/premium/status/([a-fA-F0-9]{32})$") {
        $paymentId = $Matches[1]
        $payment = @($script:db.payments | Where-Object { $_.id -eq $paymentId -and $_.email -eq $user.email } | Select-Object -First 1)
        if (-not $payment) {
            Send-Error $Context 404 "That payment request was not found."
            return
        }
        Send-Json $Context 200 @{ status = $payment.status; account = (Get-AccountView $user) }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/profile") {
        $body = Read-JsonBody $request
        $profile = @{
            country = ([string]$body.country).Trim()
            system = ([string]$body.system).Trim()
            level = ([string]$body.level).Trim()
            path = ([string]$body.path).Trim()
        }
        if (-not $profile.country -or -not $profile.system -or -not $profile.level -or $profile.country.Length -gt 100 -or $profile.system.Length -gt 100 -or $profile.level.Length -gt 100 -or $profile.path.Length -gt 100) {
            Send-Error $Context 400 "Choose a country, education system, and level."; return
        }
        $previousProfile = @{}
        foreach ($field in $profile.Keys) {
            $previousProfile[$field] = $user.$field
            $user.$field = $profile[$field]
        }
        try { Save-Database }
        catch {
            foreach ($field in $previousProfile.Keys) { $user.$field = $previousProfile[$field] }
            throw
        }
        Send-Json $Context 200 @{ user = (Get-AccountView $user) }
        return
    }
    if ($method -eq "GET" -and $path -eq "/api/documents") {
        Send-Json $Context 200 @{ documents = (Get-DocumentList $user); account = (Get-AccountView $user) }
        return
    }
    if ($method -eq "DELETE" -and $path -match "^/api/documents/([a-fA-F0-9]+)$") {
        $document = Get-DocumentById $Matches[1]
        if (-not $document -or $document.by -ne $user.email) { Send-Error $Context 404 "That uploaded document is not available to delete."; return }
        $filePath = Join-Path $uploadRoot $document.file
        $previousDocuments = @($script:db.documents)
        $script:db.documents = @($script:db.documents | Where-Object { $_.id -ne $document.id })
        try { Save-Database }
        catch { $script:db.documents = $previousDocuments; throw }
        if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force }
        Send-Json $Context 200 @{ ok = $true }
        return
    }
    if ($method -eq "POST" -and $path -eq "/api/documents") {
        $metadataHeader = $request.Headers["X-Document-Metadata"]
        if (-not $metadataHeader) { Send-Error $Context 400 "Document details are missing."; return }
        try { $metadata = ConvertFrom-Json -InputObject ([System.Uri]::UnescapeDataString($metadataHeader)) -ErrorAction Stop }
        catch { Send-Error $Context 400 "Document details must be valid JSON."; return }
        $title = ([string]$metadata.title).Trim()
        $subject = ([string]$metadata.subject).Trim()
        $name = [System.IO.Path]::GetFileName([string]$metadata.fileName)
        $extension = [System.IO.Path]::GetExtension($name).ToLowerInvariant()
        if (-not $title -or -not $subject -or -not $name -or $extension -notin $script:allowedExtensions) {
            Send-Error $Context 400 "Add a title, subject, and a supported document file."; return
        }
        if ($request.ContentLength64 -gt 9437184) { Send-Error $Context 413 "Files must be 9 MB or smaller."; return }
        $bytes = Read-RequestBytes $request
        if ($bytes.Length -eq 0 -or $bytes.Length -gt 9437184) { Send-Error $Context 400 "The document is empty or exceeds the 9 MB limit."; return }
        $id = [Guid]::NewGuid().ToString("N")
        $storedName = "$id$extension"
        $tempPath = Join-Path $uploadRoot "$id.tmp"
        $storedPath = Join-Path $uploadRoot $storedName
        [System.IO.File]::WriteAllBytes($tempPath, $bytes)
        if ((Get-Item -LiteralPath $tempPath).Length -ne $bytes.Length) {
            Remove-Item -LiteralPath $tempPath -Force
            Send-Error $Context 500 "The uploaded document could not be verified."; return
        }
        Move-Item -LiteralPath $tempPath -Destination $storedPath
        $document = [pscustomobject]@{
            id = $id; file = $storedName; name = $name; title = $title; subject = $subject
            kind = if ($metadata.kind -in @("Notes", "Past paper", "Other")) { [string]$metadata.kind } else { "Notes" }
            country = [string]$user.country; system = [string]$user.system; level = [string]$user.level
            path = [string]$user.path; by = $user.email; byName = $user.name; date = (Get-Date -Format "yyyy-MM-dd")
            mime = $script:contentTypes[$extension]; size = $bytes.Length
        }
        $script:db.documents = @($script:db.documents) + @($document)
        $user.uploads = [int]$user.uploads + 1
        try { Save-Database }
        catch {
            $script:db.documents = @($script:db.documents | Where-Object { $_.id -ne $id })
            $user.uploads = [Math]::Max(0, [int]$user.uploads - 1)
            Remove-Item -LiteralPath $storedPath -Force -ErrorAction SilentlyContinue
            throw
        }
        Send-Json $Context 201 @{ document = $document; account = (Get-AccountView $user); verified = $true }
        return
    }
    if ($method -eq "POST" -and $path -match "^/api/downloads/start/([a-zA-Z0-9-]+)$") {
        $info = Get-DocumentInfo $Matches[1]
        if (-not $info) { Send-Error $Context 404 "That document is not available on the server."; return }
        $account = Get-AccountView $user
        if (-not $account.isPremium -and $account.availableDownloads -lt 1) {
            Send-Error $Context 403 "Upload a document to unlock one download, or get Premium for unlimited downloads."; return
        }
        $leaseId = [Guid]::NewGuid().ToString("N")
        $script:leases[$leaseId] = @{ owner = $user.email; documentId = $Matches[1]; expires = [DateTime]::UtcNow.AddMinutes(10); served = $false; premium = $account.isPremium }
        Send-Json $Context 200 @{ leaseId = $leaseId; fileUrl = "/api/downloads/file/$leaseId" }
        return
    }
    if ($method -eq "GET" -and $path -match "^/api/downloads/file/([a-fA-F0-9]+)$") {
        $lease = $script:leases[$Matches[1]]
        if (-not $lease -or $lease.owner -ne $user.email -or $lease.expires -lt [DateTime]::UtcNow) { Send-Error $Context 410 "This download request has expired. Please try again."; return }
        if ($lease.served) { Send-Error $Context 409 "This download request was already used. Please try again."; return }
        $info = Get-DocumentInfo $lease.documentId
        if (-not $info) { Send-Error $Context 404 "That document is no longer available."; return }
        $lease.served = $true
        Send-Document $Context $info $false
        return
    }
    if ($method -eq "POST" -and $path -match "^/api/downloads/(complete|cancel)/([a-fA-F0-9]+)$") {
        $action = $Matches[1]
        $leaseId = $Matches[2]
        $lease = $script:leases[$leaseId]
        if (-not $lease -or $lease.owner -ne $user.email) { Send-Error $Context 404 "This download request is no longer active."; return }
        if ($action -eq "complete") {
            if (-not $lease.served) { Send-Error $Context 409 "The document transfer has not completed."; return }
            if (-not $lease.premium) {
                $previousDownloads = [int]$user.downloads
                $user.downloads = $previousDownloads + 1
                try { Save-Database }
                catch { $user.downloads = $previousDownloads; throw }
            }
        }
        $script:leases.Remove($leaseId)
        Send-Json $Context 200 @{ ok = $true; account = (Get-AccountView $user) }
        return
    }
    if ($method -eq "GET" -and $path -match "^/api/documents/([a-zA-Z0-9-]+)/open$") {
        $info = Get-DocumentInfo $Matches[1]
        if (-not $info) { Send-Error $Context 404 "That document is not available on the server."; return }
        Send-Document $Context $info $true
        return
    }
    Send-Error $Context 404 "API route not found."
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Prefix)
try {
    $listener.Start()
} catch {
    $startError = $_.Exception.Message
    try {
        $running = Invoke-WebRequest -UseBasicParsing -Uri $Prefix -TimeoutSec 2
        if ($running.StatusCode -eq 200 -and $running.Content -match '<title>Nextstep Academy</title>') {
            Write-Host "Nextstep Academy is already running at $Prefix"
            if ($env:NEXTSTEP_OPEN_BROWSER -eq "1") { Start-Process $Prefix }
            exit 0
        }
    } catch { }
    Write-Error "Could not start the local server at $Prefix. Close any app using this port, then try again. $startError"
    exit 1
}
Write-Host "Nextstep Academy is running at $Prefix"
Write-Host "Press Ctrl+C to stop the server."
if ($env:NEXTSTEP_OPEN_BROWSER -eq "1") { Start-Process $Prefix }

while ($listener.IsListening) {
    $context = $null
    try {
        $context = $listener.GetContext()
        Handle-Request $context
    } catch {
        Write-Host "Request error: $($_.Exception.Message)"
        if ($context -and -not $context.Response.OutputStream.CanWrite) { }
        elseif ($context) {
            try {
                if ($_.Exception -is [System.FormatException]) { Send-Error $context 400 "Request body must contain valid JSON." }
                else { Send-Error $context 500 "The server could not complete this request." }
            } catch { }
        }
    } finally {
        if ($context) { try { $context.Response.Close() } catch { } }
    }
}
$listener.Stop()
$listener.Close()
