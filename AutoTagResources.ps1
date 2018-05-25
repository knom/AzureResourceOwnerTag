function Send-MailMessageEx {
    # Helper function for sending mails with attachments
    # From https://gallery.technet.microsoft.com/scriptcenter/Send-MailMessage-3a920a6d
    # By David Wyatt
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [Alias('PsPath')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Attachments,

        [ValidateNotNullOrEmpty()]
        [Collections.HashTable]
        $InlineAttachments,
    
        [ValidateNotNullOrEmpty()]
        [Net.Mail.MailAddress[]]
        $Bcc,

        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,
    
        [Alias('BAH')]
        [switch]
        $BodyAsHtml,

        [ValidateNotNullOrEmpty()]
        [Net.Mail.MailAddress[]]
        $Cc,

        [Alias('DNO')]
        [ValidateNotNullOrEmpty()]
        [Net.Mail.DeliveryNotificationOptions]
        $DeliveryNotificationOption,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Net.Mail.MailAddress]
        $From,

        [Parameter(Mandatory = $true, Position = 3)]
        [Alias('ComputerName')]
        [string]
        $SmtpServer,

        [ValidateNotNullOrEmpty()]
        [Net.Mail.MailPriority]
        $Priority,
    
        [Parameter(Mandatory = $true, Position = 1)]
        [Alias('sub')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Subject,

        [Parameter(Mandatory = $true, Position = 0)]
        [Net.Mail.MailAddress[]]
        $To,

        [ValidateNotNullOrEmpty()]
        [Management.Automation.PSCredential]
        $Credential,

        [switch]
        $UseSsl,

        [ValidateRange(0, 2147483647)]
        [int]
        $Port = 25
    )

    begin {
        function FileNameToContentType {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string]
                $FileName
            )

            $mimeMappings = @{
                '.323'         = 'text/h323'
                '.aaf'         = 'application/octet-stream'
                '.aca'         = 'application/octet-stream'
                '.accdb'       = 'application/msaccess'
                '.accde'       = 'application/msaccess'
                '.accdt'       = 'application/msaccess'
                '.acx'         = 'application/internet-property-stream'
                '.afm'         = 'application/octet-stream'
                '.ai'          = 'application/postscript'
                '.aif'         = 'audio/x-aiff'
                '.aifc'        = 'audio/aiff'
                '.aiff'        = 'audio/aiff'
                '.application' = 'application/x-ms-application'
                '.art'         = 'image/x-jg'
                '.asd'         = 'application/octet-stream'
                '.asf'         = 'video/x-ms-asf'
                '.asi'         = 'application/octet-stream'
                '.asm'         = 'text/plain'
                '.asr'         = 'video/x-ms-asf'
                '.asx'         = 'video/x-ms-asf'
                '.atom'        = 'application/atom+xml'
                '.au'          = 'audio/basic'
                '.avi'         = 'video/x-msvideo'
                '.axs'         = 'application/olescript'
                '.bas'         = 'text/plain'
                '.bcpio'       = 'application/x-bcpio'
                '.bin'         = 'application/octet-stream'
                '.bmp'         = 'image/bmp'
                '.c'           = 'text/plain'
                '.cab'         = 'application/octet-stream'
                '.calx'        = 'application/vnd.ms-office.calx'
                '.cat'         = 'application/vnd.ms-pki.seccat'
                '.cdf'         = 'application/x-cdf'
                '.chm'         = 'application/octet-stream'
                '.class'       = 'application/x-java-applet'
                '.clp'         = 'application/x-msclip'
                '.cmx'         = 'image/x-cmx'
                '.cnf'         = 'text/plain'
                '.cod'         = 'image/cis-cod'
                '.cpio'        = 'application/x-cpio'
                '.cpp'         = 'text/plain'
                '.crd'         = 'application/x-mscardfile'
                '.crl'         = 'application/pkix-crl'
                '.crt'         = 'application/x-x509-ca-cert'
                '.csh'         = 'application/x-csh'
                '.css'         = 'text/css'
                '.csv'         = 'application/octet-stream'
                '.cur'         = 'application/octet-stream'
                '.dcr'         = 'application/x-director'
                '.deploy'      = 'application/octet-stream'
                '.der'         = 'application/x-x509-ca-cert'
                '.dib'         = 'image/bmp'
                '.dir'         = 'application/x-director'
                '.disco'       = 'text/xml'
                '.dll'         = 'application/x-msdownload'
                '.dll.config'  = 'text/xml'
                '.dlm'         = 'text/dlm'
                '.doc'         = 'application/msword'
                '.docm'        = 'application/vnd.ms-word.document.macroEnabled.12'
                '.docx'        = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                '.dot'         = 'application/msword'
                '.dotm'        = 'application/vnd.ms-word.template.macroEnabled.12'
                '.dotx'        = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template'
                '.dsp'         = 'application/octet-stream'
                '.dtd'         = 'text/xml'
                '.dvi'         = 'application/x-dvi'
                '.dwf'         = 'drawing/x-dwf'
                '.dwp'         = 'application/octet-stream'
                '.dxr'         = 'application/x-director'
                '.eml'         = 'message/rfc822'
                '.emz'         = 'application/octet-stream'
                '.eot'         = 'application/octet-stream'
                '.eps'         = 'application/postscript'
                '.etx'         = 'text/x-setext'
                '.evy'         = 'application/envoy'
                '.exe'         = 'application/octet-stream'
                '.exe.config'  = 'text/xml'
                '.fdf'         = 'application/vnd.fdf'
                '.fif'         = 'application/fractals'
                '.fla'         = 'application/octet-stream'
                '.flr'         = 'x-world/x-vrml'
                '.flv'         = 'video/x-flv'
                '.gif'         = 'image/gif'
                '.gtar'        = 'application/x-gtar'
                '.gz'          = 'application/x-gzip'
                '.h'           = 'text/plain'
                '.hdf'         = 'application/x-hdf'
                '.hdml'        = 'text/x-hdml'
                '.hhc'         = 'application/x-oleobject'
                '.hhk'         = 'application/octet-stream'
                '.hhp'         = 'application/octet-stream'
                '.hlp'         = 'application/winhlp'
                '.hqx'         = 'application/mac-binhex40'
                '.hta'         = 'application/hta'
                '.htc'         = 'text/x-component'
                '.htm'         = 'text/html'
                '.html'        = 'text/html'
                '.htt'         = 'text/webviewhtml'
                '.hxt'         = 'text/html'
                '.ico'         = 'image/x-icon'
                '.ics'         = 'application/octet-stream'
                '.ief'         = 'image/ief'
                '.iii'         = 'application/x-iphone'
                '.inf'         = 'application/octet-stream'
                '.ins'         = 'application/x-internet-signup'
                '.isp'         = 'application/x-internet-signup'
                '.IVF'         = 'video/x-ivf'
                '.jar'         = 'application/java-archive'
                '.java'        = 'application/octet-stream'
                '.jck'         = 'application/liquidmotion'
                '.jcz'         = 'application/liquidmotion'
                '.jfif'        = 'image/pjpeg'
                '.jpb'         = 'application/octet-stream'
                '.jpe'         = 'image/jpeg'
                '.jpeg'        = 'image/jpeg'
                '.jpg'         = 'image/jpeg'
                '.js'          = 'application/x-javascript'
                '.jsx'         = 'text/jscript'
                '.latex'       = 'application/x-latex'
                '.lit'         = 'application/x-ms-reader'
                '.lpk'         = 'application/octet-stream'
                '.lsf'         = 'video/x-la-asf'
                '.lsx'         = 'video/x-la-asf'
                '.lzh'         = 'application/octet-stream'
                '.m13'         = 'application/x-msmediaview'
                '.m14'         = 'application/x-msmediaview'
                '.m1v'         = 'video/mpeg'
                '.m3u'         = 'audio/x-mpegurl'
                '.man'         = 'application/x-troff-man'
                '.manifest'    = 'application/x-ms-manifest'
                '.map'         = 'text/plain'
                '.mdb'         = 'application/x-msaccess'
                '.mdp'         = 'application/octet-stream'
                '.me'          = 'application/x-troff-me'
                '.mht'         = 'message/rfc822'
                '.mhtml'       = 'message/rfc822'
                '.mid'         = 'audio/mid'
                '.midi'        = 'audio/mid'
                '.mix'         = 'application/octet-stream'
                '.mmf'         = 'application/x-smaf'
                '.mno'         = 'text/xml'
                '.mny'         = 'application/x-msmoney'
                '.mov'         = 'video/quicktime'
                '.movie'       = 'video/x-sgi-movie'
                '.mp2'         = 'video/mpeg'
                '.mp3'         = 'audio/mpeg'
                '.mpa'         = 'video/mpeg'
                '.mpe'         = 'video/mpeg'
                '.mpeg'        = 'video/mpeg'
                '.mpg'         = 'video/mpeg'
                '.mpp'         = 'application/vnd.ms-project'
                '.mpv2'        = 'video/mpeg'
                '.ms'          = 'application/x-troff-ms'
                '.msi'         = 'application/octet-stream'
                '.mso'         = 'application/octet-stream'
                '.mvb'         = 'application/x-msmediaview'
                '.mvc'         = 'application/x-miva-compiled'
                '.nc'          = 'application/x-netcdf'
                '.nsc'         = 'video/x-ms-asf'
                '.nws'         = 'message/rfc822'
                '.ocx'         = 'application/octet-stream'
                '.oda'         = 'application/oda'
                '.odc'         = 'text/x-ms-odc'
                '.ods'         = 'application/oleobject'
                '.one'         = 'application/onenote'
                '.onea'        = 'application/onenote'
                '.onetoc'      = 'application/onenote'
                '.onetoc2'     = 'application/onenote'
                '.onetmp'      = 'application/onenote'
                '.onepkg'      = 'application/onenote'
                '.osdx'        = 'application/opensearchdescription+xml'
                '.p10'         = 'application/pkcs10'
                '.p12'         = 'application/x-pkcs12'
                '.p7b'         = 'application/x-pkcs7-certificates'
                '.p7c'         = 'application/pkcs7-mime'
                '.p7m'         = 'application/pkcs7-mime'
                '.p7r'         = 'application/x-pkcs7-certreqresp'
                '.p7s'         = 'application/pkcs7-signature'
                '.pbm'         = 'image/x-portable-bitmap'
                '.pcx'         = 'application/octet-stream'
                '.pcz'         = 'application/octet-stream'
                '.pdf'         = 'application/pdf'
                '.pfb'         = 'application/octet-stream'
                '.pfm'         = 'application/octet-stream'
                '.pfx'         = 'application/x-pkcs12'
                '.pgm'         = 'image/x-portable-graymap'
                '.pko'         = 'application/vnd.ms-pki.pko'
                '.pma'         = 'application/x-perfmon'
                '.pmc'         = 'application/x-perfmon'
                '.pml'         = 'application/x-perfmon'
                '.pmr'         = 'application/x-perfmon'
                '.pmw'         = 'application/x-perfmon'
                '.png'         = 'image/png'
                '.pnm'         = 'image/x-portable-anymap'
                '.pnz'         = 'image/png'
                '.pot'         = 'application/vnd.ms-powerpoint'
                '.potm'        = 'application/vnd.ms-powerpoint.template.macroEnabled.12'
                '.potx'        = 'application/vnd.openxmlformats-officedocument.presentationml.template'
                '.ppam'        = 'application/vnd.ms-powerpoint.addin.macroEnabled.12'
                '.ppm'         = 'image/x-portable-pixmap'
                '.pps'         = 'application/vnd.ms-powerpoint'
                '.ppsm'        = 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12'
                '.ppsx'        = 'application/vnd.openxmlformats-officedocument.presentationml.slideshow'
                '.ppt'         = 'application/vnd.ms-powerpoint'
                '.pptm'        = 'application/vnd.ms-powerpoint.presentation.macroEnabled.12'
                '.pptx'        = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                '.prf'         = 'application/pics-rules'
                '.prm'         = 'application/octet-stream'
                '.prx'         = 'application/octet-stream'
                '.ps'          = 'application/postscript'
                '.psd'         = 'application/octet-stream'
                '.psm'         = 'application/octet-stream'
                '.psp'         = 'application/octet-stream'
                '.pub'         = 'application/x-mspublisher'
                '.qt'          = 'video/quicktime'
                '.qtl'         = 'application/x-quicktimeplayer'
                '.qxd'         = 'application/octet-stream'
                '.ra'          = 'audio/x-pn-realaudio'
                '.ram'         = 'audio/x-pn-realaudio'
                '.rar'         = 'application/octet-stream'
                '.ras'         = 'image/x-cmu-raster'
                '.rf'          = 'image/vnd.rn-realflash'
                '.rgb'         = 'image/x-rgb'
                '.rm'          = 'application/vnd.rn-realmedia'
                '.rmi'         = 'audio/mid'
                '.roff'        = 'application/x-troff'
                '.rpm'         = 'audio/x-pn-realaudio-plugin'
                '.rtf'         = 'application/rtf'
                '.rtx'         = 'text/richtext'
                '.scd'         = 'application/x-msschedule'
                '.sct'         = 'text/scriptlet'
                '.sea'         = 'application/octet-stream'
                '.setpay'      = 'application/set-payment-initiation'
                '.setreg'      = 'application/set-registration-initiation'
                '.sgml'        = 'text/sgml'
                '.sh'          = 'application/x-sh'
                '.shar'        = 'application/x-shar'
                '.sit'         = 'application/x-stuffit'
                '.sldm'        = 'application/vnd.ms-powerpoint.slide.macroEnabled.12'
                '.sldx'        = 'application/vnd.openxmlformats-officedocument.presentationml.slide'
                '.smd'         = 'audio/x-smd'
                '.smi'         = 'application/octet-stream'
                '.smx'         = 'audio/x-smd'
                '.smz'         = 'audio/x-smd'
                '.snd'         = 'audio/basic'
                '.snp'         = 'application/octet-stream'
                '.spc'         = 'application/x-pkcs7-certificates'
                '.spl'         = 'application/futuresplash'
                '.src'         = 'application/x-wais-source'
                '.ssm'         = 'application/streamingmedia'
                '.sst'         = 'application/vnd.ms-pki.certstore'
                '.stl'         = 'application/vnd.ms-pki.stl'
                '.sv4cpio'     = 'application/x-sv4cpio'
                '.sv4crc'      = 'application/x-sv4crc'
                '.swf'         = 'application/x-shockwave-flash'
                '.t'           = 'application/x-troff'
                '.tar'         = 'application/x-tar'
                '.tcl'         = 'application/x-tcl'
                '.tex'         = 'application/x-tex'
                '.texi'        = 'application/x-texinfo'
                '.texinfo'     = 'application/x-texinfo'
                '.tgz'         = 'application/x-compressed'
                '.thmx'        = 'application/vnd.ms-officetheme'
                '.thn'         = 'application/octet-stream'
                '.tif'         = 'image/tiff'
                '.tiff'        = 'image/tiff'
                '.toc'         = 'application/octet-stream'
                '.tr'          = 'application/x-troff'
                '.trm'         = 'application/x-msterminal'
                '.tsv'         = 'text/tab-separated-values'
                '.ttf'         = 'application/octet-stream'
                '.txt'         = 'text/plain'
                '.u32'         = 'application/octet-stream'
                '.uls'         = 'text/iuls'
                '.ustar'       = 'application/x-ustar'
                '.vbs'         = 'text/vbscript'
                '.vcf'         = 'text/x-vcard'
                '.vcs'         = 'text/plain'
                '.vdx'         = 'application/vnd.ms-visio.viewer'
                '.vml'         = 'text/xml'
                '.vsd'         = 'application/vnd.visio'
                '.vss'         = 'application/vnd.visio'
                '.vst'         = 'application/vnd.visio'
                '.vsto'        = 'application/x-ms-vsto'
                '.vsw'         = 'application/vnd.visio'
                '.vsx'         = 'application/vnd.visio'
                '.vtx'         = 'application/vnd.visio'
                '.wav'         = 'audio/wav'
                '.wax'         = 'audio/x-ms-wax'
                '.wbmp'        = 'image/vnd.wap.wbmp'
                '.wcm'         = 'application/vnd.ms-works'
                '.wdb'         = 'application/vnd.ms-works'
                '.wks'         = 'application/vnd.ms-works'
                '.wm'          = 'video/x-ms-wm'
                '.wma'         = 'audio/x-ms-wma'
                '.wmd'         = 'application/x-ms-wmd'
                '.wmf'         = 'application/x-msmetafile'
                '.wml'         = 'text/vnd.wap.wml'
                '.wmlc'        = 'application/vnd.wap.wmlc'
                '.wmls'        = 'text/vnd.wap.wmlscript'
                '.wmlsc'       = 'application/vnd.wap.wmlscriptc'
                '.wmp'         = 'video/x-ms-wmp'
                '.wmv'         = 'video/x-ms-wmv'
                '.wmx'         = 'video/x-ms-wmx'
                '.wmz'         = 'application/x-ms-wmz'
                '.wps'         = 'application/vnd.ms-works'
                '.wri'         = 'application/x-mswrite'
                '.wrl'         = 'x-world/x-vrml'
                '.wrz'         = 'x-world/x-vrml'
                '.wsdl'        = 'text/xml'
                '.wvx'         = 'video/x-ms-wvx'
                '.x'           = 'application/directx'
                '.xaf'         = 'x-world/x-vrml'
                '.xaml'        = 'application/xaml+xml'
                '.xap'         = 'application/x-silverlight-app'
                '.xbap'        = 'application/x-ms-xbap'
                '.xbm'         = 'image/x-xbitmap'
                '.xdr'         = 'text/plain'
                '.xla'         = 'application/vnd.ms-excel'
                '.xlam'        = 'application/vnd.ms-excel.addin.macroEnabled.12'
                '.xlc'         = 'application/vnd.ms-excel'
                '.xlm'         = 'application/vnd.ms-excel'
                '.xls'         = 'application/vnd.ms-excel'
                '.xlsb'        = 'application/vnd.ms-excel.sheet.binary.macroEnabled.12'
                '.xlsm'        = 'application/vnd.ms-excel.sheet.macroEnabled.12'
                '.xlsx'        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                '.xlt'         = 'application/vnd.ms-excel'
                '.xltm'        = 'application/vnd.ms-excel.template.macroEnabled.12'
                '.xltx'        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template'
                '.xlw'         = 'application/vnd.ms-excel'
                '.xml'         = 'text/xml'
                '.xof'         = 'x-world/x-vrml'
                '.xpm'         = 'image/x-xpixmap'
                '.xps'         = 'application/vnd.ms-xpsdocument'
                '.xsd'         = 'text/xml'
                '.xsf'         = 'text/xml'
                '.xsl'         = 'text/xml'
                '.xslt'        = 'text/xml'
                '.xsn'         = 'application/octet-stream'
                '.xtp'         = 'application/octet-stream'
                '.xwd'         = 'image/x-xwindowdump'
                '.z'           = 'application/x-compress'
                '.zip'         = 'application/x-zip-compressed'
            }

            $extension = [System.IO.Path]::GetExtension($FileName)
            $contentType = $mimeMappings[$extension]

            if ([string]::IsNullOrEmpty($contentType)) {
                return New-Object System.Net.Mime.ContentType
            }
            else {
                return New-Object System.Net.Mime.ContentType($contentType)
            }
        }

        try {
            $_smtpClient = New-Object Net.Mail.SmtpClient
    
            $_smtpClient.Host = $SmtpServer
            $_smtpClient.Port = $Port
            $_smtpClient.EnableSsl = $UseSsl

            if ($null -ne $Credential) {
                # In PowerShell 2.0, assigning the results of GetNetworkCredential() to the SMTP client sometimes fails (with gmail, in testing), but
                # building a new NetworkCredential object containing only the UserName and Password works okay.

                $_tempCred = $Credential.GetNetworkCredential()
                $_smtpClient.Credentials = New-Object Net.NetworkCredential($Credential.UserName, $_tempCred.Password)
            }
            else {
                $_smtpClient.UseDefaultCredentials = $true
            }

            $_message = New-Object Net.Mail.MailMessage
    
            $_message.From = $From
            $_message.Subject = $Subject
        
            if ($BodyAsHtml) {
                $_bodyPart = [Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/html')
            }
            else {
                $_bodyPart = [Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/plain')
            }   

            $_message.AlternateViews.Add($_bodyPart)

            if ($PSBoundParameters.ContainsKey('DeliveryNotificationOption')) { $_message.DeliveryNotificationOptions = $DeliveryNotificationOption }
            if ($PSBoundParameters.ContainsKey('Priority')) { $_message.Priority = $Priority }

            foreach ($_address in $To) {
                if (-not $_message.To.Contains($_address)) { $_message.To.Add($_address) }
            }

            if ($null -ne $Cc) {
                foreach ($_address in $Cc) {
                    if (-not $_message.CC.Contains($_address)) { $_message.CC.Add($_address) }
                }
            }

            if ($null -ne $Bcc) {
                foreach ($_address in $Bcc) {
                    if (-not $_message.Bcc.Contains($_address)) { $_message.Bcc.Add($_address) }
                }
            }
        }
        catch {
            $_message.Dispose()
            throw
        }

        if ($PSBoundParameters.ContainsKey('InlineAttachments')) {
            foreach ($_entry in $InlineAttachments.GetEnumerator()) {
                $_file = $_entry.Value.ToString()
            
                if ([string]::IsNullOrEmpty($_file)) {
                    $_message.Dispose()
                    throw "Send-MailMessage: Values in the InlineAttachments table cannot be null."
                }

                try {
                    $_contentType = FileNameToContentType -FileName $_file
                    $_attachment = New-Object Net.Mail.LinkedResource($_file, $_contentType)
                    $_attachment.ContentId = $_entry.Key

                    $_bodyPart.LinkedResources.Add($_attachment)
                }
                catch {
                    $_message.Dispose()
                    throw
                }
            }
        }
    }

    process {
        if ($null -ne $Attachments) {
            foreach ($_file in $Attachments) {
                try {
                    $_contentType = FileNameToContentType -FileName $_file
                    $_message.Attachments.Add((New-Object Net.Mail.Attachment($_file, $_contentType)))
                }
                catch {
                    $_message.Dispose()
                    throw
                }
            }
        }
    }

    end {
        try {
            $_smtpClient.Send($_message)
        }
        catch {
            throw
        }
        finally {
            $_message.Dispose()
        }
    }
}

Param(
    [Parameter()]
    [boolean]$WhatIf = $False,
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter()]
    [ValidateRange(1, 14)] 
    [int32]$DayCount = 1
)

$days = $DayCount
if ($DayCount -gt 0) {
    $days = $DayCount * -1
}

$connectionName = "AzureRunAsConnection"

# Read parameters from Azure Automation VARIABLES
# https://docs.microsoft.com/en-us/azure/automation/automation-variables
# It's required you set them up, before you run the script

# the subscription ID of the Azure subscription 
$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"

# the template URL of the HTML Template used for the mail
$TemplateUrl = Get-AutomationVariable -Name "TemplateUrl"
$TemplateHeaderGraphicUrl = Get-AutomationVariable -Name "TemplateHeaderGraphicUrl"

# ignore some resource groups (a REGEX - e.g. "(Default-|AzureFunctions|Api-Default-).*")
$RGNamesIgnoreRegex = Get-AutomationVariable -Name "RG_NamesIgnore"

# Credentials for sending the mail - name should be Office365
# https://docs.microsoft.com/en-us/azure/automation/automation-credentials
$mailCreds = Get-AutomationPSCredential -Name 'Office365'

# The mail server
$mailServer = "smtp.office365.com";

# Single Domain that users are in
$userdomain = "@microsoft.com";

# deletion date (just a warning in the mail and another TAG, no real delete here)
# 1 month in the future
$deleteDate = (Get-Date).AddMonths(1).ToString("MM\/dd\/yy")

try {
    # Get the connection "AzureRunAsConnection"
    $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         
    
    Write-Verbose "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint | Out-Null
    
    Set-AzureRmContext -SubscriptionId $SubscriptionId | Out-Null
}
catch {
    if (!$servicePrincipalConnection) {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    }
    else {
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

if ($WhatIf) {
    Write-Warning "Running in WhatIf mode - no changes will be made."
}

$allRGs = (Get-AzureRmResourceGroup).ResourceGroupName

Write-Warning "Found $($allRGs.Length) total RGs"

$aliasedRGs = (Find-AzureRmResourceGroup -Tag @{ alias = $null }).Name

Write-Warning "Found $($aliasedRGs.Length) aliased RGs"

$notAliasedRGs = $allRGs | ? {-not ($aliasedRGs -contains $_)}

Write-Warning "Found $($notAliasedRGs.Length) un-tagged RGs"

$result = New-Object System.Collections.ArrayList

foreach ($rg in $notAliasedRGs) {
    if ($rg -match $RGNamesIgnoreRegex) {
        Write-Warning "Ignoring Resource Group $rg"
        continue
    }

    $p = 100 / ($notAliasedRGs.Length - 1 ) * $notAliasedRGs.IndexOf($rg)
    Write-Progress -Activity "Searching Resource Group Logs for last $days days..." -PercentComplete $p `
        -CurrentOperation "$p% complete" `
        -Status "Resource Group $rg"

    $callers = Get-AzureRmLog -ResourceGroup $rg -DetailedOutput `
            -StartTime (Get-Date).AddDays($days) `
            -EndTime (Get-Date)`
            | Where-Object Caller -like "*@*" `
            | Where-Object { $_.Caller -and ($_.Caller -ne "System") } `
            | Where-Object { $_.OperationName.Value -ne "Microsoft.Storage/storageAccounts/listKeys/action" }`
            | Where-Object { $_.Properties.Content -and ($_.Properties.Content["requestbody"] -ne "{""tags"":{}}" ) } `
            | Sort-Object -Property Caller -Unique `
            | Select-Object Caller

    if ($callers) {
        $alias = $callers[0].Caller -replace $userdomain, ""
				
        Write-Warning "Tagging Resource Group $rg for alias $alias"
        
        if (-not $WhatIf) {
            Set-AzureRmResourceGroup -Name $rg -Tag @{ alias = $alias; deleteAfter = $deleteDate} | Out-Null
        }
        $result.Add((New-Object PSObject -Property @{Name = $rg; Alias = $alias})) | Out-Null
        
    }
    else {
        Write-Warning "No activity found for Resource Group $rg"
    }
}

Write-Progress -Activity "Searching Resource Group Logs..." -Completed -Status "Done"

if ($result.Count -gt 0) {
    $rgString = ($result | ForEach-Object { "<tr><td>$($_.Name)</td><td>$($_.Alias)</td></tr>" })

    $toAffected = ($result | ForEach-Object { "<$($_.Alias)$($userdomain)>" }) -join ";"

    $template = Invoke-WebRequest -Uri $TemplateUrl -UseBasicParsing

    $body = $template -replace "_TABLE_", $rgString -replace "_DATE_", $deleteDate

    $subject = "$($result.Count) new resource groups automatically tagged";

    $tocomb = "$To;$toAffected"
    $toArray = $tocomb.Split(";")

    Write-Warning "Sending Mail to $tocomb"

    Invoke-WebRequest -UseBasicParsing $TemplateHeaderGraphicUrl -OutFile C:\template.png

    # This is specifically for having a tagging.png header graphic embedded to the mail
    # of course you can remove the line -InlineAttachments
    Send-MailMessageEx `
        -Body $body `
        -Subject $subject `
        -Credential $mailCreds `
        -SmtpServer $mailServer `
        -Port 587 `
        -BodyAsHtml `
        -UseSSL `
        -InlineAttachments @{ "tagging.png" = "C:\template.png" } `
        -From $mailCreds.UserName `
        -To $toArray `
        -Priority "Low"
}
else {
    Write-Warning "No Email sent - 0 Resource Groups tagged"
}

$result