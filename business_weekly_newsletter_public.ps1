<#
.SYNOPSIS
    Business Weekly Newsletter - HTML Email Template Generator
    A professional, modern newsletter template for Outlook with table-based layout

.DESCRIPTION
    This PowerShell script generates a beautifully designed HTML email newsletter
    that is fully compatible with Microsoft Outlook and other email clients.

    Features:
    - Table-based HTML layout (Outlook compatible)
    - Modern glassmorphism/neomorphism design
    - Embedded banner image support
    - 4 customizable sections: Updates, Metrics, Feedback, Upcoming
    - Placeholder system for easy customization
    - Display-only preview mode

.PARAMETER To
    Email address to send the newsletter to. Leave empty for display-only mode.

.PARAMETER BannerPath
    (Optional) Path to a custom banner image file (PNG, JPG). If not provided,
    uses a built-in styled HTML banner that works without external files.

.PARAMETER DisplayOnly
    When set, opens the email in Outlook for preview instead of sending.

.EXAMPLE
    # Preview the newsletter in Outlook
    .\business_weekly_newsletter_public.ps1 -DisplayOnly

.EXAMPLE
    # Send the newsletter to a recipient
    .\business_weekly_newsletter_public.ps1 -To "team@example.com" -DisplayOnly:$false

.EXAMPLE
    # Use custom banner
    .\business_weekly_newsletter_public.ps1 -BannerPath "C:\Images\my_banner.png" -DisplayOnly

.NOTES
    Author: Open Source Community
    Version: 1.0.0
    License: MIT

    Requirements:
    - Windows with Microsoft Outlook installed
    - PowerShell 5.1 or higher

.LINK
    https://github.com/your-repo/newsletter-template
#>

param(
    [string]$To = "",
    [string]$BannerPath = "",
    [switch]$DisplayOnly = $true
)

$ErrorActionPreference = "Stop"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Business Weekly Newsletter Template" -ForegroundColor Cyan
Write-Host "  Open Source Edition v1.0" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================
# CONFIGURATION SECTION - CUSTOMIZE THESE
# ============================================

# Newsletter Branding (Edit these for your organization)
$config = @{
    # Basic Info
    newsletterTitle    = "Weekly Product Newsletter"
    newsletterSubtitle = "Your Product Intelligence Briefing"
    senderName         = "Product Team"
    companyName        = "Your Company"

    # Colors (Hex codes)
    primaryColor       = "#667eea"    # Main accent color (purple-blue)
    secondaryColor     = "#764ba2"    # Secondary accent
    successColor       = "#10b981"    # Positive trends (green)
    warningColor       = "#f59e0b"    # Warning/In Progress (amber)
    textPrimary        = "#1f2937"    # Dark text
    textSecondary      = "#6b7280"    # Light text
    background         = "#f5f7fa"    # Page background
    cardBackground     = "#ffffff"    # Card background

    # Banner Configuration (Built-in HTML banner - no image needed!)
    bannerHeadline     = "This Week in Product"           # Large text on banner
    bannerTagline      = "Innovation &bull; Updates &bull; Insights"  # Smaller text below
    bannerIcon         = "&#10024;"                       # Sparkles emoji (&#10024;), Star(&#11088;), Bulb(&#128161;)
    bannerGradientStart = "#ec4899"                       # Pink gradient (left)
    bannerGradientEnd   = "#f97316"                       # Orange gradient (right)

    # Optional: Footer Links (set to empty string to hide)
    websiteUrl         = ""           # e.g., "https://yourcompany.com"
    unsubscribeUrl     = ""           # e.g., "https://yourcompany.com/unsubscribe"
}

$today = Get-Date -Format "MMMM dd, yyyy"

# ============================================
# SECTION 1: PRODUCT UPDATES (3 cards)
# ============================================
# Replace placeholders with your actual content

$productUpdates = @(
    @{
        title       = "{{FEATURE_1_TITLE}}"           # e.g., "New Dashboard Released"
        status      = "{{STATUS}}"                     # Options: "Shipped", "In Progress", "Planning"
        description = "{{FEATURE_1_DESCRIPTION}}"      # 1-2 sentence description
        metric      = "{{METRIC}}"                     # e.g., "10K users", "85% satisfaction"
        icon        = "&#128640;"                      # See emoji codes below
    },
    @{
        title       = "{{FEATURE_2_TITLE}}"
        status      = "{{STATUS}}"
        description = "{{FEATURE_2_DESCRIPTION}}"
        metric      = "{{METRIC}}"
        icon        = "&#9889;"
    },
    @{
        title       = "{{FEATURE_3_TITLE}}"
        status      = "{{STATUS}}"
        description = "{{FEATURE_3_DESCRIPTION}}"
        metric      = "{{METRIC}}"
        icon        = "&#10024;"
    }
)

# ============================================
# SECTION 2: KEY METRICS (4 metric cards)
# ============================================

$keyMetrics = @(
    @{
        label      = "Active Users"
        value      = "{{USER_COUNT}}"                  # e.g., "1.2M", "50K"
        trend      = "{{TREND}}"                       # e.g., "&#8593; 15%" (up arrow)
        trendColor = "#10b981"                         # Green for positive
    },
    @{
        label      = "Satisfaction Score"
        value      = "{{SCORE}}"                       # e.g., "72", "4.5/5"
        trend      = "{{TREND}}"                       # e.g., "&#8593; 8 points"
        trendColor = "#10b981"
    },
    @{
        label      = "Feature Adoption"
        value      = "{{ADOPTION_RATE}}"               # e.g., "68%"
        trend      = "{{TREND}}"
        trendColor = "#667eea"                         # Primary color for neutral
    },
    @{
        label      = "Response Time"
        value      = "{{RESPONSE_TIME}}"               # e.g., "2.3 days", "4 hours"
        trend      = "{{TREND}}"                       # e.g., "&#8595; 0.5 days" (down = good)
        trendColor = "#10b981"
    }
)

# ============================================
# SECTION 3: USER FEEDBACK HIGHLIGHTS (up to 5)
# ============================================

$userFeedback = @(
    "{{USER_FEEDBACK_1}}",                             # e.g., "The new feature saved us hours!"
    "{{USER_FEEDBACK_2}}",
    "{{USER_FEEDBACK_3}}",
    "{{USER_FEEDBACK_4}}",
    "{{USER_FEEDBACK_5}}"
)

# ============================================
# SECTION 4: WHAT'S NEXT (3 items)
# ============================================

$upcomingFeatures = @(
    @{
        title       = "{{UPCOMING_1_TITLE}}"           # e.g., "AI-Powered Analytics"
        timeline    = "{{TIMELINE}}"                   # e.g., "Q2 2026", "Next Sprint", "March"
        description = "{{UPCOMING_1_DESCRIPTION}}"
    },
    @{
        title       = "{{UPCOMING_2_TITLE}}"
        timeline    = "{{TIMELINE}}"
        description = "{{UPCOMING_2_DESCRIPTION}}"
    },
    @{
        title       = "{{UPCOMING_3_TITLE}}"
        timeline    = "{{TIMELINE}}"
        description = "{{UPCOMING_3_DESCRIPTION}}"
    }
)

# ============================================
# EMOJI/ICON REFERENCE (HTML Entity Codes)
# ============================================
# Copy these into icon fields above:
#
# Rocket:        &#128640;
# Lightning:     &#9889;
# Sparkles:      &#10024;
# Chart:         &#128202;
# Speech:        &#128172;
# Crystal Ball:  &#128302;
# Check Mark:    &#10004;
# Star:          &#11088;
# Fire:          &#128293;
# Target:        &#127919;
# Trophy:        &#127942;
# Light Bulb:    &#128161;
#
# Arrows:
# Up Arrow:      &#8593;
# Down Arrow:    &#8595;
# Right Arrow:   &#8594;
# Trending Up:   &#128200;
# Trending Down: &#128201;

# ============================================
# HTML GENERATION ENGINE
# (Modify only if you need layout changes)
# ============================================

Write-Host "[1/3] Generating newsletter HTML..." -ForegroundColor Yellow

# Build Product Updates HTML
$productUpdatesHtml = ""
foreach ($update in $productUpdates) {
    $statusColor = switch ($update.status) {
        "Shipped"     { $config.successColor }
        "In Progress" { $config.warningColor }
        "Planning"    { $config.primaryColor }
        default       { $config.textSecondary }
    }

    $productUpdatesHtml += @"
<tr>
    <td style="padding: 16px 0;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: $($config.cardBackground); border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
            <tr>
                <td style="padding: 20px;">
                    <table cellpadding="0" cellspacing="0" border="0" width="100%">
                        <tr>
                            <td width="40" style="font-size: 28px; vertical-align: top;">$($update.icon)</td>
                            <td style="vertical-align: top;">
                                <div style="color: $($config.textPrimary); font-size: 18px; font-weight: 700; margin-bottom: 8px;">$($update.title)</div>
                                <div style="margin-bottom: 12px;">
                                    <span style="display: inline-block; background: $statusColor; color: white; font-size: 11px; font-weight: 600; padding: 4px 12px; border-radius: 12px; text-transform: uppercase; letter-spacing: 0.5px;">$($update.status)</span>
                                </div>
                                <div style="color: $($config.textSecondary); font-size: 14px; line-height: 1.6; margin-bottom: 12px;">$($update.description)</div>
                                <div style="color: $($config.primaryColor); font-size: 15px; font-weight: 700;">&#128202; $($update.metric)</div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </td>
</tr>
"@
}

# Build Key Metrics HTML
$metricsHtml = @"
<tr>
    <td style="padding: 0;">
        <table cellpadding="8" cellspacing="0" border="0" width="100%">
            <tr>
"@

foreach ($metric in $keyMetrics) {
    $metricsHtml += @"
                <td width="25%" style="vertical-align: top;">
                    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: rgba(255, 255, 255, 0.95); border-radius: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                        <tr>
                            <td style="padding: 20px; text-align: center;">
                                <div style="color: $($config.textSecondary); font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">$($metric.label)</div>
                                <div style="color: $($config.textPrimary); font-size: 32px; font-weight: 700; margin-bottom: 8px;">$($metric.value)</div>
                                <div style="color: $($metric.trendColor); font-size: 13px; font-weight: 600;">$($metric.trend)</div>
                            </td>
                        </tr>
                    </table>
                </td>
"@
}

$metricsHtml += @"
            </tr>
        </table>
    </td>
</tr>
"@

# Build User Feedback HTML
$feedbackHtml = ""
foreach ($feedback in $userFeedback) {
    if (-not [string]::IsNullOrWhiteSpace($feedback) -and $feedback -notlike "{{*}}") {
        $feedbackHtml += @"
<tr>
    <td style="padding: 8px 0;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: $($config.cardBackground); border-left: 4px solid $($config.primaryColor); border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.1);">
            <tr>
                <td style="padding: 16px;">
                    <div style="color: $($config.textPrimary); font-size: 14px; line-height: 1.6; font-style: italic;">&ldquo;$feedback&rdquo;</div>
                </td>
            </tr>
        </table>
    </td>
</tr>
"@
    }
}

# Show placeholder note if feedback is empty
if ($feedbackHtml -eq "") {
    $feedbackHtml = @"
<tr>
    <td style="padding: 8px 0;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: #fef3c7; border-left: 4px solid $($config.warningColor); border-radius: 8px;">
            <tr>
                <td style="padding: 16px;">
                    <div style="color: #92400e; font-size: 13px;"><strong>&#128221; Customization Area:</strong> Replace {{USER_FEEDBACK_X}} placeholders with actual user quotes</div>
                </td>
            </tr>
        </table>
    </td>
</tr>
"@
}

# Build Upcoming Features HTML
$upcomingHtml = ""
foreach ($upcoming in $upcomingFeatures) {
    $upcomingHtml += @"
<tr>
    <td style="padding: 12px 0;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: linear-gradient(135deg, rgba(102, 126, 234, 0.08) 0%, rgba(118, 75, 162, 0.08) 100%); border-left: 4px solid $($config.primaryColor); border-radius: 12px;">
            <tr>
                <td style="padding: 20px;">
                    <table cellpadding="0" cellspacing="0" border="0" width="100%">
                        <tr>
                            <td style="vertical-align: middle;">
                                <div style="color: $($config.primaryColor); font-size: 18px; font-weight: 700;">$($upcoming.title)</div>
                            </td>
                            <td width="120" style="text-align: right; vertical-align: middle;">
                                <span style="display: inline-block; background: $($config.primaryColor); color: white; font-size: 11px; font-weight: 600; padding: 6px 12px; border-radius: 12px; white-space: nowrap;">$($upcoming.timeline)</span>
                            </td>
                        </tr>
                    </table>
                    <div style="color: #4b5563; font-size: 14px; line-height: 1.6; margin-top: 8px;">$($upcoming.description)</div>
                </td>
            </tr>
        </table>
    </td>
</tr>
"@
}

# Check for placeholders to show customization hint
$hasPlaceholders = $false
foreach ($update in $productUpdates) {
    if ($update.title -like "{{*}}" -or $update.description -like "{{*}}") {
        $hasPlaceholders = $true
        break
    }
}

$customizationNote = ""
if ($hasPlaceholders) {
    $customizationNote = @"
<tr>
    <td style="padding: 20px 40px;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: #fef3c7; border-left: 4px solid $($config.warningColor); border-radius: 8px;">
            <tr>
                <td style="padding: 16px;">
                    <div style="color: #92400e; font-size: 13px; line-height: 1.6;">
                        <strong>&#128221; Template Preview Mode:</strong> This newsletter contains placeholder values. Edit the script's configuration section (lines 50-150) to replace all {{PLACEHOLDER}} values with your actual data before sending.
                    </div>
                </td>
            </tr>
        </table>
    </td>
</tr>
"@
}

# Build optional footer links
$footerLinks = ""
if ($config.websiteUrl -or $config.unsubscribeUrl) {
    $links = @()
    if ($config.websiteUrl) { $links += "<a href=`"$($config.websiteUrl)`" style=`"color: $($config.primaryColor); text-decoration: none;`">Visit Website</a>" }
    if ($config.unsubscribeUrl) { $links += "<a href=`"$($config.unsubscribeUrl)`" style=`"color: #9ca3af; text-decoration: none;`">Unsubscribe</a>" }
    $footerLinks = "<div style=`"margin-top: 12px;`">$($links -join ' &bull; ')</div>"
}

# ============================================
# MAIN HTML TEMPLATE (Table-Based for Outlook)
# ============================================

$htmlBody = @"
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$($config.newsletterTitle)</title>
</head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background-color: $($config.background);">
    <!-- Main Container Table -->
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color: $($config.background);">
        <tr>
            <td align="center" style="padding: 20px 0;">
                <!-- Content Container (800px max width) -->
                <table cellpadding="0" cellspacing="0" border="0" width="800" style="max-width: 800px; background-color: $($config.cardBackground); border-radius: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">

                    <!-- Header -->
                    <tr>
                        <td style="padding: 40px 40px 20px 40px; text-align: center;">
                            <h1 style="margin: 0; font-size: 32px; font-weight: 700; color: $($config.textPrimary); letter-spacing: -0.5px;">$($config.newsletterTitle)</h1>
                            <div style="margin-top: 8px; font-size: 14px; color: $($config.textSecondary);">$($config.newsletterSubtitle) &bull; $today</div>
                        </td>
                    </tr>

                    <!-- Styled Banner -->
                    <tr>
                        <td style="padding: 20px 40px;">
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background: linear-gradient(90deg, #ec4899 0%, #f97316 100%); background-color: #ec4899; border-radius: 12px;">
                                <tr>
                                    <td align="center" valign="middle" style="padding: 32px 40px; text-align: center;">
                                        <!-- Icon -->
                                        <div style="font-size: 40px; margin-bottom: 12px;">$($config.bannerIcon)</div>
                                        <!-- Headline -->
                                        <div style="color: #ffffff; font-size: 32px; font-weight: 700; letter-spacing: -0.5px; margin-bottom: 8px;">$($config.bannerHeadline)</div>
                                        <!-- Tagline -->
                                        <div style="color: rgba(255,255,255,0.9); font-size: 14px; font-weight: 500; letter-spacing: 2px; text-transform: uppercase;">$($config.bannerTagline)</div>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    $customizationNote

                    <!-- SECTION 1: PRODUCT UPDATES -->
                    <tr>
                        <td style="padding: 32px 40px 0 40px;">
                            <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                <tr>
                                    <td style="color: $($config.primaryColor); font-size: 22px; font-weight: 700; padding: 16px 0; border-left: 4px solid $($config.primaryColor); padding-left: 16px;">
                                        &#128640; Product Updates
                                    </td>
                                </tr>
                                $productUpdatesHtml
                            </table>
                        </td>
                    </tr>

                    <!-- SECTION 2: KEY METRICS -->
                    <tr>
                        <td style="padding: 32px 40px 0 40px;">
                            <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                <tr>
                                    <td style="color: $($config.primaryColor); font-size: 22px; font-weight: 700; padding: 16px 0; border-left: 4px solid $($config.primaryColor); padding-left: 16px;">
                                        &#128202; Key Metrics
                                    </td>
                                </tr>
                                $metricsHtml
                            </table>
                        </td>
                    </tr>

                    <!-- SECTION 3: USER FEEDBACK -->
                    <tr>
                        <td style="padding: 32px 40px 0 40px;">
                            <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                <tr>
                                    <td style="color: $($config.primaryColor); font-size: 22px; font-weight: 700; padding: 16px 0; border-left: 4px solid $($config.primaryColor); padding-left: 16px;">
                                        &#128172; User Feedback Highlights
                                    </td>
                                </tr>
                                $feedbackHtml
                            </table>
                        </td>
                    </tr>

                    <!-- SECTION 4: WHAT'S NEXT -->
                    <tr>
                        <td style="padding: 32px 40px 0 40px;">
                            <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                <tr>
                                    <td style="color: $($config.primaryColor); font-size: 22px; font-weight: 700; padding: 16px 0; border-left: 4px solid $($config.primaryColor); padding-left: 16px;">
                                        &#128302; What's Next
                                    </td>
                                </tr>
                                $upcomingHtml
                            </table>
                        </td>
                    </tr>

                    <!-- Footer -->
                    <tr>
                        <td style="padding: 40px; text-align: center; border-top: 1px solid #e5e7eb; margin-top: 32px;">
                            <div style="color: $($config.textPrimary); font-size: 14px; font-weight: 600; margin-bottom: 4px;">$($config.senderName)</div>
                            <div style="color: #9ca3af; font-size: 12px;">$($config.companyName)</div>
                            <div style="color: #9ca3af; font-size: 11px; margin-top: 8px;">Generated on $today</div>
                            $footerLinks
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"@

Write-Host "  HTML generated successfully" -ForegroundColor Green

# ============================================
# EMAIL SENDING VIA OUTLOOK
# ============================================

if ($DisplayOnly) {
    Write-Host "[2/3] Opening preview in Outlook..." -ForegroundColor Yellow
} else {
    Write-Host "[2/3] Preparing to send email..." -ForegroundColor Yellow
}

try {
    # Create Outlook COM object
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)  # 0 = olMailItem

    # Set recipient if provided
    if ($To) {
        $mail.To = $To
    }

    # Set subject and body
    $mail.Subject = "$($config.newsletterTitle) - $today"
    $mail.HTMLBody = $htmlBody

    # Banner handling: Built-in HTML banner is always included
    # Optional: Override with custom image if BannerPath is provided
    if ($BannerPath -and (Test-Path $BannerPath)) {
        $attachment = $mail.Attachments.Add($BannerPath)
        $attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "banner")
        Write-Host "  Custom banner image embedded: $BannerPath" -ForegroundColor Gray
    } elseif ($BannerPath) {
        Write-Host "  Warning: Custom banner not found at $BannerPath (using built-in banner)" -ForegroundColor Yellow
    } else {
        Write-Host "  Using built-in styled HTML banner (no external image needed)" -ForegroundColor Gray
    }

    # Display or Send
    if ($DisplayOnly) {
        $mail.Display()
        Write-Host ""
        Write-Host "[3/3] Complete!" -ForegroundColor Green
        Write-Host ""
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "  Newsletter Preview Opened in Outlook" -ForegroundColor Cyan
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Next Steps:" -ForegroundColor White
        Write-Host "  1. Replace all {{PLACEHOLDER}} values in the script" -ForegroundColor Gray
        Write-Host "  2. Re-run the script to preview changes" -ForegroundColor Gray
        Write-Host "  3. Set -DisplayOnly:`$false to send" -ForegroundColor Gray
        Write-Host ""
    } else {
        if (-not $To) {
            Write-Host ""
            Write-Host "  Error: No recipient specified. Use -To parameter." -ForegroundColor Red
            $mail.Display()
            Write-Host "  Opening in Outlook for manual sending..." -ForegroundColor Yellow
        } else {
            $mail.Send()
            Write-Host ""
            Write-Host "[3/3] Complete!" -ForegroundColor Green
            Write-Host ""
            Write-Host "==================================================" -ForegroundColor Cyan
            Write-Host "  Newsletter Sent Successfully!" -ForegroundColor Cyan
            Write-Host "==================================================" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  Sent to: $To" -ForegroundColor Gray
            Write-Host "  Subject: $($config.newsletterTitle) - $today" -ForegroundColor Gray
            Write-Host ""
        }
    }

} catch {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host "  Error Occurred" -ForegroundColor Red
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  - Ensure Microsoft Outlook is installed" -ForegroundColor Gray
    Write-Host "  - Run PowerShell as Administrator if needed" -ForegroundColor Gray
    Write-Host "  - Check if Outlook is already open" -ForegroundColor Gray
    Write-Host ""
    exit 1

} finally {
    # Clean up COM object
    if ($null -ne $outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}

# ============================================
# USAGE EXAMPLES (for reference)
# ============================================
<#
QUICK START:
------------
1. Open this script in a text editor
2. Scroll to the CONFIGURATION SECTION (line ~50)
3. Edit the $config hashtable with your branding
4. Edit $productUpdates, $keyMetrics, $userFeedback, $upcomingFeatures
5. Run: .\business_weekly_newsletter_public.ps1 -DisplayOnly

CUSTOMIZATION TIPS:
-------------------
- Colors: Use hex codes (e.g., "#667eea")
- Icons: Use HTML entity codes (see EMOJI REFERENCE section)
- Status options: "Shipped", "In Progress", "Planning"
- Trends: Use &#8593; (up) or &#8595; (down) arrows

BANNER IMAGE:
-------------
- Recommended size: 1200x400 pixels
- Supported formats: PNG, JPG
- The image will be embedded in the email (not linked)

SENDING OPTIONS:
----------------
# Preview only (default)
.\business_weekly_newsletter_public.ps1

# Preview with custom banner
.\business_weekly_newsletter_public.ps1 -BannerPath "C:\banner.png"

# Send to recipient
.\business_weekly_newsletter_public.ps1 -To "team@company.com" -DisplayOnly:$false

# Send with banner
.\business_weekly_newsletter_public.ps1 -To "team@company.com" -BannerPath "C:\banner.png" -DisplayOnly:$false
#>
