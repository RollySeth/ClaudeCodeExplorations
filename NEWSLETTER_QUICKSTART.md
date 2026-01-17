# Business Weekly Newsletter

A professional, self-contained HTML email newsletter template for Microsoft Outlook. Works great with Claude Code for AI-assisted customization.

**Self-contained** — No external images or dependencies required. Just run the script!

---

## Features

- **100% Self-Contained** — Built-in styled banner, no external images needed
- **Outlook Compatible** — Table-based HTML optimized for Outlook desktop
- **Modern Design** — Pink-orange gradient banner, clean cards, professional styling
- **4 Sections** — Product Updates, Key Metrics, User Feedback, What's Next
- **Easy Customization** — Simple placeholder system for your content
- **Optional Custom Banner** — Override with your own image if desired
- **Preview Mode** — Review before sending
- **Claude Code Ready** — AI-assisted content generation and customization

---

## Requirements

- Windows 10/11
- Microsoft Outlook (desktop application)
- PowerShell 5.1 or higher

---

## Quick Start (5 Minutes)

### Step 1: Download the Script

Save `business_weekly_newsletter_public.ps1` to your computer.

### Step 2: Preview the Template

```powershell
# Open PowerShell and navigate to the script location
cd "C:\path\to\script"

# Run in preview mode (opens in Outlook)
.\business_weekly_newsletter_public.ps1 -DisplayOnly
```

This opens the template in Outlook with placeholder values so you can see the layout.

### Step 3: Customize Your Content

Open the script in any text editor (VS Code, Notepad++, etc.) and edit the **Configuration Section** (lines 50-150):

```powershell
# BRANDING (line ~55)
$config = @{
    newsletterTitle    = "Acme Weekly Update"        # Your newsletter name
    newsletterSubtitle = "Product News & Insights"   # Tagline
    senderName         = "Product Team"              # From name
    companyName        = "Acme Corp"                 # Company name
    primaryColor       = "#667eea"                   # Main accent color
}

# PRODUCT UPDATES (line ~80)
$productUpdates = @(
    @{
        title       = "New Dashboard Released"
        status      = "Shipped"
        description = "Track all your metrics in one place with our redesigned dashboard."
        metric      = "500+ active users"
        icon        = "&#128640;"  # Rocket emoji
    },
    # ... add more updates
)

# KEY METRICS (line ~100)
$keyMetrics = @(
    @{
        label      = "Active Users"
        value      = "12.5K"
        trend      = "&#8593; 23%"  # Up arrow
        trendColor = "#10b981"      # Green
    },
    # ... add more metrics
)

# USER FEEDBACK (line ~130)
$userFeedback = @(
    "This update saved our team 10 hours per week!",
    "Finally, a feature we've been waiting for.",
    "The new UI is so much cleaner."
)

# UPCOMING FEATURES (line ~145)
$upcomingFeatures = @(
    @{
        title       = "AI-Powered Analytics"
        timeline    = "Q2 2026"
        description = "Get intelligent insights automatically generated from your data."
    },
    # ... add more
)
```

### Step 4: Customize the Banner

The script includes a **built-in styled HTML banner** — no external image needed!

**Edit banner settings in the config section (line ~89):**

```powershell
# Banner Configuration (Built-in HTML banner - no image needed!)
bannerHeadline     = "This Week in Product"           # Large text
bannerTagline      = "Innovation &bull; Updates &bull; Insights"  # Subtext
bannerIcon         = "&#128640;"                      # Emoji icon
bannerGradientStart = "#667eea"                       # Left gradient color
bannerGradientEnd   = "#764ba2"                       # Right gradient color
```

**Banner Icon Options:**
| Icon | Code | Description |
|------|------|-------------|
| 🚀 | `&#128640;` | Rocket (default) |
| ⭐ | `&#11088;` | Star |
| ✨ | `&#10024;` | Sparkles |
| 📊 | `&#128202;` | Chart |
| 💡 | `&#128161;` | Light bulb |
| 🎯 | `&#127919;` | Target |

**Optional: Use a Custom Image Instead**

If you prefer your own banner image:
- **Recommended size:** 1200 x 400 pixels
- **Format:** PNG or JPG

```powershell
.\business_weekly_newsletter_public.ps1 -BannerPath "C:\Images\my_banner.png" -DisplayOnly
```

### Step 5: Send the Newsletter

Once you're happy with the preview:

```powershell
# Send to a single recipient
.\business_weekly_newsletter_public.ps1 -To "team@company.com" -BannerPath "C:\banner.png" -DisplayOnly:$false

# Or open in Outlook and add recipients manually
.\business_weekly_newsletter_public.ps1 -BannerPath "C:\banner.png" -DisplayOnly
# Then add recipients in Outlook and click Send
```

---

## Customization Reference

### Status Options

| Status | Color | Use For |
|--------|-------|---------|
| `Shipped` | Green | Completed features |
| `In Progress` | Amber | Features in development |
| `Planning` | Purple | Future roadmap items |

### Emoji/Icon Codes

Copy these into the `icon` fields:

| Icon | Code | Description |
|------|------|-------------|
| 🚀 | `&#128640;` | Rocket (launches, releases) |
| ⚡ | `&#9889;` | Lightning (performance) |
| ✨ | `&#10024;` | Sparkles (new features) |
| 📊 | `&#128202;` | Chart (metrics, data) |
| 💬 | `&#128172;` | Speech (feedback, communication) |
| 🔮 | `&#128302;` | Crystal ball (future, roadmap) |
| ✔ | `&#10004;` | Check mark (completed) |
| ⭐ | `&#11088;` | Star (highlights) |
| 🔥 | `&#128293;` | Fire (trending, hot) |
| 🎯 | `&#127919;` | Target (goals) |
| 🏆 | `&#127942;` | Trophy (wins, achievements) |
| 💡 | `&#128161;` | Light bulb (ideas) |

### Trend Arrows

| Arrow | Code | Use For |
|-------|------|---------|
| ↑ | `&#8593;` | Positive increase |
| ↓ | `&#8595;` | Decrease (good for costs/bugs) |
| → | `&#8594;` | Stable/unchanged |
| 📈 | `&#128200;` | Trending up |
| 📉 | `&#128201;` | Trending down |

### Color Codes

| Purpose | Default | Alternatives |
|---------|---------|--------------|
| Primary (accent) | `#667eea` | `#3b82f6` (blue), `#8b5cf6` (violet) |
| Success (positive) | `#10b981` | `#22c55e` (green) |
| Warning | `#f59e0b` | `#f97316` (orange) |
| Text (dark) | `#1f2937` | `#111827` (darker) |
| Text (light) | `#6b7280` | `#9ca3af` (lighter) |

---

## Examples

### Example 1: Minimal Newsletter

```powershell
$productUpdates = @(
    @{
        title       = "Version 2.0 Released"
        status      = "Shipped"
        description = "Major update with new features and performance improvements."
        metric      = "Available now"
        icon        = "&#128640;"
    }
)

$keyMetrics = @(
    @{ label = "Users"; value = "10K"; trend = "&#8593; 15%"; trendColor = "#10b981" }
)

$userFeedback = @(
    "Great update! Love the new features."
)

$upcomingFeatures = @(
    @{
        title       = "Mobile App"
        timeline    = "Coming Soon"
        description = "Access everything on the go."
    }
)
```

### Example 2: Full Newsletter

```powershell
$productUpdates = @(
    @{
        title       = "AI-Powered Search"
        status      = "Shipped"
        description = "Find anything instantly with our new semantic search engine."
        metric      = "3x faster results"
        icon        = "&#128640;"
    },
    @{
        title       = "Dark Mode"
        status      = "In Progress"
        description = "Easier on the eyes for late-night work sessions."
        metric      = "Beta testing"
        icon        = "&#9889;"
    },
    @{
        title       = "API v3"
        status      = "Planning"
        description = "More endpoints, better documentation, faster responses."
        metric      = "Q2 Roadmap"
        icon        = "&#10024;"
    }
)

$keyMetrics = @(
    @{ label = "Monthly Users"; value = "125K"; trend = "&#8593; 32%"; trendColor = "#10b981" },
    @{ label = "NPS Score"; value = "72"; trend = "&#8593; 8 pts"; trendColor = "#10b981" },
    @{ label = "Uptime"; value = "99.9%"; trend = "&#8594; Stable"; trendColor = "#667eea" },
    @{ label = "Avg Response"; value = "45ms"; trend = "&#8595; 12ms"; trendColor = "#10b981" }
)

$userFeedback = @(
    "The new search is incredible. Found what I needed in seconds!",
    "Finally, an API that's actually well-documented.",
    "Your support team resolved my issue in under an hour. Impressed!",
    "Dark mode can't come soon enough. Take my money!",
    "Been using this for 2 years. It keeps getting better."
)

$upcomingFeatures = @(
    @{
        title       = "Mobile App"
        timeline    = "March 2026"
        description = "Full-featured iOS and Android apps with offline support."
    },
    @{
        title       = "Team Workspaces"
        timeline    = "Q2 2026"
        description = "Collaborate with your team in shared project spaces."
    },
    @{
        title       = "Integrations Hub"
        timeline    = "Q3 2026"
        description = "Connect with Slack, Teams, Notion, and 50+ other tools."
    }
)
```

---

## Troubleshooting

### "Cannot create Outlook object"

**Cause:** Outlook isn't installed or not configured.

**Fix:** Install Microsoft Outlook desktop app (not just web) and set up an email account.

### "Custom banner not found" (warning)

**Cause:** You specified `-BannerPath` but the file doesn't exist.

**Fix:** Either use the full path (`C:\Users\YourName\Pictures\banner.png`) or remove the `-BannerPath` parameter to use the built-in HTML banner instead.

### Email looks different in other clients

**Note:** This template is optimized for Outlook. Gmail, Apple Mail, etc. may render slightly differently. The table-based layout ensures maximum compatibility.

### PowerShell execution policy error

**Fix:** Run this command first:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## File Structure

```
email_automation/
├── business_weekly_newsletter_public.ps1   # Main script
├── NEWSLETTER_QUICKSTART.md                # This guide
└── banners/                                # (Optional) Store your banners here
    ├── weekly_banner.png
    └── monthly_banner.png
```

---

## Tips & Best Practices

1. **Keep it scannable** — Use short descriptions (1-2 sentences max)
2. **Lead with impact** — Put your biggest news first
3. **Use real numbers** — Metrics build credibility
4. **Include user quotes** — Social proof is powerful
5. **Preview before sending** — Always use `-DisplayOnly` first
6. **Test with yourself** — Send to your own email first
7. **Consistent schedule** — Weekly newsletters work best on the same day/time

---

## Using with Claude Code

This template works seamlessly with [Claude Code](https://claude.ai/claude-code) for AI-assisted newsletter creation.

### Generate Newsletter Content

Ask Claude to help populate your newsletter:

```
Fill in my newsletter template with:
- 3 product updates we shipped this week (dashboard redesign, API improvements, mobile app beta)
- Key metrics: 15K users (+23%), 99.9% uptime, 45ms response time
- 5 pieces of positive user feedback
- 2 upcoming features for Q2
```

### Customize Styling

```
Change the newsletter primary color to blue (#3b82f6) and update
the banner headline to "Engineering Weekly"
```

### Preview in Outlook

```
Run the newsletter script in preview mode so I can see how it looks
```

### Iterate Quickly

```
The metrics section looks too crowded. Can you reduce it to 3 metrics
and make the cards wider?
```

### Automate with Scheduling

```
Help me set up a Windows scheduled task to send this newsletter
every Monday at 9am to team@company.com
```

### Example Claude Code Workflow

1. **Start a session:** Open Claude Code in your project folder
2. **Generate content:** "Create newsletter content for this week's product updates"
3. **Preview:** "Run the script with -DisplayOnly to preview"
4. **Iterate:** "Make the user feedback section shorter"
5. **Send:** "Remove -DisplayOnly and send to the team"

---

## Advanced: Scheduling Weekly Sends

Use Windows Task Scheduler to automate weekly sends:

```powershell
# Create a scheduled task (run as Administrator)
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\Scripts\business_weekly_newsletter_public.ps1`" -To `"team@company.com`" -BannerPath `"C:\Images\banner.png`" -DisplayOnly:`$false"

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 9am

Register-ScheduledTask -TaskName "Weekly Newsletter" -Action $action -Trigger $trigger -Description "Send weekly product newsletter"
```

---

## License

MIT License — Free to use, modify, and distribute.

---

## Support

For issues or feature requests, open an issue on GitHub or contact the maintainer.

---

**Happy sending!** 📬
