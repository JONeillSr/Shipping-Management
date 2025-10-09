<#
.SYNOPSIS
    Creates starter templates for common auction vendors
    
.DESCRIPTION
    Generates pre-configured templates for frequently used auction houses
    to speed up the configuration process.
    
.EXAMPLE
    .\Create-StarterTemplates.ps1
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-07
    Version: 1.0.0
    Change Date: 
    Change Purpose: Initial Release
#>

$TemplateDirectory = ".\Templates"

# Create Templates directory if it doesn't exist
if (!(Test-Path $TemplateDirectory)) {
    New-Item -ItemType Directory -Path $TemplateDirectory -Force | Out-Null
    Write-Host "✅ Created Templates directory: $TemplateDirectory" -ForegroundColor Green
}

#region Template Definitions

# Brolyn Auctions Template
$BrolynTemplate = @{
    email_subject = "Freight Quote Request - [Location] to Ashtabula, OH - Pickup [Date]"
    auction_info = @{
        auction_name = "Brolyn Auctions"
        pickup_address = "[To be filled from invoice]"
        logistics_contact = @{
            phone = "(574) 891-3111"
            email = "logistics@brolynauctions.com"
        }
        pickup_datetime = "[Specify pickup date/time from invoice]"
        delivery_datetime = "[Next business day] between 9:00am and 5:00pm EST"
        delivery_notice = "Driver must call at least one hour prior to delivery"
        special_notes = @(
            "Forklift available on site for loading",
            "Delivery location has dock with pallet jacks only",
            "Trucks must back up to dock for unloading",
            "Total weight will NOT exceed standard truck capacity"
        )
    }
    delivery_address = "1218 Lake Avenue, Ashtabula, OH 44004"
    shipping_requirements = @{
        total_pallets = "TBD"
        truck_types = "Two trucks: 53' dry van + Flatbed (tarped)"
        labor_needed = "1-2 people for consolidation, freight prep, and loading"
        weight_notes = "Total weight unknown but will NOT exceed standard load capacity"
    }
}

# Ritchie Bros Template
$RitchieBrosTemplate = @{
    email_subject = "Freight Quote Request - [Location] to Ashtabula, OH - Pickup [Date]"
    auction_info = @{
        auction_name = "Ritchie Bros"
        pickup_address = "[To be filled]"
        logistics_contact = @{
            phone = "[To be filled]"
            email = "[To be filled]"
        }
        pickup_datetime = "[Specify pickup date/time]"
        delivery_datetime = "[Specify delivery window]"
        delivery_notice = "Driver must call prior to delivery"
        special_notes = @(
            "Loading assistance may be available on site",
            "Delivery location has dock with pallet jacks only",
            "Driver must call 1 hour before delivery"
        )
    }
    delivery_address = "1218 Lake Avenue, Ashtabula, OH 44004"
    shipping_requirements = @{
        total_pallets = "TBD"
        truck_types = "TBD - Please recommend based on items"
        labor_needed = "TBD based on items"
        weight_notes = "Total weight will NOT exceed standard truck capacity"
    }
}

# Purple Wave Template
$PurpleWaveTemplate = @{
    email_subject = "Freight Quote Request - [Location] to Ashtabula, OH - Pickup [Date]"
    auction_info = @{
        auction_name = "Purple Wave"
        pickup_address = "[To be filled]"
        logistics_contact = @{
            phone = "[To be filled]"
            email = "[To be filled]"
        }
        pickup_datetime = "[Specify pickup date/time]"
        delivery_datetime = "[Specify delivery window]"
        delivery_notice = "Driver must schedule delivery time in advance"
        special_notes = @(
            "Buyer responsible for loading arrangements",
            "Delivery location has dock with pallet jacks only"
        )
    }
    delivery_address = "1218 Lake Avenue, Ashtabula, OH 44004"
    shipping_requirements = @{
        total_pallets = "TBD"
        truck_types = "TBD - Please recommend based on items"
        labor_needed = "Loading assistance required"
        weight_notes = "Weight details to be provided"
    }
}

# GovDeals Template
$GovDealsTemplate = @{
    email_subject = "Freight Quote Request - [Location] to Ashtabula, OH - Pickup [Date]"
    auction_info = @{
        auction_name = "GovDeals"
        pickup_address = "[Government facility address]"
        logistics_contact = @{
            phone = "[To be filled]"
            email = "[To be filled]"
        }
        pickup_datetime = "[Specify pickup appointment]"
        delivery_datetime = "[Specify delivery window]"
        delivery_notice = "Driver must call ahead - government facility"
        special_notes = @(
            "Government facility - strict pickup appointments required",
            "May require ID and vehicle inspection at gate",
            "Loading assistance typically NOT available",
            "Delivery location has dock with pallet jacks only"
        )
    }
    delivery_address = "1218 Lake Avenue, Ashtabula, OH 44004"
    shipping_requirements = @{
        total_pallets = "TBD"
        truck_types = "TBD - Please recommend based on items"
        labor_needed = "Loading crew required - no on-site assistance"
        weight_notes = "Weight details to be provided"
    }
}

# Generic Template
$GenericTemplate = @{
    email_subject = "Freight Quote Request - [Pickup City, ST] to [Delivery City, ST] - Pickup [Date]"
    auction_info = @{
        auction_name = "[Auction Company Name]"
        pickup_address = "[Full pickup address]"
        logistics_contact = @{
            phone = "[Contact phone]"
            email = "[Contact email]"
        }
        pickup_datetime = "[Pickup date and time]"
        delivery_datetime = "[Delivery date and time window]"
        delivery_notice = "[Any delivery notice requirements]"
        special_notes = @(
            "[Add special notes as needed]"
        )
    }
    delivery_address = "1218 Lake Avenue, Ashtabula, OH 44004"
    shipping_requirements = @{
        total_pallets = "TBD"
        truck_types = "TBD - Please recommend based on items"
        labor_needed = "TBD based on requirements"
        weight_notes = "Total weight will NOT exceed standard truck capacity"
    }
}

#endregion

#region Save Templates

$templates = @{
    "Brolyn_Auctions" = $BrolynTemplate
    "Ritchie_Bros" = $RitchieBrosTemplate
    "Purple_Wave" = $PurpleWaveTemplate
    "GovDeals" = $GovDealsTemplate
    "Generic_Template" = $GenericTemplate
}

Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       CREATING STARTER TEMPLATES                      ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

foreach ($templateName in $templates.Keys) {
    $templatePath = Join-Path $TemplateDirectory "$templateName.json"
    
    try {
        $templates[$templateName] | ConvertTo-Json -Depth 10 | Out-File $templatePath -Encoding UTF8
        Write-Host "✅ Created: $templateName" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to create: $templateName - $_" -ForegroundColor Red
    }
}

#endregion

Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║              TEMPLATES CREATED SUCCESSFULLY            ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Green

Write-Host "Templates saved to: $TemplateDirectory`n" -ForegroundColor Cyan

Write-Host "Available Templates:" -ForegroundColor Yellow
Write-Host "  1. Brolyn_Auctions     - Pre-configured for Brolyn Auctions" -ForegroundColor White
Write-Host "  2. Ritchie_Bros        - Pre-configured for Ritchie Bros" -ForegroundColor White
Write-Host "  3. Purple_Wave         - Pre-configured for Purple Wave" -ForegroundColor White
Write-Host "  4. GovDeals            - Pre-configured for GovDeals (government)" -ForegroundColor White
Write-Host "  5. Generic_Template    - Blank template for any auction`n" -ForegroundColor White

Write-Host "💡 Usage:" -ForegroundColor Cyan
Write-Host "   Open the Configuration GUI and select a template from the" -ForegroundColor White
Write-Host "   Template Manager panel to quickly load pre-configured settings.`n" -ForegroundColor White

Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")