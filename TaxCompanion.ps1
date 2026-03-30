param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AppName   = 'Tax Companion (Bangladesh) - PowerShell'
$AppAuthor = 'rhshourav'
$AppGitHub = 'github.com/rhshourav'
$Version   = '1.0'

function Get-DefaultConfig {
    [pscustomobject]@{
        TaxYear               = '2025-26'
        SalaryExemptionCap    = 500000
        MinimumTax            = 3000
        FestivalBonusRatio    = 42714.0 / 344475.0
        RebatePctOfTaxable    = 0.03
        RebatePctOfInvestment = 0.15
        RebateMaxAmount       = 1000000
        SurchargeThreshold    = 40000000
        SurchargeAutoTiers    = @(
            [pscustomobject]@{ MaxNetWealth = 50000000;  Rate = 0.10 },
            [pscustomobject]@{ MaxNetWealth = 100000000; Rate = 0.20 },
            [pscustomobject]@{ MaxNetWealth = [double]::MaxValue; Rate = 0.35 }
        )
        TaxSlabs = @(
            [pscustomobject]@{ Limit = 350000; Rate = 0.00; Label = 'First Tk 3,50,000 (0%)' },
            [pscustomobject]@{ Limit = 100000; Rate = 0.05; Label = 'Next Tk 1,00,000 (5%)' },
            [pscustomobject]@{ Limit = 400000; Rate = 0.10; Label = 'Next Tk 4,00,000 (10%)' },
            [pscustomobject]@{ Limit = 500000; Rate = 0.15; Label = 'Next Tk 5,00,000 (15%)' },
            [pscustomobject]@{ Limit = 500000; Rate = 0.20; Label = 'Next Tk 5,00,000 (20%)' },
            [pscustomobject]@{ Limit = 2000000; Rate = 0.25; Label = 'Next Tk 20,00,000 (25%)' },
            [pscustomobject]@{ Limit = -1; Rate = 0.25; Label = 'On Balance (25%)' }
        )
    }
}

function Get-Config {
    $cfg = Get-DefaultConfig
    if (Test-Path -LiteralPath 'config.json') {
        try {
            $loaded = Get-Content -LiteralPath 'config.json' -Raw | ConvertFrom-Json
            foreach ($p in $loaded.PSObject.Properties) {
                if ($null -ne $p.Value -and $p.Name -ne 'SurchargeAutoTiers' -and $p.Name -ne 'TaxSlabs') {
                    $cfg.$($p.Name) = $p.Value
                }
            }
            if ($loaded.SurchargeAutoTiers) { $cfg.SurchargeAutoTiers = $loaded.SurchargeAutoTiers }
            if ($loaded.TaxSlabs) { $cfg.TaxSlabs = $loaded.TaxSlabs }
        } catch {
            Write-Warning "config.json could not be read; defaults used."
        }
    }
    return $cfg
}

function Format-Money {
    param([double]$Value)
    $n = [math]::Round($Value)
    return '{0:N0}' -f $n
}

function Round-Taka {
    param([double]$Value)
    [math]::Round($Value)
}

function Test-Bool {
    param(
        [string]$Value,
        [bool]$Default = $false
    )
    $s = $Value
    if ($null -eq $s) { $s = '' }
    $s = $s.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($s)) { return $Default }
    switch ($s) {
        'y' { return $true }
        'yes' { return $true }
        'true' { return $true }
        '1' { return $true }
        'on' { return $true }
        'n' { return $false }
        'no' { return $false }
        'false' { return $false }
        '0' { return $false }
        'off' { return $false }
        default { return $Default }
    }
}

function Convert-HumanNumber {
    param([string]$Text)

    $s = $Text
    if ($null -eq $s) { $s = '' }
    $s = $s.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($s)) { return 0.0 }

    $s = $s -replace ',', ''
    $s = [regex]::Replace($s, '(\d+(?:\.\d+)?)\s*(k|thousand)\b', '$1*1000')
    $s = [regex]::Replace($s, '(\d+(?:\.\d+)?)\s*(m|mn|million)\b', '$1*1000000')
    $s = [regex]::Replace($s, '(\d+(?:\.\d+)?)\s*(lakh|lac|lacs)\b', '$1*100000')
    $s = [regex]::Replace($s, '(\d+(?:\.\d+)?)\s*(cr|crore|crores)\b', '$1*10000000')
    $s = [regex]::Replace($s, '(\d+(?:\.\d+)?)\s*%', '($1/100)')

    try {
        $dt = New-Object System.Data.DataTable
        $dt.Compute($s, $null) | ForEach-Object { [double]$_ }
    } catch {
        return 0.0
    }
}

function Get-Val {
    param([string]$Value, [string]$Default)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    return $Value.Trim()
}

function Calculate-Tax {
    param(
        [double]$Taxable,
        $Slabs,
        [double]$MinimumTax
    )

    $remaining = $Taxable
    $total = 0.0
    $lines = New-Object System.Collections.Generic.List[object]

    foreach ($slab in $Slabs) {
        if ($remaining -le 0) { break }

        $amt = $remaining
        if ($slab.Limit -gt 0 -and $amt -gt $slab.Limit) { $amt = $slab.Limit }

        $tax = $amt * [double]$slab.Rate
        $total += $tax
        $lines.Add([pscustomobject]@{
            Label  = $slab.Label
            Amount = $amt
            Rate   = [double]$slab.Rate
            Tax    = $tax
        })

        $remaining -= $amt
    }

    if ($total -gt 0 -and $total -lt $MinimumTax) {
        $lines.Add([pscustomobject]@{
            Label  = 'Minimum tax floor'
            Amount = 0
            Rate   = 0
            Tax    = $MinimumTax - $total
        })
        $total = $MinimumTax
    }

    [pscustomobject]@{
        Tax  = [math]::Round($total)
        Lines = $lines
    }
}

function Calculate-Rebate {
    param(
        [double]$Taxable,
        [double]$Investment,
        $Cfg,
        [double]$TaxBefore
    )

    if ($Taxable -le 0 -or $Investment -le 0 -or $TaxBefore -le 0) { return 0.0 }
    $c1 = $Taxable * [double]$Cfg.RebatePctOfTaxable
    $c2 = $Investment * [double]$Cfg.RebatePctOfInvestment
    $rebate = [math]::Min($c1, [math]::Min($c2, [double]$Cfg.RebateMaxAmount))
    $rebate = [math]::Round($rebate)
    if ($rebate -gt $TaxBefore) { $rebate = $TaxBefore }
    if ($rebate -lt 0) { $rebate = 0 }
    return $rebate
}


function Estimate-Allowances {
    param(
        [double]$BaseGross
    )

    if ($BaseGross -le 0) {
        return [pscustomobject]@{
            Basic = 0.0
            HRA = 0.0
            Medical = 0.0
            Conveyance = 0.0
        }
    }

    # Standardized display-only split for non-custom salary mode.
    # Assumes the gross package is composed roughly as:
    # basic 60.6%, house rent 30.3%, medical 6.1%, conveyance 3.0%
    # This keeps the fields visible without changing the tax result.
    $basic = [math]::Round($BaseGross / 1.65)
    $hra = [math]::Round($basic * 0.50)
    $medical = [math]::Round($basic * 0.10)
    $conveyance = [math]::Round($basic * 0.05)

    return [pscustomobject]@{
        Basic = $basic
        HRA = $hra
        Medical = $medical
        Conveyance = $conveyance
    }
}

function Determine-SurchargeRate {
    param(
        [double]$NetWealth,
        [bool]$Apply,
        [string]$Mode,
        $Cfg
    )

    if (-not $Apply -or $NetWealth -le [double]$Cfg.SurchargeThreshold) { return 0.0 }

    if ([string]::IsNullOrWhiteSpace($Mode) -or $Mode.ToLowerInvariant() -eq 'auto') {
        foreach ($tier in $Cfg.SurchargeAutoTiers) {
            if ($NetWealth -le [double]$tier.MaxNetWealth) { return [double]$tier.Rate }
        }
        return 0.0
    }

    $p = $Mode.Trim().TrimEnd('%')
    try {
        $v = [double]::Parse($p, [System.Globalization.CultureInfo]::InvariantCulture)
        return $v / 100.0
    } catch {
        return [double]$Cfg.SurchargeAutoTiers[0].Rate
    }
}

function Get-ExpenseKeys {
    @(
        'Food, Clothing and Essentials'
        'Accommodation Expense'
        'Electricity'
        'Gas, Water, Sewer and Garbage'
        'Phone, Internet, TV & Subs'
        'Home-Support & Other Expenses'
        'Education Expenses'
        'Festival, Party, Events'
    )
}

function Compute-Allocation {
    param(
        [double]$Total,
        [string]$Loc,
        [int]$FamilySize,
        [bool]$HasKids,
        [bool]$OwnHome,
        [bool]$Staff,
        [string]$Mode
    )

    $weights = [ordered]@{
        'Food, Clothing and Essentials' = 30.0
        'Accommodation Expense' = 28.0
        'Electricity' = 2.5
        'Gas, Water, Sewer and Garbage' = 3.0
        'Phone, Internet, TV & Subs' = 3.5
        'Home-Support & Other Expenses' = 7.0
        'Education Expenses' = 10.0
        'Festival, Party, Events' = 6.0
    }

    if (-not $HasKids) { $weights['Education Expenses'] = 0.0 }

    if ($Loc -match 'dhaka|city|metro') {
        $weights['Accommodation Expense'] *= 1.20
        $weights['Food, Clothing and Essentials'] *= 1.05
        $weights['Home-Support & Other Expenses'] *= 1.10
    } else {
        $weights['Accommodation Expense'] *= 0.90
    }

    $extra = [math]::Max(0, $FamilySize - 2)
    if ($extra -gt 0) {
        $weights['Food, Clothing and Essentials'] *= (1 + 0.05 * $extra)
        if ($HasKids) { $weights['Education Expenses'] *= (1 + 0.04 * $extra) }
    }

    if ($OwnHome) {
        $weights['Accommodation Expense'] *= 0.60
        $weights['Home-Support & Other Expenses'] *= 1.05
    }
    if (-not $Staff) {
        $weights['Home-Support & Other Expenses'] *= 0.40
    }

    switch ($Mode.ToLowerInvariant()) {
        'conservative' {
            $weights['Festival, Party, Events'] *= 0.60
            $weights['Home-Support & Other Expenses'] *= 0.70
            $weights['Food, Clothing and Essentials'] *= 1.08
        }
        'comfortable' {
            $weights['Festival, Party, Events'] *= 1.30
            $weights['Home-Support & Other Expenses'] *= 1.20
            $weights['Food, Clothing and Essentials'] *= 1.05
        }
    }

    $totalWeight = ($weights.Values | Measure-Object -Sum).Sum
    if ($totalWeight -le 0) {
        return [pscustomobject]@{ Pcts = @{}; Amts = @{} }
    }

    $pcts = @{}
    $remainders = New-Object System.Collections.Generic.List[object]
    $sumPct = 0

    foreach ($k in $weights.Keys) {
        $raw = ($weights[$k] / $totalWeight) * 100.0
        $base = [math]::Floor($raw)
        $pcts[$k] = [int]$base
        $sumPct += [int]$base
        $remainders.Add([pscustomobject]@{ Key = $k; Rem = $raw - $base })
    }

    $remainders = $remainders | Sort-Object Rem -Descending
    for ($i = 0; $i -lt (100 - $sumPct); $i++) {
        $k = $remainders[$i % $remainders.Count].Key
        $pcts[$k]++
    }

    $amts = @{}
    $amtRem = New-Object System.Collections.Generic.List[object]
    $totalInt = [int](Round-Taka $Total)
    $allocated = 0

    foreach ($k in $pcts.Keys) {
        $rawAmt = $totalInt * $pcts[$k] / 100.0
        $base = [math]::Floor($rawAmt)
        $amts[$k] = [int]$base
        $allocated += [int]$base
        $amtRem.Add([pscustomobject]@{ Key = $k; Rem = $rawAmt - $base })
    }

    $drift = $totalInt - $allocated
    if ($drift -gt 0) {
        $amtRem = $amtRem | Sort-Object Rem -Descending
        for ($i = 0; $i -lt $drift; $i++) {
            $amts[$amtRem[$i % $amtRem.Count].Key]++
        }
    } elseif ($drift -lt 0) {
        $amtRem = $amtRem | Sort-Object Rem
        for ($i = 0; $i -lt (-$drift); $i++) {
            $amts[$amtRem[$i % $amtRem.Count].Key]--
        }
    }

    [pscustomobject]@{ Pcts = $pcts; Amts = $amts }
}

function Estimate-NetFromGross {
    param(
        [double]$BaseGross,
        [double]$Deductions,
        $Cfg
    )

    $bonus = [math]::Round($BaseGross * [double]$Cfg.FestivalBonusRatio)
    $total = $BaseGross + $bonus
    $exempt = [math]::Min($total / 3.0, [double]$Cfg.SalaryExemptionCap)
    $taxable = [math]::Max(0, $total - $exempt)
    $tax = (Calculate-Tax -Taxable $taxable -Slabs $Cfg.TaxSlabs -MinimumTax ([double]$Cfg.MinimumTax)).Tax
    return $total - $tax - $Deductions
}

function Estimate-GrossFromNet {
    param(
        [double]$TargetNet,
        [double]$Deductions,
        $Cfg
    )

    $low = 0.0
    $high = [math]::Max(1000000, $TargetNet * 3 + 1000000)
    $best = 0.0
    $bestDiff = [double]::MaxValue

    for ($i = 0; $i -lt 80; $i++) {
        $mid = ($low + $high) / 2.0
        $net = Estimate-NetFromGross -BaseGross $mid -Deductions $Deductions -Cfg $Cfg
        $diff = $net - $TargetNet
        $abs = [math]::Abs($diff)
        if ($abs -lt $bestDiff) {
            $bestDiff = $abs
            $best = $mid
        }
        if ($diff -gt 0) {
            $high = $mid
        } else {
            $low = $mid
        }
    }

    return [math]::Round($best)
}

function Build-Report {
    param(
        [hashtable]$Inputs,
        $Cfg
    )

    $totalInput    = Convert-HumanNumber $Inputs.TotalSalary
    $bonusIncluded = Test-Bool $Inputs.BonusIncluded $false
    $customSalary  = Test-Bool $Inputs.CustomSalary $false

    $basic    = Convert-HumanNumber $Inputs.Basic
    $hra      = Convert-HumanNumber $Inputs.HRA
    $med      = Convert-HumanNumber $Inputs.Medical
    $food     = Convert-HumanNumber $Inputs.Food
    $trans    = Convert-HumanNumber $Inputs.Transport
    $mobile   = Convert-HumanNumber $Inputs.Mobile
    $festivalPct = Convert-HumanNumber $Inputs.FestivalPct
    $totalExpense = Convert-HumanNumber $Inputs.TotalExpense
    $location = (Get-Val $Inputs.Location 'other').ToLowerInvariant()
    $familySize = [int][math]::Round((Convert-HumanNumber $Inputs.FamilySize))
    if ($familySize -le 0) { $familySize = 3 }
    $hasKids = Test-Bool $Inputs.HasKids $false
    $ownHome = Test-Bool $Inputs.OwnHome $false
    $hasStaff = Test-Bool $Inputs.HasStaff $false
    $mode = (Get-Val $Inputs.Mode 'balanced').ToLowerInvariant()
    $prevGross = Convert-HumanNumber $Inputs.PreviousGross
    $netWealthInput = Convert-HumanNumber $Inputs.NetWealth
    $openingWealth = Convert-HumanNumber $Inputs.OpeningNetWealth
    $totalAssets = Convert-HumanNumber $Inputs.TotalAssets
    $bankLoan = Convert-HumanNumber $Inputs.BankLoan
    $personalLoan = Convert-HumanNumber $Inputs.PersonalLoan
    $creditCard = Convert-HumanNumber $Inputs.CreditCard
    $otherLiabilities = Convert-HumanNumber $Inputs.OtherLiabilities
    $applySurcharge = Test-Bool $Inputs.ApplySurcharge $false
    $surchargeMode = (Get-Val $Inputs.SurchargeMode 'auto').ToLowerInvariant()
    $investment = Convert-HumanNumber $Inputs.Investment
    $reverseEnabled = Test-Bool $Inputs.ReverseEnabled $false
    $targetNet = Convert-HumanNumber $Inputs.TargetNet
    $deductions = Convert-HumanNumber $Inputs.Deductions

    $baseGross = 0.0
    $bonus = 0.0
    $totalComp = 0.0

    if ($customSalary) {
        $baseGross = Round-Taka ($basic + $hra + $med + $food + $trans + $mobile)
        $totalComp = Round-Taka $totalInput
        if ($totalComp -le 0) { $totalComp = $baseGross }
        if ($festivalPct -le 0) { $festivalPct = 100.0 }

        if ($bonusIncluded) {
            if ($totalComp -lt $baseGross) {
                $bonus = 0
                $totalComp = $baseGross
            } else {
                $bonus = $totalComp - $baseGross
            }
        } else {
            $monthlyBasic = $basic / 12.0
            $bonus = Round-Taka ($monthlyBasic * ($festivalPct / 100.0))
            $totalComp = $baseGross + $bonus
        }
    } else {
        $est = Estimate-Allowances -BaseGross $totalInput
        $basic = $est.Basic
        $hra = $est.HRA
        $med = $est.Medical
        $trans = $est.Conveyance

        if ($bonusIncluded) {
            $totalComp = Round-Taka $totalInput
            if ($totalComp -le 0) { $totalComp = 0 }
            $baseGross = Round-Taka ($totalComp / (1 + [double]$Cfg.FestivalBonusRatio))
            $bonus = $totalComp - $baseGross
        } else {
            $baseGross = Round-Taka $totalInput
            $totalComp = $baseGross
            if ($baseGross -gt 0) {
                $bonus = Round-Taka ($baseGross * [double]$Cfg.FestivalBonusRatio)
                $totalComp = $baseGross + $bonus
            }
        }
    }

    $salaryExempt = [math]::Min($totalComp / 3.0, [double]$Cfg.SalaryExemptionCap)
    $taxableSalary = [math]::Max(0, $totalComp - $salaryExempt)

    $taxCurrent = Calculate-Tax -Taxable $taxableSalary -Slabs $Cfg.TaxSlabs -MinimumTax ([double]$Cfg.MinimumTax)
    $rebate = Calculate-Rebate -Taxable $taxableSalary -Investment $investment -Cfg $Cfg -TaxBefore $taxCurrent.Tax
    $taxAfterRebate = [math]::Max(0, $taxCurrent.Tax - $rebate)

    $prevTax = 0.0
    $prevTaxLines = @()
    if ($prevGross -gt 0) {
        $prevExempt = [math]::Min($prevGross / 3.0, [double]$Cfg.SalaryExemptionCap)
        $prevTaxable = [math]::Max(0, $prevGross - $prevExempt)
        $prevTaxRes = Calculate-Tax -Taxable $prevTaxable -Slabs $Cfg.TaxSlabs -MinimumTax ([double]$Cfg.MinimumTax)
        $prevTax = $prevTaxRes.Tax
        $prevTaxLines = $prevTaxRes.Lines
    }

    $combinedBeforeSurch = $taxAfterRebate + $prevTax
    $totalLiabilities = $bankLoan + $personalLoan + $creditCard + $otherLiabilities
    $netWealthUsed = $netWealthInput
    if ($totalAssets -gt 0) { $netWealthUsed = $totalAssets - $totalLiabilities }
    $surchargeRate = Determine-SurchargeRate -NetWealth $netWealthUsed -Apply $applySurcharge -Mode $surchargeMode -Cfg $Cfg
    $surchargeAmount = 0.0
    if ($surchargeRate -gt 0 -and $netWealthUsed -gt [double]$Cfg.SurchargeThreshold) {
        $surchargeAmount = Round-Taka ($combinedBeforeSurch * $surchargeRate)
    }
    $finalTax = $combinedBeforeSurch + $surchargeAmount

    $alloc = $null
    if ($totalExpense -gt 0) {
        $alloc = Compute-Allocation -Total $totalExpense -Loc $location -FamilySize $familySize -HasKids $hasKids -OwnHome $ownHome -Staff $hasStaff -Mode $mode
    }

    $estimatedGross = 0.0
    if ($reverseEnabled -and $targetNet -gt 0) {
        $estimatedGross = Estimate-GrossFromNet -TargetNet $targetNet -Deductions $deductions -Cfg $Cfg
    }

    $wealthIncrease = $netWealthUsed - $openingWealth
    $expectedSavings = $totalComp - $deductions - $totalExpense - $finalTax
    $wealthDifference = $wealthIncrease - $expectedSavings
    $tol = [math]::Max(10000, [math]::Abs($expectedSavings) * 0.01)

    if ([math]::Abs($wealthDifference) -le $tol) {
        $wealthStatus = 'OK — wealth increase matches estimated savings within tolerance.'
        $auditRisk = 'LOW'
    } elseif ($wealthDifference -gt 0) {
        $wealthStatus = 'ALERT — reported wealth is higher than estimated savings.'
        $auditRisk = 'HIGH'
    } else {
        $wealthStatus = 'Wealth increase is lower than estimated savings.'
        $auditRisk = 'MEDIUM'
    }

    [pscustomobject]@{
        GeneratedAt = Get-Date
        TaxYear = $Cfg.TaxYear

        SalaryInput = $totalInput
        BonusIncluded = $bonusIncluded
        CustomSalary = $customSalary
        BaseGrossSalary = $baseGross
        HouseRentAllowance = $hra
        MedicalAllowance = $med
        ConveyanceAllowance = $trans
        BonusAmount = $bonus
        TotalComp = $totalComp
        SalaryExempt = $salaryExempt
        TaxableSalary = $taxableSalary

        CurrentTaxBeforeRebate = $taxCurrent.Tax
        RebateEligibleInvest = $investment
        RebateAmount = $rebate
        CurrentTaxAfterRebate = $taxAfterRebate

        PreviousGrossInput = $prevGross
        PreviousTax = $prevTax
        CombinedBeforeSurch = $combinedBeforeSurch

        ApplySurcharge = $applySurcharge
        SurchargeRate = $surchargeRate
        SurchargeAmount = $surchargeAmount
        FinalTax = $finalTax

        TotalExpense = $totalExpense
        ExpensePcts = if ($alloc) { $alloc.Pcts } else { @{} }
        ExpenseAmts = if ($alloc) { $alloc.Amts } else { @{} }

        TotalAssets = [math]::Round($totalAssets)
        TotalLiabilities = [math]::Round($totalLiabilities)
        NetWealthCurrent = [math]::Round($netWealthUsed)
        OpeningNetWealth = $openingWealth
        WealthIncrease = $wealthIncrease
        ExpectedSavings = $expectedSavings
        WealthDifference = $wealthDifference
        WealthStatus = $wealthStatus

        AuditRisk = $auditRisk

        ReverseEnabled = $reverseEnabled
        TargetNetTakeHome = $targetNet
        Deductions = $deductions
        EstimatedGrossFromNet = $estimatedGross

        CurrentTaxLines = $taxCurrent.Lines
        PrevTaxLines = $prevTaxLines
    }
}

function Get-ExpenseOrder {
    Get-ExpenseKeys
}

function Format-TaxTable {
    param($Lines)
    if (-not $Lines -or $Lines.Count -eq 0) { return "No taxable amount.`n" }
    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine(('SLAB'.PadRight(34) + ' | ' + 'TAXED'.PadRight(12) + ' | ' + 'RATE'.PadRight(8) + ' | ' + 'TAX'.PadRight(12)))
    [void]$out.AppendLine((''.PadRight(78, '-')))
    foreach ($ln in $Lines) {
        [void]$out.AppendLine((
            $ln.Label.PadRight(34) + ' | ' +
            (Format-Money $ln.Amount).PadLeft(12) + ' | ' +
            ('{0:0}%' -f ($ln.Rate * 100)).PadLeft(7) + ' | ' +
            (Format-Money $ln.Tax).PadLeft(12)
        ))
    }
    [void]$out.AppendLine((''.PadRight(78, '-')))
    return $out.ToString()
}

function Format-KvLine {
    param([string]$Label, [double]$Value)
    '{0,-34} Tk {1}' -f $Label, (Format-Money $Value)
}

function New-ReportText {
    param($R)

    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine('SUMMARY')
    [void]$sb.AppendLine(('Assessment year'.PadRight(34) + ' ' + $R.TaxYear))
    [void]$sb.AppendLine(('Generated'.PadRight(34) + ' ' + $R.GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')))
    [void]$sb.AppendLine()

    [void]$sb.AppendLine('SALARY / INCOME')
    [void]$sb.AppendLine((Format-KvLine 'Annual salary input' $R.SalaryInput))
    [void]$sb.AppendLine((Format-KvLine 'House rent allowance' $R.HouseRentAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Medical allowance' $R.MedicalAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Conveyance allowance' $R.ConveyanceAllowance))
    if (-not $R.CustomSalary) {
        [void]$sb.AppendLine('Allowance note: estimated split from the gross salary because custom breakdown is off.')
    }
    if ($R.CustomSalary) {
        [void]$sb.AppendLine((Format-KvLine 'Basic gross salary' $R.BaseGrossSalary))
        [void]$sb.AppendLine((Format-KvLine 'Festival bonus (one-month %)' $R.BonusAmount))
    } else {
        [void]$sb.AppendLine((Format-KvLine 'Base gross salary' $R.BaseGrossSalary))
        [void]$sb.AppendLine((Format-KvLine 'Festival bonus' $R.BonusAmount))
    }
    [void]$sb.AppendLine((Format-KvLine 'Total compensation' $R.TotalComp))
    [void]$sb.AppendLine((Format-KvLine 'Salary exemption' $R.SalaryExempt))
    [void]$sb.AppendLine((Format-KvLine 'Taxable income' $R.TaxableSalary))
    [void]$sb.AppendLine()

    [void]$sb.AppendLine('CURRENT YEAR TAX')
    [void]$sb.AppendLine((Format-TaxTable $R.CurrentTaxLines))
    [void]$sb.AppendLine((Format-KvLine 'Tax before rebate' $R.CurrentTaxBeforeRebate))
    [void]$sb.AppendLine((Format-KvLine 'Eligible investment' $R.RebateEligibleInvest))
    [void]$sb.AppendLine((Format-KvLine 'Tax rebate (Section 44)' $R.RebateAmount))
    [void]$sb.AppendLine((Format-KvLine 'Tax after rebate' $R.CurrentTaxAfterRebate))
    [void]$sb.AppendLine()

    if ($R.PreviousGrossInput -gt 0) {
        [void]$sb.AppendLine('PREVIOUS YEAR TAX')
        [void]$sb.AppendLine((Format-KvLine 'Previous gross income' $R.PreviousGrossInput))
        [void]$sb.AppendLine((Format-TaxTable $R.PrevTaxLines))
        [void]$sb.AppendLine((Format-KvLine 'Previous year tax' $R.PreviousTax))
        [void]$sb.AppendLine()
    }

    if ($R.ApplySurcharge) {
        [void]$sb.AppendLine('NET WEALTH SURCHARGE')
        [void]$sb.AppendLine(('Applied'.PadRight(34) + ' ' + ($(if ($R.SurchargeRate -gt 0) { 'Yes' } else { 'No' }))))
        [void]$sb.AppendLine(('Surcharge rate'.PadRight(34) + ' ' + ('{0:N2}%' -f ($R.SurchargeRate * 100))))
        [void]$sb.AppendLine((Format-KvLine 'Surcharge amount' $R.SurchargeAmount))
        [void]$sb.AppendLine((Format-KvLine 'Combined tax before surcharge' $R.CombinedBeforeSurch))
        [void]$sb.AppendLine((Format-KvLine 'Final tax' $R.FinalTax))
        [void]$sb.AppendLine()
    } else {
        [void]$sb.AppendLine('FINAL TAX')
        [void]$sb.AppendLine((Format-KvLine 'Final tax' $R.FinalTax))
        [void]$sb.AppendLine()
    }

    if ($R.TotalExpense -gt 0) {
        [void]$sb.AppendLine('IT-10BB EXPENSE ALLOCATION')
        foreach ($k in (Get-ExpenseOrder)) {
            if ($R.ExpensePcts[$k] -eq 0 -and $R.ExpenseAmts[$k] -eq 0) { continue }
            [void]$sb.AppendLine(('{0,-34} {1,3}%   Tk {2}' -f $k, $R.ExpensePcts[$k], (Format-Money $R.ExpenseAmts[$k])))
        }
        [void]$sb.AppendLine(('{0,-34} {1,3}%   Tk {2}' -f 'TOTAL', 100, (Format-Money ([int](Round-Taka $R.TotalExpense)))))
        [void]$sb.AppendLine()
    }

    [void]$sb.AppendLine('WEALTH SUMMARY')
    [void]$sb.AppendLine((Format-KvLine 'Total assets (provided)' $R.TotalAssets))
    [void]$sb.AppendLine((Format-KvLine 'Total liabilities' $R.TotalLiabilities))
    [void]$sb.AppendLine((Format-KvLine 'Net wealth (used)' $R.NetWealthCurrent))
    [void]$sb.AppendLine(('Audit risk'.PadRight(34) + ' ' + $R.AuditRisk))
    [void]$sb.AppendLine()

    [void]$sb.AppendLine('WEALTH CHECK')
    [void]$sb.AppendLine((Format-KvLine 'Opening net wealth' $R.OpeningNetWealth))
    [void]$sb.AppendLine((Format-KvLine 'Wealth increase' $R.WealthIncrease))
    [void]$sb.AppendLine((Format-KvLine 'Estimated after-tax savings' $R.ExpectedSavings))
    [void]$sb.AppendLine($R.WealthStatus)
    [void]$sb.AppendLine()

    if ($R.ReverseEnabled -and $R.TargetNetTakeHome -gt 0) {
        [void]$sb.AppendLine('REVERSE CALCULATOR')
        [void]$sb.AppendLine((Format-KvLine 'Target net take-home' $R.TargetNetTakeHome))
        [void]$sb.AppendLine((Format-KvLine 'Extra deductions' $R.Deductions))
        [void]$sb.AppendLine((Format-KvLine 'Estimated gross salary' $R.EstimatedGrossFromNet))
        [void]$sb.AppendLine()
    }

    return $sb.ToString()
}

function Save-Session {
    param($Path, $Inputs)
    $session = [pscustomobject]@{
        Version = 1
        Timestamp = (Get-Date).ToString('o')
        Inputs = $Inputs
    }
    $raw = $session | ConvertTo-Json -Depth 12
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hashBytes = $sha.ComputeHash($bytes)
        $hash = ([BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
        $wrapped = [pscustomobject]@{
            Hash = $hash
            Session = $session
        }
        $wrapped | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
    } finally {
        $sha.Dispose()
    }
}

function Load-Session {
    param($Path)
    $wrapped = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $raw = $wrapped.Session | ConvertTo-Json -Depth 12
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hashBytes = $sha.ComputeHash($bytes)
        $hash = ([BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
        if ($hash -ne $wrapped.Hash) {
            throw 'SHA256 mismatch: file may be corrupted or tampered with.'
        }
        return $wrapped.Session.Inputs
    } finally {
        $sha.Dispose()
    }
}

function Export-CSVReport {
    param($Path, $R)
    $rows = @(
        [pscustomobject]@{ Field='Tax year'; Value=$R.TaxYear }
        [pscustomobject]@{ Field='Annual salary input'; Value=(Format-Money $R.SalaryInput) }
        [pscustomobject]@{ Field='Base gross salary'; Value=(Format-Money $R.BaseGrossSalary) }
        [pscustomobject]@{ Field='House rent allowance'; Value=(Format-Money $R.HouseRentAllowance) }
        [pscustomobject]@{ Field='Medical allowance'; Value=(Format-Money $R.MedicalAllowance) }
        [pscustomobject]@{ Field='Conveyance allowance'; Value=(Format-Money $R.ConveyanceAllowance) }
        [pscustomobject]@{ Field='Festival bonus'; Value=(Format-Money $R.BonusAmount) }
        [pscustomobject]@{ Field='Total compensation'; Value=(Format-Money $R.TotalComp) }
        [pscustomobject]@{ Field='Salary exemption'; Value=(Format-Money $R.SalaryExempt) }
        [pscustomobject]@{ Field='Taxable income'; Value=(Format-Money $R.TaxableSalary) }
        [pscustomobject]@{ Field='Tax before rebate'; Value=(Format-Money $R.CurrentTaxBeforeRebate) }
        [pscustomobject]@{ Field='Rebate eligible investment'; Value=(Format-Money $R.RebateEligibleInvest) }
        [pscustomobject]@{ Field='Tax rebate'; Value=(Format-Money $R.RebateAmount) }
        [pscustomobject]@{ Field='Tax after rebate'; Value=(Format-Money $R.CurrentTaxAfterRebate) }
        [pscustomobject]@{ Field='Previous year tax'; Value=(Format-Money $R.PreviousTax) }
        [pscustomobject]@{ Field='Surcharge amount'; Value=(Format-Money $R.SurchargeAmount) }
        [pscustomobject]@{ Field='Final tax'; Value=(Format-Money $R.FinalTax) }
        [pscustomobject]@{ Field='Total assets (provided)'; Value=(Format-Money $R.TotalAssets) }
        [pscustomobject]@{ Field='Total liabilities'; Value=(Format-Money $R.TotalLiabilities) }
        [pscustomobject]@{ Field='Current net wealth'; Value=(Format-Money $R.NetWealthCurrent) }
        [pscustomobject]@{ Field='Opening net wealth'; Value=(Format-Money $R.OpeningNetWealth) }
        [pscustomobject]@{ Field='Wealth increase'; Value=(Format-Money $R.WealthIncrease) }
        [pscustomobject]@{ Field='Estimated after-tax savings'; Value=(Format-Money $R.ExpectedSavings) }
        [pscustomobject]@{ Field='Wealth status'; Value=$R.WealthStatus }
        [pscustomobject]@{ Field='Audit risk'; Value=$R.AuditRisk }
    )
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Export-MarkdownReport {
    param($Path, $R)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# Tax Companion Report')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("* Tax year: $($R.TaxYear)")
    [void]$sb.AppendLine("* Generated: $($R.GeneratedAt.ToString('o'))")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('## Summary')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('| Field | Value |')
    [void]$sb.AppendLine('|---|---:|')
    foreach ($pair in @(
        @('Annual salary input', $R.SalaryInput)
        @('Base gross salary', $R.BaseGrossSalary)
        @('House rent allowance', $R.HouseRentAllowance)
        @('Medical allowance', $R.MedicalAllowance)
        @('Conveyance allowance', $R.ConveyanceAllowance)
        @('Festival bonus', $R.BonusAmount)
        @('Total compensation', $R.TotalComp)
        @('Salary exemption', $R.SalaryExempt)
        @('Taxable income', $R.TaxableSalary)
        @('Tax before rebate', $R.CurrentTaxBeforeRebate)
        @('Eligible investment', $R.RebateEligibleInvest)
        @('Tax rebate', $R.RebateAmount)
        @('Tax after rebate', $R.CurrentTaxAfterRebate)
        @('Previous year tax', $R.PreviousTax)
        @('Surcharge amount', $R.SurchargeAmount)
        @('Final tax', $R.FinalTax)
        @('Total assets', $R.TotalAssets)
        @('Total liabilities', $R.TotalLiabilities)
        @('Net wealth (used)', $R.NetWealthCurrent)
        @('Opening net wealth', $R.OpeningNetWealth)
        @('Wealth increase', $R.WealthIncrease)
        @('Estimated after-tax savings', $R.ExpectedSavings)
        @('Wealth status', $R.WealthStatus)
        @('Audit risk', $R.AuditRisk)
    )) {
        [void]$sb.AppendLine("| $($pair[0]) | $([string]$pair[1]) |")
    }
    [void]$sb.AppendLine()
    if ($R.TotalExpense -gt 0) {
        [void]$sb.AppendLine('## Expense allocation')
        [void]$sb.AppendLine()
        [void]$sb.AppendLine('| Category | % | Amount |')
        [void]$sb.AppendLine('|---|---:|---:|')
        foreach ($k in (Get-ExpenseOrder)) {
            if ($R.ExpensePcts[$k] -eq 0 -and $R.ExpenseAmts[$k] -eq 0) { continue }
            [void]$sb.AppendLine("| $k | $($R.ExpensePcts[$k]) | Tk $(Format-Money $R.ExpenseAmts[$k]) |")
        }
        [void]$sb.AppendLine()
    }

    [IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

function Show-Header {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host $AppName -ForegroundColor White
    Write-Host "$AppAuthor • $AppGitHub • $Version" -ForegroundColor DarkGray
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host
}

function Read-Field {
    param([string]$Prompt, [string]$Default)
    $v = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}

function Run-App {
    $cfg = Get-Config
    $inputs = [ordered]@{
        TotalSalary      = ''
        BonusIncluded    = 'n'
        CustomSalary     = 'n'
        Basic            = '0'
        HRA              = '0'
        Medical          = '0'
        Food             = '0'
        Transport        = '0'
        Mobile           = '0'
        FestivalPct      = '100'
        TotalExpense     = '0'
        Location         = 'other'
        FamilySize       = '3'
        HasKids          = 'n'
        OwnHome          = 'n'
        HasStaff         = 'n'
        Mode             = 'balanced'
        PreviousGross    = '0'
        NetWealth        = '0'
        OpeningNetWealth = '0'
        TotalAssets      = '0'
        BankLoan         = '0'
        PersonalLoan     = '0'
        CreditCard       = '0'
        OtherLiabilities = '0'
        ApplySurcharge   = 'n'
        SurchargeMode    = 'auto'
        Investment       = '0'
        ReverseEnabled   = 'n'
        TargetNet        = '0'
        Deductions       = '0'
    }

    Show-Header
    Write-Host "Enter values. Expressions like 12809*23, 1 lakh, 2cr, 4.5k, and 50% are accepted." -ForegroundColor DarkCyan
    Write-Host

    $inputs.TotalSalary      = Read-Field '1. Annual Salary / Total Package (BDT)' $inputs.TotalSalary
    $inputs.BonusIncluded    = Read-Field '2. Festival bonus already included? (y/n)' $inputs.BonusIncluded
    $inputs.CustomSalary     = Read-Field '3. Enter custom salary breakdown? (y/n)' $inputs.CustomSalary

    if (Test-Bool $inputs.CustomSalary $false) {
        $inputs.Basic        = Read-Field '   -> Basic pay (annual BDT)' $inputs.Basic
        $inputs.HRA          = Read-Field '   -> House rent allowance (annual BDT)' $inputs.HRA
        $inputs.Medical      = Read-Field '   -> Medical allowance (annual BDT)' $inputs.Medical
        $inputs.Food         = Read-Field '   -> Food allowance (annual BDT)' $inputs.Food
        $inputs.Transport    = Read-Field '   -> Transport / conveyance (annual BDT)' $inputs.Transport
        $inputs.Mobile       = Read-Field '   -> Mobile & other allowance (annual BDT)' $inputs.Mobile
        $inputs.FestivalPct  = Read-Field '   -> Festival bonus % of ONE MONTH BASIC' $inputs.FestivalPct
    }

    $inputs.TotalExpense     = Read-Field '4. Total annual expense (BDT)' $inputs.TotalExpense
    $inputs.Location         = Read-Field '5. Location (dhaka/other)' $inputs.Location
    $inputs.FamilySize       = Read-Field '6. Family size' $inputs.FamilySize
    $inputs.HasKids          = Read-Field '7. Do you have kids? (y/n)' $inputs.HasKids
    $inputs.OwnHome          = Read-Field '8. Do you own your home? (y/n)' $inputs.OwnHome
    $inputs.HasStaff         = Read-Field '9. Home-support staff? (y/n)' $inputs.HasStaff
    $inputs.Mode             = Read-Field '10. Mode (balanced/conservative/comfortable)' $inputs.Mode
    $inputs.PreviousGross    = Read-Field "11. Previous year's gross income (BDT)" $inputs.PreviousGross
    $inputs.NetWealth        = Read-Field '12. Net wealth (current) (BDT) (fallback)' $inputs.NetWealth
    $inputs.OpeningNetWealth = Read-Field '13. Opening net wealth (previous year) (BDT)' $inputs.OpeningNetWealth
    $inputs.TotalAssets      = Read-Field '14. Total assets (BDT) (overrides net wealth if set)' $inputs.TotalAssets
    $inputs.BankLoan         = Read-Field '15. Bank loan outstanding (BDT)' $inputs.BankLoan
    $inputs.PersonalLoan     = Read-Field '16. Personal loan from others (BDT)' $inputs.PersonalLoan
    $inputs.CreditCard       = Read-Field '17. Credit card dues (BDT)' $inputs.CreditCard
    $inputs.OtherLiabilities = Read-Field '18. Other liabilities (BDT)' $inputs.OtherLiabilities
    $inputs.ApplySurcharge   = Read-Field '19. Apply net wealth surcharge? (y/n)' $inputs.ApplySurcharge
    $inputs.SurchargeMode    = Read-Field "20. Surcharge percent (number or 'auto')" $inputs.SurchargeMode
    $inputs.Investment       = Read-Field '21. Total eligible investment for rebate (BDT)' $inputs.Investment
    $inputs.ReverseEnabled   = Read-Field '22. Reverse calculator? (y/n)' $inputs.ReverseEnabled
    $inputs.TargetNet        = Read-Field '23. Target net take-home pay (BDT)' $inputs.TargetNet
    $inputs.Deductions       = Read-Field '24. Extra deductions / PF / other (BDT)' $inputs.Deductions

    $report = Build-Report -Inputs $inputs -Cfg $cfg
    $reportText = New-ReportText $report

    Show-Header
    Write-Host $reportText

    while ($true) {
        Write-Host
        Write-Host "Commands: [s]ave JSON  [l]oad JSON  [c]sv  [m]d  [r]e-run  [q]uit" -ForegroundColor DarkGray
        $cmd = (Read-Host 'Enter command').Trim().ToLowerInvariant()

        switch ($cmd) {
            's' {
                $path = Read-Host 'Session JSON path' 
                if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_session.json' }
                if (-not $path.EndsWith('.json')) { $path += '.json' }
                Save-Session -Path $path -Inputs $inputs
                Write-Host "Saved: $path" -ForegroundColor Green
            }
            'l' {
                $path = Read-Host 'Session JSON path'
                if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_session.json' }
                if (-not $path.EndsWith('.json')) { $path += '.json' }
                $loaded = Load-Session -Path $path
                $idx = 0
                foreach ($k in $inputs.Keys) {
                    if ($idx -lt $loaded.Count) { $inputs[$k] = [string]$loaded[$idx] }
                    $idx++
                }
                $report = Build-Report -Inputs $inputs -Cfg $cfg
                $reportText = New-ReportText $report
                Show-Header
                Write-Host $reportText
            }
            'c' {
                $path = Read-Host 'CSV path'
                if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_summary.csv' }
                if (-not $path.EndsWith('.csv')) { $path += '.csv' }
                Export-CSVReport -Path $path -R $report
                Write-Host "CSV exported: $path" -ForegroundColor Green
            }
            'm' {
                $path = Read-Host 'Markdown path'
                if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_summary.md' }
                if (-not $path.EndsWith('.md')) { $path += '.md' }
                Export-MarkdownReport -Path $path -R $report
                Write-Host "Markdown exported: $path" -ForegroundColor Green
            }
            'r' {
                Run-App
                return
            }
            'q' {
                return
            }
            default {
                Write-Host 'Unknown command.' -ForegroundColor Yellow
            }
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Run-App
}
