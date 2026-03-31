param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# IT-10BB Tax Companion v3.0.0-clean
# Author: rhshourav
# Repo:   https://github.com/rhshourav/IT10BB-TaxCompanion
# ============================================================

$script:AppName   = 'IT-10BB Tax Companion'
$script:AppAuthor = 'rhshourav'
$script:AppRepo   = 'https://github.com/rhshourav/IT10BB-TaxCompanion'
$script:AppVersion = 'v3.0.0-clean'

function Get-DefaultConfig {
    [pscustomobject]@{
        TaxYear               = '2025-26'
        SalaryExemptionCap    = 500000
        MinimumTax            = 3000
        MonthlyFestivalBonus  = 1100.0
        FestivalBonusRatio    = 42714.0 / 344475.0   # legacy fallback
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

function Load-Config {
    param([string]$Path = 'config.json')
    $cfg = Get-DefaultConfig
    if (Test-Path -LiteralPath $Path) {
        try {
            $loaded = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
            foreach ($p in $loaded.PSObject.Properties) {
                if ($p.Name -ne 'SurchargeAutoTiers' -and $p.Name -ne 'TaxSlabs' -and $null -ne $p.Value) {
                    $cfg.$($p.Name) = $p.Value
                }
            }
            if ($loaded.SurchargeAutoTiers) { $cfg.SurchargeAutoTiers = $loaded.SurchargeAutoTiers }
            if ($loaded.TaxSlabs) { $cfg.TaxSlabs = $loaded.TaxSlabs }
        } catch {
            Write-Warning 'config.json could not be read; defaults used.'
        }
    }
    return $cfg
}

function Format-Money {
    param([double]$Value)
    '{0:N0}' -f ([math]::Round($Value))
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
        $result = $dt.Compute($s, $null)
        return [double]$result
    } catch {
        return 0.0
    }
}

function Get-Val {
    param([string]$Value, [string]$Default)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    return $Value.Trim()
}

function Get-FestivalBonusRate {
    param([string]$EmployeeType)

    switch ((Get-Val $EmployeeType 'staff').Trim().ToLowerInvariant()) {
        'worker' { return 80.0 }
        default { return 100.0 }
    }
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
        if ($slab.Limit -gt 0 -and $amt -gt $slab.Limit) {
            $amt = $slab.Limit
        }

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

    return [pscustomobject]@{
        Tax   = [math]::Round($total)
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

function Determine-SurchargeRate {
    param(
        [double]$NetWealth,
        [bool]$Apply,
        [string]$Mode,
        $Cfg
    )

    if (-not $Apply -or $NetWealth -le [double]$Cfg.SurchargeThreshold) { return 0.0 }

    $modeText = $Mode
    if ($null -eq $modeText) { $modeText = '' }
    $modeText = $modeText.Trim().ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($modeText) -or $modeText -eq 'auto') {
        foreach ($tier in $Cfg.SurchargeAutoTiers) {
            if ($NetWealth -le [double]$tier.MaxNetWealth) { return [double]$tier.Rate }
        }
        return 0.0
    }

    $p = $modeText.TrimEnd('%')
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

    $locText = $Loc
    if ($null -eq $locText) { $locText = '' }
    $locText = $locText.ToLowerInvariant()

    if ($locText -match 'dhaka|city|metro') {
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

    $modeText = $Mode
    if ($null -eq $modeText) { $modeText = '' }
    $modeText = $modeText.ToLowerInvariant()

    switch ($modeText) {
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

    $totalWeight = 0.0
    foreach ($v in $weights.Values) { $totalWeight += $v }
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

    return [pscustomobject]@{ Pcts = $pcts; Amts = $amts }
}

function New-MonthlyBreakdown {
    param(
        [double]$MonthlyGross
    )

    if ($MonthlyGross -le 0) {
        return [pscustomobject]@{
            MonthlyGross      = 0.0
            GrossAllowance    = 0.0
            Basic             = 0.0
            HRA               = 0.0
            Medical           = 0.0
            Food              = 0.0
            Conveyance        = 0.0
            FestivalBonus     = 0.0
            AnnualGross       = 0.0
        }
    }

    # Monthly gross salary includes basic + HRA + fixed allowances.
    # Gross allowance = gross salary - fixed allowances.
    # Basic = gross allowance / 1.4
    # HRA = gross allowance - basic
    $medical = 250.0
    $food = 650.0
    $conveyance = 200.0
    $fixedAllowances = $medical + $food + $conveyance

    $grossSalary = [math]::Round([math]::Max(0, $MonthlyGross))
    $grossAllowance = [math]::Max(0, $grossSalary - $fixedAllowances)

    $basic = [math]::Round($grossAllowance / 1.4)
    $hra = [math]::Round($grossAllowance - $basic)

    # Keep the total monthly salary aligned after rounding.
    $current = $basic + $hra + $fixedAllowances
    $drift = [math]::Round($grossSalary - $current)
    if ($drift -ne 0) {
        $basic += $drift
    }

    $grossAllowance = [math]::Round($basic + $hra)
    $annual = [math]::Round($grossSalary * 12)

    return [pscustomobject]@{
        MonthlyGross      = $grossSalary
        GrossAllowance    = $grossAllowance
        Basic             = [math]::Round($basic)
        HRA               = [math]::Round($hra)
        Medical           = [math]::Round($medical)
        Food              = [math]::Round($food)
        Conveyance        = [math]::Round($conveyance)
        FestivalBonus     = 0.0
        AnnualGross       = $annual
    }
}

function Estimate-NetFromGross {
    param(
        [double]$BaseGross,
        [double]$Deductions,
        [string]$EmployeeType,
        $Cfg
    )

    $monthlyGross = [double]$BaseGross / 12.0
    $monthly = New-MonthlyBreakdown -MonthlyGross $monthlyGross
    $festivalBonusRate = Get-FestivalBonusRate $EmployeeType
    $festivalBonusPerEid = [math]::Round(([double]$monthly.Basic * ($festivalBonusRate / 100.0)))
    $bonus = [math]::Round($festivalBonusPerEid * 2)
    $total = [math]::Round($BaseGross + $bonus)
    $exempt = [math]::Min($total / 3.0, [double]$Cfg.SalaryExemptionCap)
    $taxable = [math]::Max(0, $total - $exempt)
    $tax = (Calculate-Tax -Taxable $taxable -Slabs $Cfg.TaxSlabs -MinimumTax ([double]$Cfg.MinimumTax)).Tax
    return $total - $tax - $Deductions
}

function Estimate-GrossFromNet {
    param(
        [double]$TargetNet,
        [double]$Deductions,
        [string]$EmployeeType,
        $Cfg
    )

    $low = 0.0
    $high = [math]::Max(1000000, $TargetNet * 3 + 1000000)
    $best = 0.0
    $bestDiff = [double]::MaxValue

    for ($i = 0; $i -lt 80; $i++) {
        $mid = ($low + $high) / 2.0
        $net = Estimate-NetFromGross -BaseGross $mid -Deductions $Deductions -EmployeeType $EmployeeType -Cfg $Cfg
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

    $salaryPeriod = (Get-Val $Inputs.SalaryPeriod 'monthly').ToLowerInvariant()
    if ($salaryPeriod -ne 'yearly') { $salaryPeriod = 'monthly' }
    $salaryAmount = Convert-HumanNumber $Inputs.TotalSalary
    $customSalary = Test-Bool $Inputs.CustomSalary $false
    $employeeType = (Get-Val $Inputs.EmployeeType 'staff').ToLowerInvariant()
    $festivalBonusRate = Get-FestivalBonusRate $employeeType
    $festivalBonusCount = 2

    $customBasic = Convert-HumanNumber $Inputs.Basic
    $customHRA = Convert-HumanNumber $Inputs.HRA
    $customMedical = Convert-HumanNumber $Inputs.Medical
    $customFood = Convert-HumanNumber $Inputs.Food
    $customTransport = Convert-HumanNumber $Inputs.Transport

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
    # If the user provides a custom breakdown, use it.
    # Otherwise, infer the full monthly payroll structure from the salary input.
    $monthly = $null
    $salaryInputAnnual = 0.0
    if ($customSalary) {
        $monthlyTotal = Round-Taka ($customBasic + $customHRA + $customMedical + $customFood + $customTransport)
        $monthly = [pscustomobject]@{
            MonthlyGross      = $monthlyTotal
            GrossAllowance    = Round-Taka ($customBasic + $customHRA)
            Basic             = Round-Taka $customBasic
            HRA               = Round-Taka $customHRA
            Medical           = Round-Taka $customMedical
            Food              = Round-Taka $customFood
            Conveyance        = Round-Taka $customTransport
            FestivalBonus     = 0.0
            AnnualGross       = [math]::Round($monthlyTotal * 12)
        }
        $salaryInputAnnual = [math]::Round($monthlyTotal * 12)
    } elseif ($salaryPeriod -eq 'yearly') {
        $salaryInputAnnual = [math]::Round($salaryAmount)
        $monthly = New-MonthlyBreakdown -MonthlyGross ([double]$salaryInputAnnual / 12.0)
    } else {
        $monthly = New-MonthlyBreakdown -MonthlyGross $salaryAmount
        $salaryInputAnnual = [math]::Round([double]$monthly.MonthlyGross * 12)
    }

    $monthlyGross = [double]$monthly.MonthlyGross
    $annualHRA = [math]::Round(([double]$monthly.HRA * 12))
    $annualMedical = [math]::Round(([double]$monthly.Medical * 12))
    $annualFood = [math]::Round(([double]$monthly.Food * 12))
    $annualConveyance = [math]::Round(([double]$monthly.Conveyance * 12))

    # Make the annual salary components add up exactly to the annual salary input.
    # The yearly basic salary is the remainder after annual allowances are removed.
    $annualCore = [math]::Round($salaryInputAnnual - $annualHRA - $annualMedical - $annualFood - $annualConveyance)

    $festivalBonusPerEid = [math]::Round(([double]$monthly.Basic * ($festivalBonusRate / 100.0)))
    $annualBonus = [math]::Round($festivalBonusPerEid * $festivalBonusCount)
    $annualGross = [math]::Round($salaryInputAnnual + $annualBonus)

    $salaryExempt = [math]::Min($annualGross / 3.0, [double]$Cfg.SalaryExemptionCap)
    $taxableSalary = [math]::Max(0, $annualGross - $salaryExempt)

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
        $estimatedGross = Estimate-GrossFromNet -TargetNet $targetNet -Deductions $deductions -EmployeeType $employeeType -Cfg $Cfg
    }

    $wealthIncrease = $netWealthUsed - $openingWealth
    $expectedSavings = $annualGross - $deductions - $totalExpense - $finalTax
    $wealthDifference = $wealthIncrease - $expectedSavings
    $tol = [math]::Max(10000, [math]::Abs($expectedSavings) * 0.01)

    if ([math]::Abs($wealthDifference) -le $tol) {
        $wealthStatus = 'OK - wealth increase matches estimated savings within tolerance.'
        $auditRisk = 'LOW'
    } elseif ($wealthDifference -gt 0) {
        $wealthStatus = 'ALERT - reported wealth is higher than estimated savings.'
        $auditRisk = 'HIGH'
    } else {
        $wealthStatus = 'Wealth increase is lower than estimated savings.'
        $auditRisk = 'MEDIUM'
    }

    return [pscustomobject]@{
        GeneratedAt = Get-Date
        TaxYear = $Cfg.TaxYear

        MonthlyGrossSalary = $monthlyGross
        MonthlyGrossAllowance = [double]$monthly.GrossAllowance
        MonthlyBasicSalary = [double]$monthly.Basic
        MonthlyHouseRentAllowance = [double]$monthly.HRA
        MonthlyMedicalAllowance = [double]$monthly.Medical
        MonthlyFoodAllowance = [double]$monthly.Food
        MonthlyConveyanceAllowance = [double]$monthly.Conveyance
        SalaryPeriod = $salaryPeriod

        SalaryInput = $salaryInputAnnual
        CustomSalary = $customSalary
        AnnualBasicSalary = $annualCore
        BaseGrossSalary = $annualCore
        HouseRentAllowance = $annualHRA
        MedicalAllowance = $annualMedical
        FoodAllowance = $annualFood
        ConveyanceAllowance = $annualConveyance
        FestivalBonusPerEid = $festivalBonusPerEid
        BonusAmount = $annualBonus
        TotalComp = $annualGross
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
    return (Get-ExpenseKeys)
}

function Bar-ChartLine {
    param(
        [string]$Label,
        [double]$Part,
        [double]$Total,
        [int]$Width = 26
    )

    if ($Total -le 0) { return $Label.PadRight(28) + ' | ' }
    $filled = [int][math]::Round(($Part / $Total) * $Width)
    if ($filled -lt 0) { $filled = 0 }
    if ($filled -gt $Width) { $filled = $Width }
    $bar = ('#' * $filled).PadRight($Width, '.')
    $percent = [math]::Round((($Part / $Total) * 100), 1)
    $percentText = $percent.ToString('0.0')
    return $Label.PadRight(28) + ' | ' + $bar + ' | ' + $percentText.PadLeft(6) + '%'
}

function Format-TaxTable {
    param($Lines)
    if (-not $Lines -or $Lines.Count -eq 0) { return "No taxable amount.`n" }

    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine('SLAB'.PadRight(34) + ' | ' + 'TAXED'.PadRight(12) + ' | ' + 'RATE'.PadRight(8) + ' | ' + 'TAX'.PadRight(12))
    [void]$out.AppendLine((''.PadRight(78, '-')))
    foreach ($ln in $Lines) {
        $rateText = ([math]::Round(($ln.Rate * 100), 0)).ToString('0') + '%'
        [void]$out.AppendLine(
            $ln.Label.PadRight(34) + ' | ' +
            (Format-Money $ln.Amount).PadLeft(12) + ' | ' +
            $rateText.PadLeft(7) + ' | ' +
            (Format-Money $ln.Tax).PadLeft(12)
        )
    }
    [void]$out.AppendLine((''.PadRight(78, '-')))
    return $out.ToString()
}

function Format-KvLine {
    param([string]$Label, [double]$Value)
    return $Label.PadRight(34) + ' Tk ' + (Format-Money $Value)
}

function New-TextReport {
    param($R)

    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine('================================================================================')
    [void]$sb.AppendLine(" $($script:AppName)  $($script:AppVersion)")
    [void]$sb.AppendLine(" Author: $($script:AppAuthor)")
    [void]$sb.AppendLine(" Repo  : $($script:AppRepo)")
    [void]$sb.AppendLine('SALARY / INCOME (MONTHLY)')
    [void]$sb.AppendLine(('Salary input mode'.PadRight(34) + ' ' + $R.SalaryPeriod))
    [void]$sb.AppendLine((Format-KvLine 'Monthly gross salary' $R.MonthlyGrossSalary))
    [void]$sb.AppendLine((Format-KvLine 'Gross allowance' $R.MonthlyGrossAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Basic salary' $R.MonthlyBasicSalary))
    [void]$sb.AppendLine((Format-KvLine 'House rent allowance' $R.MonthlyHouseRentAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Medical allowance' $R.MonthlyMedicalAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Food allowance' $R.MonthlyFoodAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Conveyance allowance' $R.MonthlyConveyanceAllowance))
    [void]$sb.AppendLine()

    [void]$sb.AppendLine('SALARY / INCOME (ANNUALIZED)')
    [void]$sb.AppendLine((Format-KvLine 'Annual salary input' $R.SalaryInput))
    [void]$sb.AppendLine((Format-KvLine 'Annual basic salary' $R.BaseGrossSalary))
    [void]$sb.AppendLine((Format-KvLine 'Annual house rent allowance' $R.HouseRentAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Annual medical allowance' $R.MedicalAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Annual food allowance' $R.FoodAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Annual conveyance allowance' $R.ConveyanceAllowance))
    [void]$sb.AppendLine((Format-KvLine 'Festival bonus per Eid' $R.FestivalBonusPerEid))
    [void]$sb.AppendLine((Format-KvLine 'Annual festival bonus (2 Eid)' $R.BonusAmount))
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

    [void]$sb.AppendLine('VISUAL SUMMARY')
    [void]$sb.AppendLine((Bar-ChartLine -Label 'Tax before rebate' -Part $R.CurrentTaxBeforeRebate -Total ([math]::Max(1, $R.CurrentTaxBeforeRebate + $R.SurchargeAmount)) -Width 36))
    if ($R.SurchargeAmount -gt 0) {
        [void]$sb.AppendLine((Bar-ChartLine -Label 'Surcharge' -Part $R.SurchargeAmount -Total ([math]::Max(1, $R.CurrentTaxBeforeRebate + $R.SurchargeAmount)) -Width 36))
    }
    if ($R.TotalExpense -gt 0) {
        [void]$sb.AppendLine()
        foreach ($k in (Get-ExpenseOrder)) {
            if ($R.ExpenseAmts[$k] -gt 0) {
                [void]$sb.AppendLine((Bar-ChartLine -Label $k -Part ([double]$R.ExpenseAmts[$k]) -Total $R.TotalExpense -Width 36))
            }
        }
    }
    [void]$sb.AppendLine()

    [void]$sb.AppendLine('TIP: use export commands for JSON, CSV, Markdown, or PDF.')
    return $sb.ToString()
}

function Save-Session {
    param(
        [string]$Path,
        $Inputs,
        [int]$Step = 0,
        [int]$Section = 0
    )

    $session = [pscustomobject]@{
        Version   = 1
        Timestamp = (Get-Date).ToString('o')
        Inputs    = $Inputs
        Step      = $Step
        Section   = $Section
    }

    $raw = $session | ConvertTo-Json -Depth 12
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hashBytes = $sha.ComputeHash($bytes)
        $hash = ([BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
        $wrapped = [pscustomobject]@{
            Hash    = $hash
            Session = $session
        }
        $wrapped | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
    } finally {
        $sha.Dispose()
    }
}

function Load-Session {
    param([string]$Path)

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
        return $wrapped.Session
    } finally {
        $sha.Dispose()
    }
}

function Export-CSVReport {
    param([string]$Path, $R)

    $rows = @(
        [pscustomobject]@{ Field='Tax year'; Value=$R.TaxYear }
        [pscustomobject]@{ Field='Monthly gross salary'; Value=(Format-Money $R.MonthlyGrossSalary) }
        [pscustomobject]@{ Field='Monthly gross allowance'; Value=(Format-Money $R.MonthlyGrossAllowance) }
        [pscustomobject]@{ Field='Monthly basic salary'; Value=(Format-Money $R.MonthlyBasicSalary) }
        [pscustomobject]@{ Field='Monthly house rent allowance'; Value=(Format-Money $R.MonthlyHouseRentAllowance) }
        [pscustomobject]@{ Field='Monthly medical allowance'; Value=(Format-Money $R.MonthlyMedicalAllowance) }
        [pscustomobject]@{ Field='Monthly food allowance'; Value=(Format-Money $R.MonthlyFoodAllowance) }
        [pscustomobject]@{ Field='Monthly conveyance allowance'; Value=(Format-Money $R.MonthlyConveyanceAllowance) }
        [pscustomobject]@{ Field='Annual salary input'; Value=(Format-Money $R.SalaryInput) }
        [pscustomobject]@{ Field='Annual basic salary'; Value=(Format-Money $R.BaseGrossSalary) }
        [pscustomobject]@{ Field='House rent allowance'; Value=(Format-Money $R.HouseRentAllowance) }
        [pscustomobject]@{ Field='Medical allowance'; Value=(Format-Money $R.MedicalAllowance) }
        [pscustomobject]@{ Field='Food allowance'; Value=(Format-Money $R.FoodAllowance) }
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
    param([string]$Path, $R)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# IT-10BB Tax Companion Report')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("* Author: $($script:AppAuthor)")
    [void]$sb.AppendLine("* Version: $($script:AppVersion)")
    [void]$sb.AppendLine("* Repo: $($script:AppRepo)")
    [void]$sb.AppendLine("* Tax year: $($R.TaxYear)")
    [void]$sb.AppendLine("* Generated: $($R.GeneratedAt.ToString('o'))")
    [void]$sb.AppendLine()

    [void]$sb.AppendLine('## Monthly salary breakdown')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('| Field | Value |')
    [void]$sb.AppendLine('|---|---:|')
    foreach ($pair in @(
        @('Monthly gross salary', $R.MonthlyGrossSalary)
        @('Monthly gross allowance', $R.MonthlyGrossAllowance)
        @('Monthly basic salary', $R.MonthlyBasicSalary)
        @('Monthly house rent allowance', $R.MonthlyHouseRentAllowance)
        @('Monthly medical allowance', $R.MonthlyMedicalAllowance)
        @('Monthly food allowance', $R.MonthlyFoodAllowance)
        @('Monthly conveyance allowance', $R.MonthlyConveyanceAllowance)
    )) {
        [void]$sb.AppendLine("| $($pair[0]) | Tk $(Format-Money $pair[1]) |")
    }

    [void]$sb.AppendLine()
    [void]$sb.AppendLine('## Annualized salary / tax summary')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('| Field | Value |')
    [void]$sb.AppendLine('|---|---:|')
    foreach ($pair in @(
        @('Annual salary input', $R.SalaryInput)
        @('Annual basic salary', $R.BaseGrossSalary)
        @('House rent allowance', $R.HouseRentAllowance)
        @('Medical allowance', $R.MedicalAllowance)
        @('Food allowance', $R.FoodAllowance)
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
        [void]$sb.AppendLine("| $($pair[0]) | Tk $([string]$(if ($pair[0] -eq 'Wealth status' -or $pair[0] -eq 'Audit risk') { $pair[1] } else { (Format-Money $pair[1]) })) |")
    }

    if ($R.TotalExpense -gt 0) {
        [void]$sb.AppendLine()
        [void]$sb.AppendLine('## Expense allocation')
        [void]$sb.AppendLine()
        [void]$sb.AppendLine('| Category | % | Amount |')
        [void]$sb.AppendLine('|---|---:|---:|')
        foreach ($k in (Get-ExpenseOrder)) {
            if ($R.ExpensePcts[$k] -eq 0 -and $R.ExpenseAmts[$k] -eq 0) { continue }
            [void]$sb.AppendLine("| $k | $($R.ExpensePcts[$k]) | Tk $(Format-Money $R.ExpenseAmts[$k]) |")
        }
    }

    if ($R.ReverseEnabled -and $R.TargetNetTakeHome -gt 0) {
        [void]$sb.AppendLine()
        [void]$sb.AppendLine('## Reverse calculator')
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("* Target net take-home: Tk $(Format-Money $R.TargetNetTakeHome)")
        [void]$sb.AppendLine("* Extra deductions: Tk $(Format-Money $R.Deductions)")
        [void]$sb.AppendLine("* Estimated gross salary: Tk $(Format-Money $R.EstimatedGrossFromNet)")
    }

    [IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

function Escape-PdfText {
    param([string]$Text)
    $s = $Text
    if ($null -eq $s) { $s = '' }
    $s = $s -replace '\\', '\\\\'
    $s = $s -replace '\(', '\('
    $s = $s -replace '\)', '\)'
    return $s
}

function Export-PdfReport {
    param([string]$Path, $R)

    $text = New-TextReport $R
    $lines = $text -split "`r?`n"

    $pageWidth = 595.0
    $pageHeight = 842.0
    $marginX = 40.0
    $marginTop = 40.0
    $fontSize = 9.0
    $lineHeight = 11.0
    $maxLinesPerPage = [int][math]::Floor(($pageHeight - (2 * $marginTop)) / $lineHeight)
    if ($maxLinesPerPage -lt 1) { $maxLinesPerPage = 1 }

    $pages = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lines.Count; $i += $maxLinesPerPage) {
        $end = [math]::Min($i + $maxLinesPerPage, $lines.Count)
        $pages.Add($lines[$i..($end - 1)])
    }

    $objects = New-Object System.Collections.Generic.List[string]
    $objects.Add('') # 1 Catalog placeholder
    $objects.Add('') # 2 Pages placeholder
    $objects.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>') # 3 font

    $pageObjectNumbers = New-Object System.Collections.Generic.List[int]
    $contentObjectNumbers = New-Object System.Collections.Generic.List[int]

    for ($p = 0; $p -lt $pages.Count; $p++) {
        $pageObjectNumbers.Add($objects.Count + 1)
        $objects.Add('') # page placeholder
        $contentObjectNumbers.Add($objects.Count + 1)
        $objects.Add('') # content placeholder
    }

    $kidsRefs = ($pageObjectNumbers | ForEach-Object { "$_ 0 R" }) -join ' '

    # Fill page and content objects
    for ($p = 0; $p -lt $pages.Count; $p++) {
        $pageNum = $pageObjectNumbers[$p]
        $contentNum = $contentObjectNumbers[$p]
        $pageObj = " << /Type /Page /Parent 2 0 R /MediaBox [0 0 $pageWidth $pageHeight] /Resources << /Font << /F1 3 0 R >> >> /Contents $contentNum 0 R >>"
        $objects[$pageNum - 1] = $pageObj

        $contentLines = New-Object System.Collections.Generic.List[string]
        $contentLines.Add('BT')
        $contentLines.Add("/F1 $fontSize Tf")
        $contentLines.Add("1 0 0 1 $marginX " + ($pageHeight - $marginTop) + " Tm")

        foreach ($line in $pages[$p]) {
            $escaped = Escape-PdfText $line
            $contentLines.Add("($escaped) Tj")
            $contentLines.Add("0 -$lineHeight Td")
        }
        $contentLines.Add('ET')
        $contentStream = ($contentLines -join "`n")
        $objects[$contentNum - 1] = "<< /Length $([Text.Encoding]::ASCII.GetByteCount($contentStream)) >>`nstream`n$contentStream`nendstream"
    }

    $objects[0] = "<< /Type /Catalog /Pages 2 0 R >>"
    $objects[1] = "<< /Type /Pages /Kids [$kidsRefs] /Count $($pages.Count) >>"

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('%PDF-1.4')
    $offsets = New-Object System.Collections.Generic.List[int]
    $offsets.Add(0)

    for ($i = 0; $i -lt $objects.Count; $i++) {
        $offsets.Add([Text.Encoding]::ASCII.GetByteCount($sb.ToString()))
        [void]$sb.AppendLine("$($i + 1) 0 obj")
        [void]$sb.AppendLine($objects[$i])
        [void]$sb.AppendLine('endobj')
    }

    $xrefStart = [Text.Encoding]::ASCII.GetByteCount($sb.ToString())
    [void]$sb.AppendLine('xref')
    [void]$sb.AppendLine("0 $($objects.Count + 1)")
    [void]$sb.AppendLine('0000000000 65535 f ')
    for ($i = 1; $i -le $objects.Count; $i++) {
        [void]$sb.AppendLine(('{0:0000000000} 00000 n ' -f $offsets[$i]))
    }
    [void]$sb.AppendLine('trailer')
    [void]$sb.AppendLine("<< /Size $($objects.Count + 1) /Root 1 0 R >>")
    [void]$sb.AppendLine('startxref')
    [void]$sb.AppendLine($xrefStart)
    [void]$sb.AppendLine('%%EOF')

    [IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::ASCII)
}

function Write-Banner {
    Clear-Host
    Write-Host '+==============================================================================+' -ForegroundColor Cyan
    Write-Host ("|  {0,-74}|" -f $script:AppName) -ForegroundColor Cyan
    Write-Host ("|  {0,-74}|" -f "$($script:AppVersion)  |  Author: $($script:AppAuthor)") -ForegroundColor Green
    Write-Host ("|  {0,-74}|" -f "Repo: $($script:AppRepo)") -ForegroundColor DarkCyan
    Write-Host '+==============================================================================+' -ForegroundColor Cyan
    Write-Host
}

function Read-Field {
    param(
        [string]$Prompt,
        [string]$Default
    )

    $value = Read-Host ("{0} [{1}]" -f $Prompt, $Default)
    if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
    return $value
}

function Prompt-Inputs {
    $inputs = [ordered]@{
        SalaryPeriod     = 'monthly'
        TotalSalary      = ''
        CustomSalary     = 'n'
        Basic            = '0'
        HRA              = '0'
        Medical          = '0'
        Food             = '0'
        Transport        = '0'
        EmployeeType     = 'staff'
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

    Write-Banner
    Write-Host 'Use expressions like 12809*23, 1 lakh, 2cr, 4.5k, and 50%.' -ForegroundColor Yellow
    Write-Host 'Festival bonus is handled separately: staff gets 100% of basic salary per Eid, workers get 80%.' -ForegroundColor DarkGray
    Write-Host 'If you choose custom breakdown, the gross-salary input is ignored.' -ForegroundColor DarkGray
    Write-Host 'Press Enter to accept the default in brackets.' -ForegroundColor DarkGray
    Write-Host

    $inputs.SalaryPeriod = Read-Field '1. Salary input mode (monthly/yearly)' $inputs.SalaryPeriod
    $inputs.TotalSalary   = Read-Field '2. Gross salary amount (BDT)' $inputs.TotalSalary
    $inputs.CustomSalary  = Read-Field '3. Enter custom salary breakdown? (y/n)' $inputs.CustomSalary
    $inputs.EmployeeType  = Read-Field '4. Employee type (staff/worker)' $inputs.EmployeeType

    if (Test-Bool $inputs.CustomSalary $false) {
        Write-Host 'Custom breakdown is enabled: the gross salary input above will be ignored.' -ForegroundColor DarkGray
        $inputs.Basic       = Read-Field '   -> Basic salary (monthly BDT)' $inputs.Basic
        $inputs.HRA         = Read-Field '   -> House rent allowance (monthly BDT)' $inputs.HRA
        $inputs.Medical     = Read-Field '   -> Medical allowance (monthly BDT)' $inputs.Medical
        $inputs.Food        = Read-Field '   -> Food allowance (monthly BDT)' $inputs.Food
        $inputs.Transport   = Read-Field '   -> Conveyance allowance (monthly BDT)' $inputs.Transport
    }

    $inputs.TotalExpense     = Read-Field '4. Total annual expense (BDT)' $inputs.TotalExpense
    $inputs.Location         = Read-Field '5. Location (dhaka/other)' $inputs.Location
    $inputs.FamilySize       = Read-Field '6. Family size' $inputs.FamilySize
    $inputs.HasKids          = Read-Field '7. Do you have kids? (y/n)' $inputs.HasKids
    $inputs.OwnHome          = Read-Field '8. Do you own your home? (y/n)' $inputs.OwnHome
    $inputs.HasStaff         = Read-Field '9. Home-support staff? (y/n)' $inputs.HasStaff
    $inputs.Mode             = Read-Field '10. Mode (balanced/conservative/comfortable)' $inputs.Mode
    $inputs.PreviousGross    = Read-Field "11. Previous year's gross income (annual BDT)" $inputs.PreviousGross
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

    return $inputs
}

function Show-Report {
    param($R)

    Write-Banner

    Write-Host 'MONTHLY SALARY BREAKDOWN' -ForegroundColor Magenta
    Write-Host ('{0,-34} {1}' -f 'Salary input mode', $R.SalaryPeriod) -ForegroundColor DarkGray
    Write-Host ('{0,-34} Tk {1}' -f 'Monthly gross salary', (Format-Money $R.MonthlyGrossSalary)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Gross allowance', (Format-Money $R.MonthlyGrossAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Basic salary', (Format-Money $R.MonthlyBasicSalary)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'House rent allowance', (Format-Money $R.MonthlyHouseRentAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Medical allowance', (Format-Money $R.MonthlyMedicalAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Food allowance', (Format-Money $R.MonthlyFoodAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Conveyance allowance', (Format-Money $R.MonthlyConveyanceAllowance)) -ForegroundColor White
    Write-Host

    Write-Host 'ANNUALIZED SALARY / TAX' -ForegroundColor Cyan
    Write-Host ('{0,-34} Tk {1}' -f 'Annual salary input', (Format-Money $R.SalaryInput)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Annual basic salary', (Format-Money $R.BaseGrossSalary)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'House rent allowance', (Format-Money $R.HouseRentAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Medical allowance', (Format-Money $R.MedicalAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Food allowance', (Format-Money $R.FoodAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Conveyance allowance', (Format-Money $R.ConveyanceAllowance)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Festival bonus per Eid', (Format-Money $R.FestivalBonusPerEid)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Festival bonus (2 Eid)', (Format-Money $R.BonusAmount)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Total compensation', (Format-Money $R.TotalComp)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Salary exemption', (Format-Money $R.SalaryExempt)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Taxable income', (Format-Money $R.TaxableSalary)) -ForegroundColor White
    if ($R.ApplySurcharge) {
        Write-Host 'NET WEALTH SURCHARGE' -ForegroundColor Red
        Write-Host ('{0,-34} {1}' -f 'Applied', ($(if ($R.SurchargeRate -gt 0) { 'Yes' } else { 'No' }))) -ForegroundColor White
        Write-Host ('{0,-34} {1}' -f 'Surcharge rate', ('{0:N2}%' -f ($R.SurchargeRate * 100))) -ForegroundColor White
        Write-Host ('{0,-34} Tk {1}' -f 'Surcharge amount', (Format-Money $R.SurchargeAmount)) -ForegroundColor White
        Write-Host ('{0,-34} Tk {1}' -f 'Combined tax before surcharge', (Format-Money $R.CombinedBeforeSurch)) -ForegroundColor White
        Write-Host ('{0,-34} Tk {1}' -f 'Final tax', (Format-Money $R.FinalTax)) -ForegroundColor White
        Write-Host
    } else {
        Write-Host 'FINAL TAX' -ForegroundColor Green
        Write-Host ('{0,-34} Tk {1}' -f 'Final tax', (Format-Money $R.FinalTax)) -ForegroundColor White
        Write-Host
    }

    if ($R.TotalExpense -gt 0) {
        Write-Host 'IT-10BB EXPENSE ALLOCATION' -ForegroundColor Magenta
        foreach ($k in (Get-ExpenseOrder)) {
            if ($R.ExpensePcts[$k] -eq 0 -and $R.ExpenseAmts[$k] -eq 0) { continue }
            Write-Host ('{0,-34} {1,3}%   Tk {2}' -f $k, $R.ExpensePcts[$k], (Format-Money $R.ExpenseAmts[$k])) -ForegroundColor White
        }
        Write-Host ('{0,-34} {1,3}%   Tk {2}' -f 'TOTAL', 100, (Format-Money ([int](Round-Taka $R.TotalExpense)))) -ForegroundColor White
        Write-Host
    }

    Write-Host 'WEALTH SUMMARY' -ForegroundColor Cyan
    Write-Host ('{0,-34} Tk {1}' -f 'Total assets (provided)', (Format-Money $R.TotalAssets)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Total liabilities', (Format-Money $R.TotalLiabilities)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Net wealth (used)', (Format-Money $R.NetWealthCurrent)) -ForegroundColor White
    Write-Host ('{0,-34} {1}' -f 'Audit risk', $R.AuditRisk) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Opening net wealth', (Format-Money $R.OpeningNetWealth)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Wealth increase', (Format-Money $R.WealthIncrease)) -ForegroundColor White
    Write-Host ('{0,-34} Tk {1}' -f 'Estimated after-tax savings', (Format-Money $R.ExpectedSavings)) -ForegroundColor White
    Write-Host $R.WealthStatus -ForegroundColor Yellow
    Write-Host

    if ($R.ReverseEnabled -and $R.TargetNetTakeHome -gt 0) {
        Write-Host 'REVERSE CALCULATOR' -ForegroundColor DarkCyan
        Write-Host ('{0,-34} Tk {1}' -f 'Target net take-home', (Format-Money $R.TargetNetTakeHome)) -ForegroundColor White
        Write-Host ('{0,-34} Tk {1}' -f 'Extra deductions', (Format-Money $R.Deductions)) -ForegroundColor White
        Write-Host ('{0,-34} Tk {1}' -f 'Estimated gross salary', (Format-Money $R.EstimatedGrossFromNet)) -ForegroundColor White
        Write-Host
    }

    Write-Host 'VISUAL SUMMARY' -ForegroundColor Green
    Write-Host (Bar-ChartLine -Label 'Tax before rebate' -Part $R.CurrentTaxBeforeRebate -Total ([math]::Max(1, $R.CurrentTaxBeforeRebate + $R.SurchargeAmount)) -Width 36) -ForegroundColor White
    if ($R.SurchargeAmount -gt 0) {
        Write-Host (Bar-ChartLine -Label 'Surcharge' -Part $R.SurchargeAmount -Total ([math]::Max(1, $R.CurrentTaxBeforeRebate + $R.SurchargeAmount)) -Width 36) -ForegroundColor White
    }
    if ($R.TotalExpense -gt 0) {
        foreach ($k in (Get-ExpenseOrder)) {
            if ($R.ExpenseAmts[$k] -gt 0) {
                Write-Host (Bar-ChartLine -Label $k -Part ([double]$R.ExpenseAmts[$k]) -Total $R.TotalExpense -Width 36) -ForegroundColor White
            }
        }
    }

    Write-Host
    Write-Host 'Commands: [s]ave JSON  [l]oad JSON  [c]sv  [m]d  [p]df  [r]e-run  [q]uit' -ForegroundColor DarkGray
}

function Invoke-App {
    $cfg = Load-Config
    while ($true) {
        $inputs = Prompt-Inputs
        $report = Build-Report -Inputs $inputs -Cfg $cfg
        Show-Report $report

        while ($true) {
            $cmd = (Read-Host 'Enter command').Trim().ToLowerInvariant()
            switch ($cmd) {
                's' {
                    $path = Read-Host 'Session JSON path'
                    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_session.json' }
                    if (-not $path.EndsWith('.json')) { $path += '.json' }
                    Save-Session -Path $path -Inputs $inputs -Step 0 -Section 0
                    Write-Host "Saved: $path" -ForegroundColor Green
                }
                'l' {
                    $path = Read-Host 'Session JSON path'
                    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_session.json' }
                    if (-not $path.EndsWith('.json')) { $path += '.json' }
                    $session = Load-Session -Path $path
                    $loaded = $session.Inputs
                    foreach ($k in $inputs.Keys) {
                        if ($loaded.PSObject.Properties[$k]) {
                            $inputs[$k] = [string]$loaded.$k
                        }
                    }
                    $report = Build-Report -Inputs $inputs -Cfg $cfg
                    Show-Report $report
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
                'p' {
                    $path = Read-Host 'PDF path'
                    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'tax_summary.pdf' }
                    if (-not $path.EndsWith('.pdf')) { $path += '.pdf' }
                    Export-PdfReport -Path $path -R $report
                    Write-Host "PDF exported: $path" -ForegroundColor Green
                }
                'r' {
                    break
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
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-App
}
