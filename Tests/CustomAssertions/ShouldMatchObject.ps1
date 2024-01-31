#Requires -Module Pester
function Get-Properties {
    param(
        [object]
        $Object
    )

    Get-Member -InputObject $Object | Where-Object {
        $_.MemberType -ne [System.Management.Automation.PSMemberTypes]::Method
    }
}

function Test-Object {
    param(
        [object]
        $ActualValue,

        [object]
        $ExpectedValue
    )

    if ($ActualValue.pstypenames[0] -ne $ExpectedValue.pstypenames[0]) {
        return [PSCustomObject]@{
            Succeeded      = $false
            FailureMessage = "Expected an object with a type name of $($ExpectedValue.pstypenames[0]), but got $($ActualValue.pstypenames[0])."
        }
    }

    $actualProperties = (Get-Properties $ActualValue).Name
    $expectedProperties = (Get-Properties $ExpectedValue).Name

    $actualNotInExpected = $actualProperties.Where({ $_ -notin $expectedProperties })
    $expectedNotInActual = $expectedProperties.Where({ $_ -notin $actualProperties })

    if ($actualNotInExpected.Length -gt 0 -and $expectedNotInActual.Length -gt 0) {
        return [PSCustomObject]@{
            Succeeded      = $false
            FailureMessage = @(
                "Expected an object with @($($expectedProperties -join ', ')) properties, but got @($($actualProperties -join ', '))."
                "Received the following extra properties: $($actualNotInExpected -join ', ')"
                "Missing the following properties: $($expectedNotInActual -join ', ')"
            ) -join "`n"
        }
    } elseif ($actualNotInExpected.Length -gt 0) {
        return [PSCustomObject]@{
            Succeeded      = $false
            FailureMessage = @(
                "Expected an object with @($($expectedProperties -join ', ')) properties, but got @($($actualProperties -join ', '))."
                "Received the following extra properties: $($actualNotInExpected -join ', ')"
            ) -join "`n"
        }
    } elseif ($expectedNotInActual.Length -gt 0) {
        return [PSCustomObject]@{
            Succeeded      = $false
            FailureMessage = @(
                "Expected an object with @($($expectedProperties -join ', ')) properties, but got @($($actualProperties -join ', '))."
                "Missing the following properties: $($expectedNotInActual -join ', ')"
            ) -join "`n"
        }
    }

    foreach ($name in $actualProperties) {
        if ($ActualValue.$name -ne $ExpectedValue.$name) {
            return [PSCustomObject]@{
                Succeeded      = $false
                FailureMessage = "Expected an object with a $name property equal to $($ExpectedValue.$name), but got $($ActualValue.$name)."
            }
        }
    }

    return [PSCustomObject]@{
        Succeeded      = $true
        FailureMessage = $null
    }
}

function Test-Objects {
    param(
        [object]
        $ActualValue,

        [object]
        $ExpectedValue
    )

    if (
        ($ActualValue -is [array]) -xor ($ExpectedValue -is [array])
    ) {
        return [PSCustomObject]@{
            Succeeded      = $false
            FailureMessage = "Expected an object of type $($ExpectedValue.GetType().Name), but got $($ActualValue.GetType().Name)."
        }
    } elseif ($ActualValue -is [array]) {
        if ($ActualValue.Length -ne $ExpectedValue.Length) {
            return [PSCustomObject]@{
                Succeeded      = $false
                FailureMessage = "Expected an array of length $($ExpectedValue.Length), but got $($ActualValue.Length)."
            }
        }

        for ($i = 0; $i -lt $ActualValue.Length; $i++) {
            $result = Test-Object -ActualValue $ActualValue[$i] -ExpectedValue $ExpectedValue[$i]

            if (-not $result.Succeeded) {
                return $result
            }
        }

        return [PSCustomObject]@{
            Succeeded      = $true
            FailureMessage = $null
        }
    }

    return Test-Object -ActualValue $ActuallValue -ExpectedValue $ExpectedValue
}

Add-ShouldOperator -Name MatchObject -InternalName Test-Objects -Test ${function:Test-Objects} -SupportsArrayInput
