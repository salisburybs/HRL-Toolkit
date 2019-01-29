#Import-Module $PSScriptRoot\Logging_Functions.ps1

Function Persona-GetSQLLocation{
    return "HRL-DBS02.domain.hrl\AMT2"
}

Function Persona-GetDB{
    return "OFS"
}

Function Validate-Pin{
    param([string]$pinCode)
    $badPins = "000000", "111111", "222222", "333333", "444444",`
        "555555", "666666", "777777", "888888", "999999", "123456", "654321"
    
    if($pinCode -in $badPins){
        return $false
    }
    return $true
}

Function Get-RandomPin{
    $pin = ""
    Do {
        $pin = (Get-Random -Minimum 000000 -Maximum 999999).ToString("000000")
    } Until (Validate-Pin -pinCode $pin)
    
    return $pin
}

Function Get-ExpiredPinOnlyCredentials {

    $SQLServer = Persona-GetSQLLocation
    $SQLDBName = Persona-GetDB
    $SQLQuery = "
        SELECT
	        1 as UserID,
	        idh.LastName,
	        idh.FirstName,
	        idh.AMTIDHolder,
	        crd.CardID,
	        ccs.CardSetID,
	        crd.EncodedID,
	        crd.EmbossedID,
	        crd.CredentialType,
	        crd.PINCode,
	        crd.InputSuppression,
	        crd.ExtendedAccess,
	        crd.PassbackExempt,
	        crd.PINCommandAccess,
	        crd.ActivationDateTime,
	        crd.ExpirationDateTime,
	        crd.EscortCardID,
	        crd.Deleted,
	        crd.RawBitPattern,
	        crs.DefaultUserType
        FROM [OFS].[VertXBase].[Card] crd
        Join AmtOfBase.AMTIDs amt on 
	        amt.EncodedID=crd.EncodedID 
        Join Custom.IDHolders idh on 
	        idh.AMTIDHolder=amt.AMTIDHolder
        left join Application.VertXCardToAHG420CredentialSet vrtx on
	        vrtx.VertXAMTID = crd.CardID
        left join AHG420Base.CredentialSets crs on
	        crs.AMTID = vrtx.AHG420AMTID
        left join [OFS].[VertXBase].[CardsToCardSet] ccs on
	        ccs.CardID = crd.CardID
        where 
	        crd.CredentialType=3 and   --PIN only credential 
	        DefaultUserType = 4 and
            crd.ExpirationDateTime < GETDATE() and 
	        crd.CardID not in (93562, 199404) -- lapoint and admission
        ORDER BY
	        idh.LastName, idh.FirstName
    "
    # Much easier. Assumes Windows Integrated Sec.
    # Returns array without the count as the first object
    return @(Invoke-Sqlcmd -Query $SQLQuery -ServerInstance $SQLServer)
}

Function PCO-UpdateCardAndCredentialSet{
param([int]$UserID,
    [string]$AMTIDHolder,
    [int]$CardID,
    [int]$CardSetID,
    [string]$EncodedID = '',
    [string]$EmbossedID = '',
    [int]$CredentialType,
    [string]$PINCode = '',
    [int]$InputSuppression = 0,
    [int]$ExtendedAccess = 0,
    [int]$PassbackExempt = 0,
    [int]$PINCommandAccess = 0,
    [DateTime]$ActivationDateTime = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
    [DateTime]$ExpirationDateTime,
    [int]$EscortCardID = 0,
    [int]$Deleted = 0,
    [string]$RawBitPattern = '',
    [int]$ReUseEncodedID = 0,
    [int]$ShowDebugInfo = 0,
    [int]$TranslateAllCards = 0,
    [int]$DefaultCached = 0,
    [int]$DefaultUserType = 4,
    [string]$logname)

    $SqlServer = Persona-GetSQLLocation
    $SqlDBName = Persona-GetDB

    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$SqlServer;Database=$SqlDBName;Integrated Security=True"

    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandType = [System.Data.CommandType]::StoredProcedure
    $SqlCmd.Connection = $SqlConn
    $SqlCmd.CommandText = "[Application].[AddUpdateCardAndCredentialSet]"


    $SqlCmd.Parameters.Add("@UserID", $UserID) | Out-Null
    $SqlCmd.Parameters.Add("@AMTIDHolder", $AMTIDHolder) | Out-Null

    $param_CardID = New-Object System.Data.SqlClient.SqlParameter
    $param_CardID.Direction = [System.Data.ParameterDirection]::InputOutput
    $param_CardID.ParameterName = "@CardID"
    $param_CardID.Value = $CardID
    $SqlCmd.Parameters.Add($param_CardID) | Out-Null

    $SqlCmd.Parameters.Add("@CardSetID", $CardSetID) | Out-Null
    $SqlCmd.Parameters.Add("@EncodedID", $EncodedID) | Out-Null
    $SqlCmd.Parameters.Add("@EmbossedID", $EmbossedID) | Out-Null
    $SqlCmd.Parameters.Add("@CredentialType", $CredentialType) | Out-Null
    $SqlCmd.Parameters.Add("@PINCode", $PINCode) | Out-Null
    $SqlCmd.Parameters.Add("@InputSuppression", $InputSuppression) | Out-Null
    $SqlCmd.Parameters.Add("@ExtendedAccess", $ExtendedAccess) | Out-Null
    $SqlCmd.Parameters.Add("@PassbackExempt", $PassbackExempt) | Out-Null
    $SqlCmd.Parameters.Add("@PINCommandAccess", $PINCommandAccess) | Out-Null
    $SqlCmd.Parameters.Add("@ActivationDateTime", $ActivationDateTime) | Out-Null
    $SqlCmd.Parameters.Add("@ExpirationDateTime", $ExpirationDateTime) | Out-Null
    $SqlCmd.Parameters.Add("@EscortCardID", $EscortCardID) | Out-Null
    $SqlCmd.Parameters.Add("@Deleted", $Deleted) | Out-Null
    $SqlCmd.Parameters.Add("@RawBitPattern", $RawBitPattern) | Out-Null
    
    $param_ReUseEncodedID = New-Object System.Data.SqlClient.SqlParameter
    $param_ReUseEncodedID.Direction = [System.Data.ParameterDirection]::Output
    $param_ReUseEncodedID.ParameterName = "@ReUseEncodedID"
    $param_ReUseEncodedID.Value = 0
    $SqlCmd.Parameters.Add($param_ReUseEncodedID) | Out-Null

    $param_ErrorText = New-Object System.Data.SqlClient.SqlParameter
    $param_ErrorText.Direction = [System.Data.ParameterDirection]::Output
    $param_ErrorText.ParameterName = "@ErrorText"
    $param_ErrorText.Value = ""
    $SqlCmd.Parameters.Add($param_ErrorText) | Out-Null
    
    $SqlCmd.Parameters.Add("@ShowDebugInfo", $ShowDebugInfo) | Out-Null
    $SqlCmd.Parameters.Add("@TranslateAllCards", $TranslateAllCards) | Out-Null
    $SqlCmd.Parameters.Add("@DefaultCached", $DefaultCached) | Out-Null
    $SqlCmd.Parameters.Add("@DefaultUserType", $DefaultUserType) | Out-Null

    try{
        $SqlConn.Open()
        $result = $SqlCmd.ExecuteNonQuery()
        #Write-Host "INFO Sucessfully created/updated credential"
        if($result -ne -1){
            @($result, $param_CardID, $param_ReUseEncodedID, $param_ErrorText)
        }
        $SqlConn.Close()
        return $result
    } catch {
        if($SqlConn.State -eq "Open"){
            $SqlConn.Close()
        }
        Write-Host "ERROR CardID=$CardID"
        return $null
    }
}

Function PCO-UpdatePins{
    $output = @()

    foreach($card in Get-ExpiredPinOnlyCredentials){ 
        $newpin = Get-RandomPin
        $updatecard = PCO-UpdateCardAndCredentialSet `
            -UserID $card.UserID `
            -AMTIDHolder $card.AMTIDHolder `
            -CardID $card.CardID `
            -CardSetID $card.CardSetID `
            -CredentialType $card.CredentialType `
            -EncodedID $card.EncodedID `
            -EmbossedID $card.EmbossedID -PINCode $newpin -ExpirationDateTime "2019-05-18 12:00:00" -RawBitPattern $card.RawBitPattern
        $output += [PSCustomObject] @{Hall=$card.LastName; Room=$card.FirstName; Pin=$newpin}
    }

    $output | Export-Csv output.csv -NoTypeInformation
}