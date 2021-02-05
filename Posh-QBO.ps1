Function Get-qTransactionMapping {
    return @{
        Begin       = "<STMTTRN>"
        Type        = "<TRNTYPE>"
        Date        = "<DTPOSTED>"
        Amount      = "<TRNAMT>"
        ID          = "<FITID>"
        Check       = "<CHECKNUM>"
        name        = "<NAME>"
        Memo        = "<MEMO>"
        END         = "</STMTTRN>"
    }
}

Function Get-qCurrentTabsactionHash {
    param(
        $headers
    )
    $transaction_hash = [ordered]@{
        "Begin"                 = "<STMTTRN>"
    }
    $headers | foreach-object {
        $transaction_hash[$_.name] = (Get-qTransactionMapping).$($_.name)
    }
    $transaction_hash["end"] = "</STMTTRN>"
    return $transaction_hash
}

function Get-qTransactions {
    param(
        $path,
        $Company = "Target",
        [Parameter(Mandatory)]
        [ValidateSet("CreditCard","BankAccount")]
        $Type,
        $BankID = "0210021",
        $AccountNumber = "1234"
    )
    # Import csv
    $csv = Import-Csv $path

    # Get CSV mapping to standarized column names.
    $mapping  = get-content '.\QBO mapping.json' | ConvertFrom-Json

    # Convert column names to stadarized names, based on mapping.
    $data = $csv | foreach-object {
        $current = $_
        $hash = [ordered]@{}
        $mapping.$company.psobject.Properties | foreach-object {
            $hash[$_.name] = $current.$($_.value)
        }
        [pscustomobject]$hash
    }

    $id = 201901141
    # Sanitize Amount value
    $data | foreach-object {
        $_.amount       = [DOUBLE]$_.amount.replace("$","").replace("(","-").replace(")","")
        $_.date         = "$(($_.date | get-date).ToString("yyyyMMddHHMMss")).000[-5]"
        if ($type -eq "CreditCard") {
            $_.amount  = $_.amount * -1
        }
        if ($_.amount -lt 0) {$_.type = "Debit"} else {$_.type = "Credit"}
        $_.type         = ($_.type).ToUpper()
        $_.id           = $id++
    }

    $transactions = $data | ForEach-Object {
        $trns_mp = (Get-qCurrentTabsactionHash -headers $mapping.$company.psobject.Properties)
        $current = $_
        $mapping.$company.psobject.Properties.name | ForEach-Object {
            $trns_mp.$psitem += $current.$_
        }
        $trns_mp
    }

    switch ($type) {
        "CreditCard"        {
            $BeginType = "<CREDITCARDMSGSRSV1>`n<CCSTMTTRNRS>`n<TRNUID>1`n<STATUS>`n<CODE>0`n<SEVERITY>INFO`n<MESSAGE>Success`n</STATUS>`n<CCSTMTRS>`n<CURDEF>USD`n<CCACCTFROM>`n<ACCTID>$($AccountNumber)`n</CCACCTFROM>"
            $EndType = "</STMTRS>`n</STMTTRNRS>`n</BANKMSGSRSV1>`n</OFX>"
            break
        }
        "BankAccount"       {
            $BeginType = "<BANKMSGSRSV1>`n<STMTTRNRS>`n<TRNUID>1`n<STATUS>`n<CODE>0`n<SEVERITY>INFO`n<MESSAGE>Success`n</STATUS>`n<STMTRS>`n<CURDEF>USD`n<BANKACCTFROM>`n<BANKID>$($BankID)`n<ACCTID>$($AccountNumber)`n<ACCTTYPE>CHECKING`n</BANKACCTFROM>"
            $EndType = "</CCSTMTRS>`n</CCSTMTTRNRS>`n</CREDITCARDMSGSRSV1>`n</OFX>"
            break
        }
    }
    $beginfile = "OFXHEADER:100`nDATA:OFXSGML`nVERSION:102`nSECURITY:NONE`nENCODING:USASCII`nCHARSET:1252`nCOMPRESSION:NONE`nOLDFILEUID:NONE`nNEWFILEUID:NONE`n<OFX>`n<SIGNONMSGSRSV1>`n<SONRS>`n<STATUS>`n<CODE>0`n<SEVERITY>INFO`n</STATUS>`n<DTSERVER>20210110120000[0:GMT]`n<LANGUAGE>ENG`n<FI>`n<ORG>B1`n<FID>10898`n</FI>`n<INTU.BID>2430`n</SONRS>`n</SIGNONMSGSRSV1>"
    $beginfile2 = "<BANKTRANLIST>`n<DTSTART>20190114120000[0:GMT]`n<DTEND>20210104120000[0:GMT]"
    $endfile = "</BANKTRANLIST>`n<LEDGERBAL>`n<BALAMT>1`n<DTASOF>20210110120000[0:GMT]`n</LEDGERBAL>`n<AVAILBAL>`n<BALAMT>1`n<DTASOF>20210110120000[0:GMT]`n</AVAILBAL>"

    $beginfile
    $BeginType
    $beginfile2
    $transactions.values
    $endfile
    $EndType
}

Function Convert-qCSV2QBO {
    param(
        [parameter(Mandatory)]
        $Source,
        [parameter(Mandatory)]
        $Company,
        [parameter(Mandatory)]
        $Type,
        [parameter(Mandatory)]
        $Destination,
        [parameter(Mandatory)]
        $OUtputFileName
    )
    Get-qTransactions -path $source -company target -Type CreditCard | Out-File "$destination\$OUtputFileName.qbo" -Encoding ascii
}
