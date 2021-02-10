param(
    [CmdletBinding()]
    $path
)

# Parse the document using PdfPig
Add-Type -Path "$PSScriptRoot\dlls\UglyToad.PdfPig.dll"
$document = [UglyToad.PdfPig.PdfDocument]::Open($path)

# Extract rows from PDF.
$words = for ($i = 1; $i -lt $document.NumberOfPages; $i++) {
    # Sort the rows page by page, otherwisw thw rows from all pages get mixed up, as they use coordinates per page. 
    $document.GetPage($i).getwords() | Group-Object {$_.BoundingBox.bottom} | Sort-Object -Descending {[double]$_.name}
}
# Flatten text output to single string object. Needed to be able to easily split by account number.
$text = ($words | ForEach-Object {$_.group -join " "}) -join "`n"

# Split the document by accounts
$regex = [pscustomobject]@{value = ($text -split "Account number") | Where-Object {$_[0] -eq ":"}}
    
# Get the dates the statement is refering to. It is important as the dates on the transactions is missing the year so we are using the year from this date to fix that.
$dates = [regex]::Match($text,"\w+ [0-9]{2}, [0-9]{4} through ?\w+ [0-9]{2}, [0-9]{4}").value -split "through" | foreach-object {$_ | get-date}

$regex.value | foreach-object {
    $tr_regex = [regex]::Matches($_,"([0-9]+\/[0-9]+) (.*) (-?[0-9]+\,?[0-9]*\.[0-9]+) (-?[0-9]+\,?[0-9]*\.[0-9]+)")
    
    $account_number = [regex]::Match($_, "(?<=: )[0-9]+\b").value

    $transaction = $tr_regex | ForEach-Object {

        # Convert short date to full date. Utilizes the year listed in the document statement range
        if (($_.groups[1].value -split "/")[0] -lt 12) {
            $date = "$($_.groups[1].value)/$($dates[1].Year)"
        } else {
            $date = "$($_.groups[1].value)/$($dates[0].Year)"
        }

        # Fix where sometimes the PDF parser places a space between the negative symbol and the amount,
        # and the amount is wrong plus the negative is included in the memo.
        if (($_.Groups[2].Value)[-1] -eq "-") {
            $memo       = ($_.Groups[2].Value).TrimEnd(" -")
            $amount     = "-$($_.Groups[3].Value)"
        } else {
            $memo       = $_.Groups[2].Value
            $amount     = $_.Groups[3].Value
        }

        [pscustomobject]@{
            Date        = $date
            Memo        = $memo
            Amount      = $amount
            Account     = $account_number
            file        = [io.path]::GetFileName($path)
        }
    }
    [pscustomobject]@{
        account = $account_number
        transactions = $transaction
        file        = [io.path]::GetFileName($path)
    }
}