Add-Type -Path "$PSScriptRoot\dlls\itext.kernel.dll"

# Regex options
$m = [System.Text.RegularExpressions.RegexOptions]::Multiline
$s = [System.Text.RegularExpressions.RegexOptions]::Singleline

# Convert PDF to text
$path = "C:\Users\Feldmans\Downloads\20141128-statements-3099- (1).pdf"
$pdfDoc  = [iText.Kernel.Pdf.PdfDocument]::new([iText.Kernel.Pdf.PdfReader]::new($path))
$text = for ($a = 1; $a -le $pdfdoc.GetNumberOfPages(); $a++) {
    [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($pdfDoc.GetPage($a))
}
$text = $text -join "`n"

# For newer accessible PDF's (From circa Nov 2016) there is some markers we can use to break up the document based on account and extract the transactions.
$regex = [regex]::Matches($text, "^\*start\*globalproduct$.*?^\*start\*posttransaction detail message$", @($s,$m))

# Confirm the document is indeed a newer one, otherwise this will be empty.
if ([string]::IsNullOrWhiteSpace($regex)) {
    # In order not to need to have seperate code for older and newer statements, we are imitating the regex object so the code will work on both the same.
    $regex = [pscustomobject]@{value = $text}
}
    
# Get the dates the statement is refering to. It is important as the dates on the transactions is missing the year so we are using the year from this date to fix that.
$dates = [regex]::Match($text,".*[0-9]{2}, [0-9]{4} through .* [0-9]{2}, [0-9]{4}").value -split "through" | foreach {$_ | get-date}

$regex.value | foreach-object {
    $trans = ($_ -split "`n") | where-object {$_ -match  "[0-9]{2}\/[0-9]{2}"}
    $transaction = $trans | ForEach-Object {
        $tr_regex = [regex]::Match($_,"([0-9]+\/[0-9]+) (.*) (-?[0-9]+\,?[0-9]*\.[0-9]+) (-?[0-9]+\,?[0-9]*\.[0-9]+)").Groups[1..4].value
        
        # Add the year to the transaction date. 
        if (($tr_regex[0] -split "/")[0] -lt 12) {
            $Date = "$($tr_regex[0])/$($dates[1].Year)"
        } else {
            $date = "$($tr_regex[0])/$($dates[0].Year)"
        }


        [pscustomobject]@{
            Date        = $date
            Memo        = $tr_regex[1]
            Amount      = $tr_regex[2]
        }
    }
    [pscustomobject]@{
        account = [regex]::Match($_, "(?<=Account Number: ).*?$",@($s,$m)).value
        transactions = $transaction
    }
}