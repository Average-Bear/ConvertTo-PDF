function ConvertTo-Pdf($files, $outFile) {

<#param(
  [Parameter(Mandatory)]
  [string[]]$Files,

  [Parameter(Mandatory)]
  [string[]]$OutFile
)#>

    Add-Type -AssemblyName System.Drawing
    $files = @($files)

    if (!$outFile) {

        $firstFile = $files[0] 
        if ($firstFile.FullName) { $firstFile = $firstFile.FullName }
        $outFile = $firstFile.Substring(0, $firstFile.LastIndexOf(".")) + ".pdf"
    } 

    else {

        if (![System.IO.Path]::IsPathRooted($outFile)) {

            $outFile = [System.IO.Path]::Combine((Get-Location).Path, $outFile)
        }
    }

    try {

        $doc = [System.Drawing.Printing.PrintDocument]::new()
        $opt = $doc.PrinterSettings = [System.Drawing.Printing.PrinterSettings]::new()
        $opt.PrinterName = "Microsoft Print to PDF"
        $opt.PrintToFile = $true
        $opt.PrintFileName = $outFile

        $script:_pageIndex = 0
        $doc.add_PrintPage({

            param($sender, [System.Drawing.Printing.PrintPageEventArgs] $a)
            $file = $files[$script:_pageIndex]

            if ($file.FullName) {

                $file = $file.FullName
            }
            $script:_pageIndex = $script:_pageIndex + 1

            try {

                $image = [System.Drawing.Image]::FromFile($file)
                $a.Graphics.DrawImage($image, $a.PageBounds)
                $a.HasMorePages = $script:_pageIndex -lt $files.Count
            }
            finally {

                $image.Dispose()
            }
        })

        $doc.PrintController = [System.Drawing.Printing.StandardPrintController]::new()

        $doc.Print()
        return $outFile
    }
    finally {

        if ($doc) { $doc.Dispose() }
    }
}
