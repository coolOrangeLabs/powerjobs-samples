# ============================================================================#
# PowerShell module sample for coolOrange powerJobs                           #
# Generates a token for WeTransfer                                            #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
# ============================================================================#

#https://developers.wetransfer.com/documentation/#api-keys-where-and-how
$apiKey = "BdRdwz6Yiiavof27O4Tkra7YRujMeiJn1GgLvIYv"

function Get-AccessToken {
    $Headers = @{
        "x-api-key"=$apiKey;
    }    
    $res = Invoke-RestMethod -Uri "https://dev.wetransfer.com/v2/authorize" -ContentType "application/json; charset=utf-8" -Headers $Headers -Method POST

    return $res.token
}

function New-WeTransfer([String]$File, [String]$Token){
    $f = $(Get-Item $File)

    $Headers = @{
        "x-api-key"=$apiKey;
        "Authorization"="Bearer $Token"
    }
    $Body = @{
        "message"="coolOrangeUpload";
        "files"=@(
            @{
                "name"=$f.Name;
                "size"=$f.Length;
            }
        )
    } | ConvertTo-Json
    $transfer = Invoke-RestMethod -Uri "https://dev.wetransfer.com/v2/transfers" -ContentType "application/json; charset=utf-8" -Headers $Headers -Body $Body -Method POST

    return $transfer
}

function Request-UploadUrl($Transfer, [String]$File, [String]$Token){
    Write-Host $Transfer

    $transferId = $Transfer.id
    $fileId = $Transfer.files[0].id
    $partNumbers = $Transfer.files[0].multipart.part_numbers

    $Headers = @{
        "x-api-key"=$apiKey;
        "Authorization"="Bearer $Token"
    }
    [System.IO.BinaryReader]$reader = New-Object System.IO.BinaryReader([System.IO.File]::Open($File, [System.IO.FileMode]::Open))

    for ($i = 1; $i -le $partNumbers; $i++){
        $url = Invoke-RestMethod -Uri "https://dev.wetransfer.com/v2/transfers/$transferId/files/$fileId/upload-url/$i" -ContentType "application/json; charset=utf-8" -Headers $Headers -Method GET
        
        $buffer = [byte[]]::new(5242880)
        $count = $reader.Read($buffer, 0, $buffer.Length)

        [System.Array]::Resize([ref]$buffer, $count)

        Invoke-RestMethod -Uri $url.url -ContentType "application/octet-stream" -Method PUT -Body $buffer
    }
    $reader.Close()
}

function Complete-WeTransfer($Transfer, [String]$Token) {
    $transferId = $Transfer.id
    $fileId = $Transfer.files[0].id
    $partNumbers = $Transfer.files[0].multipart.part_numbers

    $Headers = @{
        "x-api-key"=$apiKey;
        "Authorization"="Bearer $Token";
    }
    $Body = @{
        "part_numbers"=$partNumbers
    } | ConvertTo-Json
    Invoke-RestMethod -Uri "https://dev.wetransfer.com/v2/transfers/$transferId/files/$fileId/upload-complete" -ContentType "application/json; charset=utf-8" -Headers $Headers -Body $Body -Method PUT
}

function Close-WeTransfer($Transfer, [String]$Token){
    $transferId = $Transfer.id

    $Headers = @{
        "x-api-key"=$apiKey;
        "Authorization"="Bearer $Token";
    }
    return $(Invoke-RestMethod -Uri "https://dev.wetransfer.com/v2/transfers/$transferId/finalize" -ContentType "application/json; charset=utf-8" -Headers $Headers -Method PUT)
}