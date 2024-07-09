#Requires -Version 7
#Requires -Modules Pester

BeforeAll {
    $testInputFile = @{
        MaxConcurrentJobs = 1
        Tasks             = @(
            @{
                ComputerName = @('PC1')
                FolderPath   = @((New-Item "TestDrive:/folder" -ItemType Directory).FullName)
                Filter       = @('*.txt')
                Recurse      = $false
                SendMail     = @{
                    Header = $null
                    To     = 'bob@contoso.com'
                    When   = 'Always'
                }
            }
        )
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName   = 'Test (Brecht)'
        ImportFile   = $testOutParams.FilePath
        SearchScript = (New-Item 'TestDrive:/s.ps1' -ItemType File).FullName
        LogFolder    = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin  = 'admin@contoso.com'
    }

    Mock Invoke-Command
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'ImportFile' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    It 'the file SearchScript cannot be found' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.SearchScript = 'c:\upDoesNotExist.ps1'

        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Search script with path '$($testNewParams.SearchScript)' not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MaxConcurrentJobs', 'Tasks'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.<_> not found' -ForEach @(
                'FolderPath', 'Filter', 'ComputerName', 'SendMail'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Tasks.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SendMail.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Tasks.SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ComputerName is not unique' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].ComputerName = @('PC1', 'PC1', 'PC2')

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*Property 'ComputerName' contains the duplicate value 'PC1'. Duplicate values are not allowed.")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Filter is not unique' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Filter = @('*a*', '*a*', '*b*')

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*Property 'Filter' contains the duplicate value '*a*'. Duplicate values are not allowed.")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Recurse is not a true false value' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Recurse = 'a'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*The value 'a' in 'Recurse' is not a true false value.")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'FolderPath contains duplicates' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].FolderPath = @(
                    'a', 'a', 'b'
                )

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*Property 'FolderPath' contains the duplicate value 'a'. Duplicate values are not allowed.")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail is not Always or OnlyWhenFilesAreFound' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SendMail.When = 'a'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*The value 'a' in 'Tasks.SendMail.When' is not supported. Only the value 'Always' or 'OnlyWhenFilesAreFound' can be used.")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'MaxConcurrentJobs is not a number' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.MaxConcurrentJobs = 'a'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' needs to be a number, the value 'a' is not supported*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'execute the search script with Invoke-Command' {
    It 'once for each computer, path and filter' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = @('PC1', 'PC2')
        $testNewInputFile.Tasks[0].FolderPath = @('z:\a', 'z:\b')
        $testNewInputFile.Tasks[0].Filter = @('*.txt', '*.pst')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        foreach (
            $testComputer in
            $testNewInputFile.Tasks[0].ComputerName
        ) {
            foreach (
                $testPath in
                $testNewInputFile.Tasks[0].FolderPath
            ) {
                foreach (
                    $testFilter in
                    $testNewInputFile.Tasks[0].Filter
                ) {
                    Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
                        ($ComputerName -eq $testComputer) -and
                        ($FilePath -eq $testParams.SearchScript) -and
                        ($EnableNetworkAccess) -and
                        ($ErrorAction -eq 'Stop') -and
                        ($ArgumentList[0] -eq $testPath) -and
                        ($ArgumentList[1] -eq $testFilter) -and
                        ($ArgumentList[2] -eq $testNewInputFile.Tasks[0].Recurse)
                    }
                }
            }
        }

        Should -Invoke Invoke-Command -Times 8 -Exactly
    }
}
Describe 'SendMail.When' {
    Context 'Always' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'Always'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams
        }
        Context 'send an e-mail to the user' {
            It "when no files are found" {
                Mock Invoke-Command {
                    [PSCustomObject]@{
                        Id        = 1
                        Files     = @()
                        StartTime = Get-Date
                        EndTime   = (Get-Date).AddMinutes(5)
                        Error     = $null
                    }
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -ParameterFilter {
                    ($to -eq $testNewInputFile.Tasks[0].SendMail.To) -and
                    ($Bcc -eq $testParams.ScriptAdmin) -and
                    ($Priority -eq 'Normal') -and
                    ($Subject -eq '0 files found') -and
                    (-not $Attachments) -and
                    ($Message -like "*Found a total of <b>0 files</b>*$($testNewInputFile.Tasks[0].ComputerName)*$($testNewInputFile.Tasks[0].FolderPath.SubString(2, $testNewInputFile.Tasks[0].FolderPath.length -2))*Filter*Files found*$($testNewInputFile.Tasks[0].Filter)*0*")
                    }
            }
            It "when files are found" {
                Mock Invoke-Command {
                    [PSCustomObject]@{
                        Id        = 1
                        Files     = @('a', 'b')
                        StartTime = Get-Date
                        EndTime   = (Get-Date).AddMinutes(5)
                        Error     = $null
                    }
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -ParameterFilter {
                ($to -eq $testNewInputFile.Tasks[0].SendMail.To) -and
                ($Bcc -eq $testParams.ScriptAdmin) -and
                ($Priority -eq 'Normal') -and
                ($Subject -eq '2 files found') -and
                ($Attachments -like '* - 0 - Log.xlsx') -and
                ($Message -like "*Found a total of <b>2 files</b>*$($testNewInputFile.Tasks[0].ComputerName)*$($testNewInputFile.Tasks[0].FolderPath.SubString(2, $testNewInputFile.Tasks[0].FolderPath.length -2))*Filter*Files found*$($testNewInputFile.Tasks[0].Filter)*2*Check the attachment for details*")
                }
            }
        }
    }
    Context 'OnlyWhenFilesAreFound' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].SendMail.When = 'OnlyWhenFilesAreFound'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams
        }
        Context 'send no e-mail to the user' {
            It "when no files are found" {
                Mock Invoke-Command {
                    [PSCustomObject]@{
                        Id        = 1
                        Files     = @()
                        StartTime = Get-Date
                        EndTime   = (Get-Date).AddMinutes(5)
                        Error     = $null
                    }
                }

                .$testScript @testParams

                Should -Not -Invoke Send-MailHC
            }
        }
        Context 'send an e-mail to the user' {
            It "when files are found" {
                Mock Invoke-Command {
                    [PSCustomObject]@{
                        Id        = 1
                        Files     = @('a', 'b')
                        StartTime = Get-Date
                        EndTime   = (Get-Date).AddMinutes(5)
                        Error     = $null
                    }
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -ParameterFilter {
                ($to -eq $testNewInputFile.Tasks[0].SendMail.To) -and
                ($Bcc -eq $testParams.ScriptAdmin) -and
                ($Priority -eq 'Normal') -and
                ($Subject -eq '2 files found') -and
                ($Attachments -like '* - 0 - Log.xlsx') -and
                ($Message -like "*Found a total of <b>2 files</b>*$($testNewInputFile.Tasks[0].ComputerName)*$($testNewInputFile.Tasks[0].FolderPath.SubString(2, $testNewInputFile.Tasks[0].FolderPath.length -2))*Filter*Files found*$($testNewInputFile.Tasks[0].Filter)*2*Check the attachment for details*")
                }
            }
        }
    }
}  -Tag test