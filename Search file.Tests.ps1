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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*The value 'a' in 'SendMail.When' is not supported. Only the value 'Always' or 'OnlyWhenFilesAreFound' can be used.")
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
    It 'once for 1 computer name and 1 database name' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = 'PC1'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
            ($ComputerName -eq $env:COMPUTERNAME) -and
            ($FilePath -eq $testParams.SqlScript) -and
            ($EnableNetworkAccess) -and
            ($ErrorAction -eq 'Stop') -and
            ($ArgumentList[0] -eq $testNewInputFile.Tasks[0].ComputerNames) -and
            ($ArgumentList[1] -eq $testNewInputFile.Tasks[0].DatabaseNames) -and
            ($ArgumentList[2][0] -eq $testData.sqlFile[0].Content) -and
            ($ArgumentList[2][1] -eq $testData.sqlFile[1].Content) -and
            ($ArgumentList[3][0] -eq $testData.sqlFile[0].Path) -and
            ($ArgumentList[3][1] -eq $testData.sqlFile[1].Path)
        }

        Should -Invoke Invoke-Command -Times 1 -Exactly
    }
    It 'once for each computer name and each database name' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerNames = @('PC1', 'PC2')
        $testNewInputFile.Tasks[0].DatabaseNames = @('db1', 'db2')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        foreach (
            $testComputer in
            $testNewInputFile.Tasks[0].ComputerNames
        ) {
            foreach (
                $testDatabase in
                $testNewInputFile.Tasks[0].DatabaseNames
            ) {
                Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
                    ($ComputerName -eq $env:COMPUTERNAME) -and
                    ($FilePath -eq $testParams.SqlScript) -and
                    ($EnableNetworkAccess) -and
                    ($ErrorAction -eq 'Stop') -and
                    ($ArgumentList[0] -eq $testComputer) -and
                    ($ArgumentList[1] -eq $testDatabase) -and
                    ($ArgumentList[2][0] -eq $testData.sqlFile[0].Content) -and
                    ($ArgumentList[2][1] -eq $testData.sqlFile[1].Content) -and
                    ($ArgumentList[3][0] -eq $testData.sqlFile[0].Path) -and
                    ($ArgumentList[3][1] -eq $testData.sqlFile[1].Path)
                }
            }
        }

        Should -Invoke Invoke-Command -Times 4 -Exactly
    }
}