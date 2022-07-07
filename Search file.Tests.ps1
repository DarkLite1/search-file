#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $realCmdLet = @{
        StartJob = Get-Command Start-Job
    }

    $testFolderPath = (New-Item "TestDrive:/folder" -ItemType Directory).FullName

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

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
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
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
            Context 'MaxConcurrentJobs' {
                It 'is missing' {
                    @{
                        # MaxConcurrentJobs = 6
                        Tasks = @(
                            @{
                                ComputerName = $null
                                FolderPath   = @($testFolderPath)
                                Filter       = '*kiwi*'
                                Recurse      = $false
                                SendMail     = @{
                                    To   = @('bob@contoso.com')
                                    When = 'Always'
                                }
                            }
                        )
                    } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

                    .$testScript @testParams
                            
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not a number' {
                    @{
                        MaxConcurrentJobs = 'a'
                        Tasks             = @(
                            @{
                                ComputerName = $null
                                FolderPath   = @($testFolderPath)
                                Filter       = '*kiwi*'
                                Recurse      = $false
                                SendMail     = @{
                                    To   = @('bob@contoso.com')
                                    When = 'Always'
                                }
                            }
                        )
                    } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                    
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
            Context 'Tasks' {
                It 'is missing' {
                    @{
                        MaxConcurrentJobs = 6
                    } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Property 'Tasks' not found.")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                Context 'Filter' {
                    It 'is missing' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = $testFolderPath
                                    # Filter     = '*kiwi*'
                                    Recurse      = $false
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'Always'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*Property 'Filter' is mandatory.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'is not unique' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = $testFolderPath
                                    Filter       = @('*a*', '*a*', '*b*')
                                    Recurse      = $false
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'Always'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*Property 'Filter' contains the duplicate value '*a*'. Duplicate values are not allowed.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                Context 'Recurse' {
                    It 'is missing' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = $testFolderPath
                                    Filter       = '*kiwi*'
                                    # Recurse    = $false
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'Always'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*Property 'Recurse' is mandatory.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'is not a true false value' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = $testFolderPath
                                    Filter       = @('*a*')
                                    Recurse      = 'a'
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'Always'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*The value 'a' in 'Recurse' is not a true false value.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                Context 'FolderPath' {
                    It 'is missing' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    # FolderPath = $testFolderPath
                                    Filter       = '*kiwi*'
                                    Recurse      = $false
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'Always'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*Property 'FolderPath' is mandatory.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'is duplicate' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = @(
                                        $testFolderPath, $testFolderPath
                                    )
                                    Filter       = '*kiwi*'
                                    Recurse      = $false
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'Always'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*Property 'FolderPath' contains the duplicate value '$testFolderPath'. Duplicate values are not allowed.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                Context 'SendMail' {
                    It 'is missing' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = $testFolderPath
                                    Filter       = '*kiwi*'
                                    Recurse      = $false
                                    # SendMail   = @{
                                    #     To   = @('bob@contoso.com')
                                    #     When = 'Always'
                                    # }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                    
                        .$testScript @testParams
        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'SendMail' is mandatory.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'When is not the value Always or OnlyWhenFilesAreFound' {
                        @{
                            MaxConcurrentJobs = 6
                            Tasks             = @(
                                @{
                                    ComputerName = $null
                                    FolderPath   = $testFolderPath
                                    Filter       = '*kiwi*'
                                    Recurse      = $false
                                    SendMail     = @{
                                        To   = @('bob@contoso.com')
                                        When = 'a'
                                    }
                                }
                            )
                        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams
                    
                        .$testScript @testParams
        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*The value 'a' in 'SendMail.When' is not supported. Only the value 'Always' or 'OnlyWhenFilesAreFound' can be used.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
            }
        }
    }
}
Describe 'when matching files are found' {
    BeforeAll {
        Remove-Item "$testFolderPath\*" -Recurse -Force

        $testFile = @(
            'a kiwi a.txt',
            'b kiwi b.txt',
            'c kiwi c.txt'
        ) | ForEach-Object {
            New-Item -Path "$testFolderPath\$_" -ItemType File
        }

        @{
            MaxConcurrentJobs = 6
            Tasks             = @(
                @{
                    ComputerName = $null
                    FolderPath   = $testFolderPath
                    Filter       = '*kiwi*'
                    Recurse      = $false
                    SendMail     = @{
                        To   = @('bob@contoso.com')
                        When = 'Always'
                    }
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Path          = $testFolderPath
                    Filter        = '*kiwi*'
                    Recurse       = $false
                    File          = $testFile[0].FullName
                    CreationTime  = $testFile[0].CreationTime
                    LastWriteTime = $testFile[0].LastWriteTime
                    Size          = [MATH]::Round($testFile[0].Length / 1GB, 2) 
                    Size_         = $testFile[0].Length
                    Duration      = '00:00:*'
                }
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Path          = $testFolderPath
                    Filter        = '*kiwi*'
                    Recurse       = $false
                    File          = $testFile[1].FullName
                    CreationTime  = $testFile[1].CreationTime
                    LastWriteTime = $testFile[1].LastWriteTime
                    Size          = [MATH]::Round($testFile[1].Length / 1GB, 2) 
                    Size_         = $testFile[1].Length
                    Duration      = '00:00:*'
                }
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Path          = $testFolderPath
                    Filter        = '*kiwi*'
                    Recurse       = $false
                    File          = $testFile[2].FullName
                    CreationTime  = $testFile[2].CreationTime
                    LastWriteTime = $testFile[2].LastWriteTime
                    Size          = [MATH]::Round($testFile[2].Length / 1GB, 2) 
                    Size_         = $testFile[2].Length
                    Duration      = '00:00:*'
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - 0 - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Files'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.File -eq $testRow.File
                }
                $actualRow.File | Should -Be $testRow.File
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.Path | Should -Be $testRow.Path
                $actualRow.Filter | Should -Be $testRow.Filter
                $actualRow.Recurse | Should -Be $testRow.Recurse
                $actualRow.CreationTime.ToString('yyyyMMdd HHmmss') | 
                Should -Be $testRow.CreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.LastWriteTime.ToString('yyyyMMdd HHmmss') | 
                Should -Be $testRow.LastWriteTime.ToString('yyyyMMdd HHmmss')
                $actualRow.Size | Should -Be $testRow.Size
                $actualRow.Size_ | Should -Be $testRow.Size_
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Duration | Should -BeLike $testRow.Duration
            }
        }
    }
    Context 'send a mail to the user when SendMail.When is Always' {
        BeforeAll {
            $testMail = @{
                To          = 'bob@contoso.com'
                Bcc         = $ScriptAdmin
                Priority    = 'Normal'
                Subject     = '3 files found'
                Message     = "*Found a total of <b>3 files</b>*$env:COMPUTERNAME*$testFolderPath*Filter*Files found**kiwi*3*Check the attachment for details*"
                Attachments = '* - 0 - Log.xlsx'
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Priority -eq $testMail.Priority) -and
                ($Subject -eq $testMail.Subject) -and
                ($Attachments -like $testMail.Attachments) -and
                ($Message -like $testMail.Message)
            }
        }
    }
}
Describe 'when no matching files are found' {
    BeforeAll {
        Remove-Item "$testFolderPath\*" -Recurse -Force

        $testFile = @(
            'a.txt',
            'b.txt',
            'c.txt'
        ) | ForEach-Object {
            New-Item -Path "$testFolderPath\$_" -ItemType File
        }

        @{
            MaxConcurrentJobs = 6
            Tasks             = @(
                @{
                    ComputerName = $null
                    FolderPath   = $testFolderPath
                    Filter       = '*.pst'
                    Recurse      = $false
                    SendMail     = @{
                        To   = @('bob@contoso.com')
                        When = 'Always'
                    }
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    It 'create no Excel file' {
        $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'
        $testExcelLogFile | Should -BeNullOrEmpty
    }
    Context 'send a mail to the user when SendMail.When is Always' {
        BeforeAll {
            $testMail = @{
                To       = 'bob@contoso.com'
                Bcc      = $ScriptAdmin
                Priority = 'Normal'
                Subject  = '0 files found'
                Message  = "*Found a total of <b>0 files</b>*$env:COMPUTERNAME*$testFolderPath*Filter*Files found*.pst*0*"
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeNullOrEmpty
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Priority -eq $testMail.Priority) -and
                ($Subject -eq $testMail.Subject) -and
                (-not $Attachments) -and
                ($Message -like $testMail.Message)
            }
        }
    }
    It 'send no mail when SendMai.When is OnlyWhenFilesAreFound' {
        @{
            MaxConcurrentJobs = 6
            Tasks             = @(
                @{
                    ComputerName = $null
                    FolderPath   = $testFolderPath
                    Filter       = '*.pst'
                    Recurse      = $false
                    SendMail     = @{
                        To   = @('bob@contoso.com')
                        When = 'OnlyWhenFilesAreFound'
                    }
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams

        Should -Not -Invoke Send-MailHC
    }
}
Describe 'when an error happens while searching for files' {
    BeforeAll {
        Mock Start-Job {
            & $realCmdLet.StartJob -Scriptblock { 
                throw 'oops'
            }
        }

        @{
            MaxConcurrentJobs = 6
            Tasks             = @(
                @{
                    ComputerName = $null
                    FolderPath   = $testFolderPath
                    Filter       = '*.pst'
                    Recurse      = $false
                    SendMail     = @{
                        To   = @('bob@contoso.com')
                        When = 'Always'
                    }
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context "export an Excel file with worksheet 'Errors'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    ComputerName = $env:COMPUTERNAME
                    Path         = $testFolderPath
                    Filters      = '*.pst'
                    Duration     = '00:00:*'
                    Error        = 'oops'
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - 0 - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Errors'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.ComputerName -eq $testRow.ComputerName
                }
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.Path | Should -Be $testRow.Path
                $actualRow.Filters | Should -Be $testRow.Filters
                $actualRow.Duration | Should -BeLike $testRow.Duration
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'send a mail to the user when SendMail.When is Always' {
        BeforeAll {
            $testMail = @{
                To          = 'bob@contoso.com'
                Bcc         = $ScriptAdmin
                Priority    = 'High'
                Subject     = '0 files found, 1 error'
                Message     = "*Detected <b>1 error</b> during execution.*Found a total of <b>0 files</b>*$env:COMPUTERNAME*$testFolderPath*Filter*Files found**.pst*0*Check the attachment for details*"
                Attachments = '* - 0 - Log.xlsx'
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Priority | Should -Be $testMail.Priority
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Priority -eq $testMail.Priority) -and
                ($Subject -eq $testMail.Subject) -and
                ($Attachments -like $testMail.Attachments) -and
                ($Message -like $testMail.Message)
            }
        }
    }
}