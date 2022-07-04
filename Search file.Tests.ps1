#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
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
Describe 'when matching file names are found' {
    BeforeEach {
        Remove-Item "$testFolderPath\*" -Recurse -Force
    }
    Context 'and SendMail is Always' {
        It 'send an e-mail' {
            $testFile = @(
                'a kiwi a',
                'b kiwi b',
                'c kiwi c'
            ) | ForEach-Object {
                (New-Item -Path "$testFolderPath\$_" -ItemType File).FullName
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

            Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
                $Priority -eq 'Normal'
            }
        }
    }
} -Tag test