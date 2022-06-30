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
            Context 'Tasks' {
                It 'is missing' {
                    @{
                        MailTo = @('bob@contoso.com')
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Property 'Tasks' not found.")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'MailTo is missing' {
                    @{
                        Tasks = @(
                            @{
                                # MailTo            = @('bob@contoso.com')
                                FolderPath        = $testFolderPath
                                Word              = 'kiwi'
                                MailOnlyWhenFound = $true
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Property 'MailTo' is mandatory.")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Word is missing' {
                    @{
                        Tasks = @(
                            @{
                                MailTo            = @('bob@contoso.com')
                                FolderPath        = $testFolderPath
                                # Word              = 'kiwi'
                                MailOnlyWhenFound = $true
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Property 'Word' is mandatory.")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                Context 'FolderPath' {
                    It 'is missing' {
                        @{
                            Tasks = @(
                                @{
                                    MailTo            = @('bob@contoso.com')
                                    # FolderPath        = $testFolderPath
                                    Word              = 'kiwi'
                                    MailOnlyWhenFound = $true
                                }
                            )
                        } | ConvertTo-Json | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*Property 'FolderPath' is mandatory.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'does not exist' {
                        @{
                            Tasks = @(
                                @{
                                    MailTo            = @('bob@contoso.com')
                                    FolderPath        = 'x:\notExisting'
                                    Word              = 'kiwi'
                                    MailOnlyWhenFound = $true
                                }
                            )
                        } | ConvertTo-Json | Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*The path 'x:\notExisting' in 'FolderPath' does not exist.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                Context 'MailOnlyWhenFound' {
                    It 'is missing' {
                        @{
                            Tasks = @(
                                @{
                                    MailTo            = @('bob@contoso.com')
                                    FolderPath        = $testFolderPath
                                    Word              = 'kiwi'
                                    # MailOnlyWhenFound = $true
                                }
                            )
                        } | ConvertTo-Json | Out-File @testOutParams
                    
                        .$testScript @testParams
        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*The property 'MailOnlyWhenFound' is mandatory.")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'MailOnlyWhenFound is not a boolean' {
                        @{
                            Tasks = @(
                                @{
                                    MailTo            = @('bob@contoso.com')
                                    FolderPath        = $testFolderPath
                                    Word              = 'kiwi'
                                    MailOnlyWhenFound = 'a'
                                }
                            )
                        } | ConvertTo-Json | Out-File @testOutParams
                    
                        .$testScript @testParams
        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*The value 'a' in 'MailOnlyWhenFound' is not a true false value*")
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
