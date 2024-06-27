#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Id      = 1
        Path    = (New-Item 'TestDrive:/f' -ItemType Directory).FullName
        Filter  = '*.txt'
        Recurse = $false
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('Id', 'Path', 'Recurse', 'Filter') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'return an object' {
    BeforeAll {
        $testFiles = @(
            'file1.txt', 'file2.txt', 'file3.doc'
        ).ForEach(
            { New-Item -Path "$($testParams.Path)\$_" -ItemType File }
        )

        $testNewParams = $testParams.clone()
    }
    Context 'when files are found' {
        BeforeAll {
            $testNewParams.Filter = '*.txt'

            $actual = & $testScript @testNewParams
        }
        Context 'with property' {
            It 'Files' {
                $expected = $testFiles.where(
                    { $_.extension -eq '.txt' }
                )

                $actual.Files | Should -HaveCount $expected.Count

                foreach ($file in $expected) {
                    $actual.Files.FullName | Should -Contain $file.FullName
                }
            }
            It 'StartTime' {
                $actual.StartTime | Should -BeOfType 'System.DateTime'
            }
            It 'EndTime' {
                $actual.EndTime | Should -BeOfType 'System.DateTime'
            }
            It 'Error' {
                $actual.Error | Should -BeNullOrEmpty
            }
        }
    }
    Context 'when no files are found' {
        BeforeAll {
            $testNewParams.Filter = '*.xlsx'

            $actual = & $testScript @testNewParams
        }
        Context 'with property' {
            It 'Id' {
                $actual.Id | Should -Be $testNewParams.Id
            }
            It 'Files' {
                $actual.Files | Should -BeNullOrEmpty
            }
            It 'StartTime' {
                $actual.StartTime | Should -BeOfType 'System.DateTime'
            }
            It 'EndTime' {
                $actual.EndTime | Should -BeOfType 'System.DateTime'
            }
            It 'Error' {
                $actual.Error | Should -BeNullOrEmpty
            }
        }
    }
    Context 'with Error when' {
        Context 'Get-ChildItem fails' {
            BeforeAll {
                Mock Get-ChildItem {
                    throw 'Oops'
                }

                $actual = & $testScript @testNewParams
            }
            Context 'with property' {
                It 'Id' {
                    $actual.Id | Should -Be $testNewParams.Id
                }
                It 'Files' {
                    $actual.Files | Should -BeNullOrEmpty
                }
                It 'StartTime' {
                    $actual.StartTime | Should -BeOfType 'System.DateTime'
                }
                It 'EndTime' {
                    $actual.EndTime | Should -BeOfType 'System.DateTime'
                }
                It 'Error' {
                    $actual.Error | Should -Be "Failed retrieving files with filter '$($testNewParams.Filter)': Oops"
                }
            }
        }
        Context 'Path is not found' {
            BeforeAll {
                $testNewParams.Path = 'z:\notExisting'

                $actual = & $testScript @testNewParams
            }
            Context 'with property' {
                It 'Id' {
                    $actual.Id | Should -Be $testNewParams.Id
                }
                It 'Files' {
                    $actual.Files | Should -BeNullOrEmpty
                }
                It 'StartTime' {
                    $actual.StartTime | Should -BeOfType 'System.DateTime'
                }
                It 'EndTime' {
                    $actual.EndTime | Should -BeOfType 'System.DateTime'
                }
                It 'Error' {
                    $actual.Error | Should -Be "Failed retrieving files with filter '$($testNewParams.Filter)': Folder '$($testNewParams.Path)' not found."
                }
            }
        }
    }
}