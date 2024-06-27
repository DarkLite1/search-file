#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path    = (New-Item 'TestDrive:/f' -ItemType Directory).FullName
        Filters = @('*.txt')
        Recurse = $false
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('Path', 'Recurse', 'Filters') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'throw a terminating error when' {
    It 'Path is not found' {
        $testNewParams = $testParams.clone()
        $testNewParams.Path = 'z:\notExisting'

        {
            & $testScript @testNewParams
        } |
        Should -Throw -ExpectedMessage "Folder '$($testNewParams.Path)' not found."
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
        $testNewParams.Filters = @('*.txt', '*.xlsx', '*.oops')

        Mock Get-ChildItem {
            throw 'Oops'
        } -ParameterFilter {
            $Filter -eq $testNewParams.Filters[2]
        }

        $actual = & $testScript @testNewParams
    }
    It 'for each filter' {
        $actual | Should -HaveCount $testNewParams.Filters.Count
    }
    Context 'when files are found' {
        Context 'with property' {
            It 'Files' {
                $expected = $testFiles.where(
                    { $_.extension -eq '.txt' }
                )

                $actual[0].Files | Should -HaveCount $expected.Count

                foreach ($file in $expected) {
                    $actual[0].Files.FullName | Should -Contain $file.FullName
                }
            }
            It 'Filter' {
                $actual[0].Filter | Should -Be $testNewParams.Filters[0]
            }
            It 'StartTime' {
                $actual[0].StartTime | Should -BeOfType 'System.DateTime'
            }
            It 'EndTime' {
                $actual[0].EndTime | Should -BeOfType 'System.DateTime'
            }
            It 'Error' {
                $actual[0].Error | Should -BeNullOrEmpty
            }
        }
    }
    Context 'when no files are found' {
        Context 'with property' {
            It 'Files' {
                $actual[1].Files | Should -BeNullOrEmpty
            }
            It 'Filter' {
                $actual[1].Filter | Should -Be $testNewParams.Filters[1]
            }
            It 'StartTime' {
                $actual[1].StartTime | Should -BeOfType 'System.DateTime'
            }
            It 'EndTime' {
                $actual[1].EndTime | Should -BeOfType 'System.DateTime'
            }
            It 'Error' {
                $actual[1].Error | Should -BeNullOrEmpty
            }
        }
    }
    Context 'when Get-ChildItem fails' {
        Context 'with property' {
            It 'Files' {
                $actual[2].Files | Should -BeNullOrEmpty
            }
            It 'Filter' {
                $actual[2].Filter | Should -Be $testNewParams.Filters[2]
            }
            It 'StartTime' {
                $actual[2].StartTime | Should -BeOfType 'System.DateTime'
            }
            It 'EndTime' {
                $actual[2].EndTime | Should -BeOfType 'System.DateTime'
            }
            It 'Error' {
                $actual[2].Error | Should -Be "Failed retrieving files with filter '$($testNewParams.Filters[2])': Oops"
            }
        }
    }
}