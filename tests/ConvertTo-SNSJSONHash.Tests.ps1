BeforeAll {
    . (Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\ConvertTo-SNSJSONHash.ps1')
}

Describe 'ConvertTo-SNSJSONHash' {
    It 'converts a flat JSON object to a hashtable' {
        $inputObject = '{"name":"Scoutnet","active":true}' | ConvertFrom-Json
        $result = $inputObject | ConvertTo-SNSJSONHash

        $result | Should -BeOfType [hashtable]
        $result['name'] | Should -Be 'Scoutnet'
        $result['active'] | Should -Be 'True'
    }

    It 'converts nested JSON objects to nested hashtables' {
        $inputObject = '{"settings":{"enabled":true,"count":3}}' | ConvertFrom-Json
        $result = $inputObject | ConvertTo-SNSJSONHash

        $result['settings'] | Should -BeOfType [hashtable]
        $result['settings']['enabled'] | Should -Be 'True'
        $result['settings']['count'] | Should -Be '3'
    }

    It 'converts each object passed through the pipeline' {
        $inputObjects = @(
            '{"id":"first"}' | ConvertFrom-Json
            '{"id":"second"}' | ConvertFrom-Json
        )
        $results = @($inputObjects | ConvertTo-SNSJSONHash)

        $results.Count | Should -Be 2
        $results[0]['id'] | Should -Be 'first'
        $results[1]['id'] | Should -Be 'second'
    }
}
