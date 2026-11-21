#### applyTo: "tests/**/*.Tests.ps1"

#### Pester v5 Testing Conventions (Office365-Scoutnet-synk)

Follow these exact patterns when writing or updating unit tests in tests/: 

#### Module Lifecycle (BeforeAll / AfterAll)

* Import the module using $PSScriptRoot relative path and clean it up in AfterAll: 

powershell

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Office365-Scoutnet-synk\Office365-Scoutnet-synk.psm1'
    Import-Module $modulePath -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module Office365-Scoutnet-synk -ErrorAction SilentlyContinue
}

Use code with caution.

Copy code text to clipboard.

#### InModuleScope & External Cmdlet Stubs

* Wrap unit tests for internal and exported functions inside InModuleScope Office365-Scoutnet-synk { ... }.
* **Define dummy stub functions before Mock:** External modules (Microsoft.Graph.*, ExchangeOnlineManagement, Az.*) may not be loaded or connected in CI runners. Always declare lightweight function <CmdletName> { param(...) } stubs inside InModuleScope with matching parameter names (including [switch] parameters) before calling Mock <CmdletName>.
* **Reset module state:** When testing functions that rely on configuration state, initialize $script:SNSConf = [SNSConfiguration]::new() explicitly inside each It block so tests never leak state into one another.

#### Mocking & Assertions (Pester v5 Syntax)

* Always mock Write-SNSLog {} to prevent file I/O during tests, and verify log levels ('Info', 'Warn', 'Error') and messages using Should -Invoke Write-SNSLog -Times 1 -Exactly -ParameterFilter { ... }.
* Use Pester v5 syntax exclusively: 

  * Should -Invoke <Cmdlet> -Times <N> -Exactly (never legacy Assert-MockCalled).
  * { ... } | Should -Throw '*expected substring*' for error paths.
  * @(...).Count | Should -Be <N> when asserting collection counts to avoid single-item unwrapping issues.
* Use example.org or example.se for all test domains and email addresses, and never include real Scoutnet IDs, tokens, or personal data.
* Ensure all test files obey PSScriptAnalyzerSettings.psd1 (no unused variables, no cmdlet aliases, no trailing whitespace).
* Always clean up any mocked or stubbed cmdlets after each test to prevent cross-test contamination (use Remove-Mock or equivalent).
* Use BeforeEach and AfterEach blocks to set up and tear down test-specific state, ensuring isolation between tests.
* Always reset $script:SNSConf and any other module-level state in BeforeEach to prevent tests from affecting each other.