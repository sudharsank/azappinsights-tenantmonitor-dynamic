// avoids "TS2304: Cannot find name 'CONSTANT'" TSC error on builds because these
// strings are not defined in the code, rather webpack replaces them on build

declare var AZURE_APPINSIGHTS_INSTRUMENTATION_KEY: string
declare var AZURE_APPINSIGHTS_APP_NAME: string
declare var AZURE_APPINSIGHTS_APP_VERSION: string
declare var AZURE_APPINSIGHTS_SOLUTION_ID: string
declare var AZURE_APPINSIGHTS_SOLUTION_VERSION: string
