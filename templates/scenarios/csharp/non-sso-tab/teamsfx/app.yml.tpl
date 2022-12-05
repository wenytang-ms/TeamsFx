version: 1.0.0

provision:
  - uses: arm/deploy # Deploy given ARM templates parallelly.
    with:
      subscriptionId: ${{AZURE_SUBSCRIPTION_ID}} # The AZURE_SUBSCRIPTION_ID is a built-in environment variable. TeamsFx will ask you select one subscription if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select subscription if it's empty in this case.
      resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}} # The AZURE_RESOURCE_GROUP_NAME is a built-in environment variable. TeamsFx will ask you to select or create one resource group if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select or create resource grouop if it's empty in this case.
      templates:
       - path: ./infra/azure.bicep # Relative path to teamsfx folder
         parameters: ./infra/azure.parameters.json # Relative path to teamsfx folder. Placeholders will be replaced with corresponding environment variable before ARM deployment.
         deploymentName: Create-resources-for-tab # Required when deploy ARM template
      bicepCliVersion: v0.9.1 # Teams Toolkit will download this bicep CLI version from github for you, will use bicep CLI in PATH if you remove this config.
    # Output: every bicep output will be persisted in current environment's .env file with certain naming conversion. Refer https://aka.ms/teamsfx-actions/arm-deploy for more details on the naming conversion rule.

deploy:
  - uses: dotnet/command
    with:
      args: publish --configuration Release --runtime win-x86 --self-contained
  - uses: azureAppService/deploy
    with:
      # deploy base folder
      distributionPath: ./bin/Release/net6.0/win-x86/publish
      # the resource id of the cloud resource to be deployed to
      resourceId: ${{TAB_AZURE_APP_SERVICE_RESOURCE_ID}}

registerApp:
  - uses: teamsApp/create # Creates a Teams app
    with:
      name: ${{TEAMS_APP_NAME}} # Teams app name
    # Output: following environment variable will be persisted in current environment's .env file.
    # TEAMS_APP_ID: the id of Teams app

configureApp:
  - uses: teamsApp/validate
    with:
      manifestTemplatePath: ./appPackage/manifest.template.json # Path to manifest template
  - uses: teamsApp/createAppPackage # Build Teams app package with latest env value
    with:
      manifestTemplatePath: ./appPackage/manifest.template.json # Path to manifest template
      outputZipPath: ./build/appPackage/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./build/appPackage/manifest.${{TEAMSFX_ENV}}.json
  - uses: teamsApp/update # Apply the Teams app manifest to an existing Teams app. Will use the app id in manifest file to determine which Teams app to update.
    with:
      appPackagePath: ./build/appPackage/appPackage.${{TEAMSFX_ENV}}.zip # Relative path to teamsfx folder. This is the path for built zip file.
    # Output: following environment variable will be persisted in current environment's .env file.
    # TEAMS_APP_ID: the id of Teams app