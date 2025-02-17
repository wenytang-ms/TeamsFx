{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch Remote (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "3-remote",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "3-remote",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Local Service"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Launch App (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Local Service"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Attach to Local Service",
            "type": "node",
            "request": "attach",
            "port": 9239,
            "restart": true,
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote (Desktop)",
            "type": "node",
            "request": "launch",
            "preLaunchTask": "Start Teams App in Desktop Client (Remote)",
            "presentation": {
                "group": "3-remote",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen",
        {{#CEAEnabled}}
        },
        {
            "name": "Launch Remote in Copilot (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&${account-hint}",
            "cascadeTerminateToConfigurations": ["Attach to Local Service"],
            "presentation": {
                "group": "3-M365",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen"
            },
            {
            "name": "Launch Remote in Copilot (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&${account-hint}",
            "cascadeTerminateToConfigurations": ["Attach to Local Service"],
            "presentation": {
                "group": "3-M365",
                "order": 4
            },
            "internalConsoleOptions": "neverOpen"
        {{/CEAEnabled}}
        }
    ],
    "compounds": [
        {
            "name": "Debug in Teams (Edge)",
            "configurations": [
                "Launch App (Edge)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Chrome)",
            "configurations": [
                "Launch App (Chrome)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Desktop)",
            "configurations": [
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start Teams App in Desktop Client",
            "presentation": {
                "group": "2-local",
                "order": 3
            },
            "stopAll": true
        },
        {
            "name": "Debug in Test Tool",
            "configurations": [
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start Teams App (Test Tool)",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        {{#CEAEnabled}}
        },
        {
            "name": "Debug in Copilot (Edge)",
            "configurations": [
                "Launch Remote in Copilot (Edge)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "3-M365",
                "order": 1
            },
            "stopAll": true
            },
            {
            "name": "Debug in Copilot (Chrome)",
            "configurations": [
                "Launch Remote in Copilot (Chrome)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "3-M365",
                "order": 2
            },
            "stopAll": true
        {{/CEAEnabled}}
        }
    ]
}
