{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.14/MicrosoftTeams.schema.json",
    "manifestVersion": "1.14",
    "version": "1.0.0",
    "id": "${{TEAMS_APP_ID}}",
    "packageName": "com.microsoft.teams.extension",
    "developer": {
        "name": "Teams App, Inc.",
        "websiteUrl": "${{TAB_ENDPOINT}}",
        "privacyUrl": "${{TAB_ENDPOINT}}/index.html#/privacy",
        "termsOfUseUrl": "${{TAB_ENDPOINT}}/index.html#/termsofuse"
    },
    "icons": {
        "color": "resources/color.png",
        "outline": "resources/outline.png"
    },
    "name": {
        "short": "${{TEAMS_APP_NAME}}",
        "full": "Full name for {%appName%}"
    },
    "description": {
        "short": "Short description of {%appName%}",
        "full": "Full description of {%appName%}"
    },
    "accentColor": "#FFFFFF",
    "bots": [
        {
            "botId": "${{BOT_ID}}",
            "scopes": [
                "personal",
                "team",
                "groupchat"
            ],
            "supportsFiles": false,
            "isNotificationOnly": false,
            "commandLists": [
                {
                    "scopes": [
                        "personal",
                        "team",
                        "groupchat"
                    ],
                    "commands": [
                        {
                            "title": "welcome",
                            "description": "Resend welcome card of this Bot"
                        },
                        {
                            "title": "learn",
                            "description": "Learn about Adaptive Card and Bot Command"
                        }
                    ]
                }
            ]
        }
    ],
    "composeExtensions": [],
    "configurableTabs": [
        {
            "configurationUrl": "${{TAB_ENDPOINT}}/index.html#/config",
            "canUpdateConfiguration": true,
            "scopes": [
                "team",
                "groupchat"
            ]
        }
    ],
    "staticTabs": [
        {
            "entityId": "index0",
            "name": "Personal Tab",
            "contentUrl": "${{TAB_ENDPOINT}}/index.html#/tab",
            "websiteUrl": "${{TAB_ENDPOINT}}/index.html#/tab",
            "scopes": [
                "personal"
            ]
        }
    ],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": [
        "${{TAB_DOMAIN}}",
        "${{BOT_DOMAIN}}"
    ]
}