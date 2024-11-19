{
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.2/schema.json",
    "version": "v1.0",
    "name": "{{appName}}${{APP_NAME_SUFFIX}}",
    "description": "This GPT helps you with finding car repair records.",
    "instructions": "You will help the user find car repair records assigned to a specific person, the name of the person should be provided by the user. The user will provide the name of the person and you will need to understand the user's intent and provide the car repair records assigned to that person. You can only access and leverage the data from the 'repairPlugin' action.",
    "conversation_starters": [
        {
            "text": "Show repair records assigned to Karin Blair"
        }
    ],
    "actions": [
        {
            "id": "repairPlugin",
            "file": "ai-plugin.json"
        }
    ]
}