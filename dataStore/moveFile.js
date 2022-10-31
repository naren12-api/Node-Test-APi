const folderConfig = {
    "folderMruConfiguration": {
      "__type": "TargetFolderMruConfiguration:#Exchange",
      "FolderMruEntries": [
        {
          "__type": "TargetFolderMRUEntry:#Exchange",
          "EwsFolderIdEntry": "AQMkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoALgAAAwnRer/eZjZBjmY5EMUoM0MBAC2svI4z/YRKsBSP3F0w8zgAAt5RBJcAAAA="
        },
        {
          "__type": "TargetFolderMRUEntry:#Exchange",
          "EwsFolderIdEntry": "AQMkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoALgAAAwnRer/eZjZBjmY5EMUoM0MBAC2svI4z/YRKsBSP3F0w8zgAAAIBDAAAAA=="
        },
        {
          "__type": "TargetFolderMRUEntry:#Exchange",
          "EwsFolderIdEntry": "AQMkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoALgAAAwnRer/eZjZBjmY5EMUoM0MBAC2svI4z/YRKsBSP3F0w8zgAAAIBTQAAAA=="
        },
        {
          "__type": "TargetFolderMRUEntry:#Exchange",
          "EwsFolderIdEntry": "AQMkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoALgAAAwnRer/eZjZBjmY5EMUoM0MBAC2svI4z/YRKsBSP3F0w8zgAAAIBCgAAAA=="
        }
      ]
    }
  };
  
const folderSwitch = {
    "__type": "ApplyConversationActionJsonRequest:#Exchange",
    "Header": {
      "__type": "JsonRequestHeaders:#Exchange",
      "RequestServerVersion": "V2018_01_08",
      "TimeZoneContext": {
        "__type": "TimeZoneContext:#Exchange",
        "TimeZoneDefinition": {
          "__type": "TimeZoneDefinitionType:#Exchange",
          "Id": "UTC"
        }
      }
    },
    "Body": {
      "__type": "ApplyConversationActionRequest:#Exchange",
      "ConversationActions": [
        {
          "__type": "ConversationAction:#Exchange",
          "Action": "Move",
          "ConversationId": {
            "__type": "ItemId:#Exchange",
            "Id": "AQQkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoAEAB3SppbycBUQ4AVJEm4AoLu"
          },
          "ContextFolderId": {
            "__type": "TargetFolderId:#Exchange",
            "BaseFolderId": {
              "__type": "FolderId:#Exchange",
              "Id": "AQMkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoALgAAAwnRer/eZjZBjmY5EMUoM0MBAC2svI4z/YRKsBSP3F0w8zgAAAIBDAAAAA=="
            }
          },
          "DestinationFolderId": {
            "__type": "TargetFolderId:#Exchange",
            "BaseFolderId": {
              "__type": "FolderId:#Exchange",
              "Id": "AQMkADAwATM3ZmYAZS0yNDljLTE3YmIALTAwAi0wMAoALgAAAwnRer/eZjZBjmY5EMUoM0MBAC2svI4z/YRKsBSP3F0w8zgAAt5RBJcAAAA="
            }
          }
        }
      ],
      "ReturnMovedItemIds": true
    }
  };
  
module.exports = { folderConfig, folderSwitch };