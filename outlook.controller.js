const { doRequest } = require('./outlook.service');

const body = require('./dataStore/data.json');
const readBody = require('./dataStore/read.json');
const deleteBody= require('./dataStore/delete.json');
const { folderConfig, folderSwitch } = require('./dataStore/moveFile');

const sendMailOutlook = async (req, res) => {
    const recipients = [];
    const item = body.Body.Items;
    item[0]['Subject'] = req.body.sub;
    if(req.body.to && Array.isArray(req.body.to)) {
        req.body.to.map((item) => {
            return recipients.push({
                EmailAddress: item,
                Name: item,
                MailboxType: "ExternalMailbox",
                RoutingType: "SMTP"
            });
        })
    }
    else {
        recipients.push({
            EmailAddress: req.body.to,
            Name: req.body.to,
            MailboxType: "ExternalMailbox",
            RoutingType: "SMTP"
        });
    }
    item[0]['ToRecipients'] = recipients;
    item[0]['CcRecipients'] = req.body.cc || [];
    item[0]['BccRecipients'] = req.body.bcc || [];
    item[0]['From']['Mailbox']['EmailAddress'] = req.body.from;
    item[0]['From']['Mailbox']['Name'] = req.body.from;
    item[0]['Body'] = {...item[0]['Body'], Value: `<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"><style type=\"text/css\" style=\"display:none;\"> P {margin-top:0;margin-bottom:0;} </style></head><body dir=\"ltr\"><div style=\"font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);\" class=\"elementToProof\">${req.body.message}</div></body></html>`};
    const mailSend = {
        ...body, 
        Body: { 
            ...body.Body,
            Items: item
        }
    };
    const options = {
        qs: { action: 'CreateItem' },
        headers: { action: "CreateItem" },
        body: JSON.stringify(mailSend)
    }
    try {
        let response = await doRequest(options)
        res.status(200).json({ success: true, msg: "mail send successfully", body: JSON.stringify(response) })
    } catch (err) {
        // console.log("response catch", err)
        res.status(500).json({ success: false, msg: err })
    }
}

const moveMailOutlook = async (req, res) => {
    const ConversationActions = folderSwitch.Body.ConversationActions;
    ConversationActions[0]['ConversationId']['id'] = req.body.ConversationId;
    ConversationActions[0]['ContextFolderId']['BaseFolderId']['Id'] = req.body.BaseFolderId;
    ConversationActions[0]['DestinationFolderId']['BaseFolderId']['Id'] = req.body.DestinationFolderId;

    const sendFolder = { 
        ...folderSwitch, 
        Body: { ...folderSwitch.Body, ConversationActions }
    };
    const options1 = {
        qs: { action: 'SetFolderMruConfiguration' },
        headers:
        {
            "action": "SetFolderMruConfiguration",
            "x-owa-urlpostdata": JSON.stringify(folderConfig)
        },
        body: null
    }
    const options2 = {
        qs: { action: 'ApplyConversationAction' },
        headers:
        {
            "action": "ApplyConversationAction",
            "x-owa-urlpostdata": JSON.stringify(sendFolder)
        },
        body: null
    }
    try {
        await doRequest(options1);
        let response = await doRequest(options2);
        res.status(200).json({ success: true, msg: "moved mail successfully", body: JSON.parse(response) })
    } catch (err) {
        console.log("response catch", err)
        res.status(500).json({ success: false, msg: err })
    }
}

const readMailOutlook = async (req, res) => {
    const readMail = {...readBody};
    readMail['Body']['ConversationActions'][0]['ConversationId']['Id'] = req.body.ConversationId;
    readMail['Body']['ConversationActions'][0]['ContextFolderId']['BaseFolderId']['Id'] = req.body.BaseFolderId;
    // console.log("caacccccccccc")
    const options = {
        qs: { action: 'ApplyConversationAction' },
        headers:
        {
            "action": "ApplyConversationAction",
            "x-owa-urlpostdata": JSON.stringify(readMail)
        },
        body: null
    }
    try {
        let response = await doRequest(options)
        res.status(200).json({ success: true, msg: "read mail successfully", body: JSON.parse(response) })
    } catch (err) {
        console.log("response catch", err)
        res.status(500).json({ success: false, msg: err })
    }
}

const deleteMailOutlook = async (req, res) => {
    console.log("deleteMailOutlook")
    /* "SoftDelete", */
    const options = {
        qs: { action: 'ApplyConversationAction' },
        headers:
        {
            "action": "ApplyConversationAction",
            "x-owa-urlpostdata": JSON.stringify(deleteBody)
        },
        body: null
    }
    try {
        let response = await doRequest(options)
        res.status(200).json({ success: true, msg: "get mail successfully", body: JSON.parse(response) })
    } catch (err) {
        console.log("response catch", err)
        res.status(500).json({ success: false, msg: err })
    }
}

module.exports = {
    sendMailOutlook,
    moveMailOutlook,
    readMailOutlook,
    deleteMailOutlook,
}