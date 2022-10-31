const {
    sendMailOutlook,
    moveMailOutlook,
    readMailOutlook,
    deleteMailOutlook,
} = require('./outlook.controller')
const routes = require('express').Router();

// Outlook endpoints
routes.post('/api/outlook/sendmail', sendMailOutlook);
routes.post('/api/outlook/move', moveMailOutlook);
routes.post('/api/outlook/read', readMailOutlook);
routes.post('/api/outlook/delete', deleteMailOutlook);

module.exports = { routes };