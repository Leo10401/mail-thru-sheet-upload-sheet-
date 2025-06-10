const mongoose = require('mongoose');

const sendedMailSchema = new mongoose.Schema({
    recipientEmail: {
        type: String,
        required: true
    },
    recipientName: {
        type: String,
        required: true
    },
    certificateDetails: {
        type: String,
        required: true
    },
    sheetName: {
        type: String,
        required: true
    },
    fileName: {
        type: String,
        required: true
    },
    sentAt: {
        type: Date,
        default: Date.now
    }
});

const SendedMail = mongoose.model('SendedMail', sendedMailSchema);

module.exports = SendedMail;
