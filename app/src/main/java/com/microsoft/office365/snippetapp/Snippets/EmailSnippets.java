/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

package com.microsoft.office365.snippetapp.Snippets;

import com.microsoft.outlookservices.Attachment;
import com.microsoft.outlookservices.BodyType;
import com.microsoft.outlookservices.EmailAddress;
import com.microsoft.outlookservices.FileAttachment;
import com.microsoft.outlookservices.Folder;
import com.microsoft.outlookservices.Item;
import com.microsoft.outlookservices.ItemAttachment;
import com.microsoft.outlookservices.ItemBody;
import com.microsoft.outlookservices.Message;
import com.microsoft.outlookservices.Recipient;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutionException;

public class EmailSnippets {

    public static final String MICROSOFT_OUTLOOK_SERVICES_ITEM_ATTACHMENT = "#Microsoft.OutlookServices.ItemAttachment";
    OutlookClient mOutlookClient;

    public EmailSnippets(OutlookClient mailClient) {
        mOutlookClient = mailClient;
    }

    /**
     * Gets a list of the 10 most recent email messages in the
     * user Inbox, sorted by date and time received.
     *
     * @return List of type com.microsoft.outlookservices.Message
     * @version 1.0
     */
    public List<Message> getMailMessages() throws ExecutionException, InterruptedException {
        List<Message> messages = mOutlookClient
                .getMe()
                .getFolders()
                .getById("Inbox")
                .getMessages()
                .select("Subject")
                .orderBy("DateTimeReceived desc")
                .top(10)
                .read().get();
        return messages;

    }

    /**
     * Gets an email message by the id of the desired message
     *
     * @return com.microsoft.outlookservices.Message
     * @version 1.0
     */
    public Message getMailMessageById(String mailId) throws ExecutionException, InterruptedException {
        return mOutlookClient
                .getMe()
                .getMessages()
                .select("ID")
                .getById(mailId)
                .read().get();

    }

    /**
     * Gets a list of all recent email messages in the
     * user Inbox whose subject matches, sorted by date and time received
     *
     * @param subjectLine The subject of the email to be matched
     * @return List of String. The mail Ids of the matching messages
     * @version 1.0
     * @see 'https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar'
     */
    public List<String> getInboxMessagesBySubject(String subjectLine) throws ExecutionException, InterruptedException {
        List<Message> inboxMessages = mOutlookClient
                .getMe()
                .getFolders()
                .getById("Inbox")
                .getMessages()
                .filter("Subject eq '" + subjectLine.trim() + "'")
                .read()
                .get();

        ArrayList<String> mailIds = new ArrayList<>();
        for (Message message : inboxMessages) {
            mailIds.add(message.getId());
        }
        return mailIds;
    }

    /**
     * Gets a list of all recent email messages in the
     * named mail folder whose subject matches, sorted by date and time received
     *
     * @param subjectLine The subject of the email to be matched
     * @param folderName  The display name of the mail folder
     * @return List of String. The mail Ids of the matching messages
     * @version 1.0
     * @see 'https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar'
     */
    public List<Message> getMailboxMessagesByFolderName_Subject(
            String subjectLine,
            String folderName) throws ExecutionException, InterruptedException {

        List<Folder> sentFolder = mOutlookClient.getMe()
                .getFolders()
                .select("ID")
                .filter("DisplayName eq '" + folderName + "'")
                .read()
                .get();
        return mOutlookClient
                .getMe()
                .getFolder(sentFolder.get(0).getId())
                .getMessages()
                .select("ID")
                .filter("Subject eq '" + subjectLine.trim() + "'")
                .read()
                .get();
    }

    /**
     * Gets a list of all recent email messages in the
     * user Inbox whose subject matches subjectLine and sent date&time are greater than the
     * sentDate parameter. Results are sorted by date and time received
     *
     * @param subjectLine The subject of the email to be matched
     * @param sentDate    The UTC (Zulu) time that the mail was sent.
     * @return List of String. The mail Ids of the matching messages
     * @version 1.0
     * @see 'https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar'
     */
    public List<String> GetInboxMessagesBySubject_DateTimeReceived(String subjectLine, Date sentDate, String mailFolder) throws ExecutionException, InterruptedException {

        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
        String filterString = "DateTimeReceived ge "
                + formatter.format(sentDate.getTime())
                + " and "
                + "Subject eq '"
                + subjectLine.trim() + "'";
        List<Message> inboxMessages = mOutlookClient
                .getMe()
                .getFolders()
                .getById(mailFolder)
                .getMessages()
                .select("ID")
                .filter(filterString)
                .read()
                .get();
        ArrayList<String> mailIds = new ArrayList<>();
        for (Message message : inboxMessages) {
            if (message.getSubject().equals(subjectLine.trim()))
                mailIds.add(message.getId());
        }
        return mailIds;
    }


    /**
     * Gets a list of all recent email messages in the
     * user Inbox whose subject matches, sorted by date and time received
     *
     * @param textContent The content of the file to be attached
     * @param fileName    The name of the file to be attached
     * @return com.microsoft.outlookservices.FileAttachment. The Attachment object
     * @version 1.0
     */
    private FileAttachment getTextFileAttachment(String textContent, String fileName) {
        FileAttachment fileAttachment = new FileAttachment();
        fileAttachment.setContentBytes(textContent.getBytes());
        fileAttachment.setName(fileName);
        fileAttachment.setSize(textContent.getBytes().length);
        return fileAttachment;
    }


    /**
     * Gets a message out of the user's draft folder by id and adds a text file attachment
     *
     * @param mailId       The id of the draft email that will get the attachment
     * @param fileContents The contents of the text file to be attached
     * @param fileName     The name of the file to be attached
     * @return Boolean. The result of the operation. True if success
     * @version 1.0
     */
    public Attachment addTextFileAttachmentToMessage(
            String mailId
            , String fileContents
            , String fileName
            , boolean isInline) throws ExecutionException, InterruptedException {


        FileAttachment attachment = getTextFileAttachment(fileContents, fileName);
        attachment.setIsInline(isInline);

        mOutlookClient
                .getMe()
                .getMessages()
                .getById(mailId)
                .getAttachments()
                .add(attachment)
                .get();
        return attachment;
    }

    /**
     * Gets a message out of the user's draft folder by id and adds a text file attachment
     *
     * @param mailId       The id of the draft email that will get the attachment
     * @param itemToAttach The mail message to attach
     * @return Boolean. The result of the operation. True if success
     * @version 1.0
     */
    public Boolean addItemAttachment(
            String mailId
            , Item itemToAttach
            , boolean isInline) throws ExecutionException, InterruptedException {
        ItemAttachment itemAttachment = new ItemAttachment();
        itemAttachment.setName(itemToAttach.getClass().getName());
        itemAttachment.setItem(itemToAttach);
        itemAttachment.setContentType(MICROSOFT_OUTLOOK_SERVICES_ITEM_ATTACHMENT);
        itemAttachment.setIsInline(false);
        itemAttachment.setId(itemToAttach.getId());
        itemAttachment.setIsInline(isInline);
        mOutlookClient
                .getMe()
                .getMessages()
                .getById(mailId)
                .getAttachments()
                .add(itemAttachment)
                .get();
        return true;
    }

    /**
     * Gets a list of Attachment objects representing the contents of a set of email attachments
     *
     * @param mailID The email id of the message whose attachments are wanted
     * @return List. A list of Byte array objects
     * @version 1.0
     */
    public List<Attachment> getAttachmentsFromEmailMessage(String mailID) throws ExecutionException, InterruptedException {
        return mOutlookClient
                .getMe()
                .getMessages()
                .getById(mailID)
                .getAttachments()
                .read()
                .get();
    }

    /**
     * Gets a message out of the user's draft folder by id and adds a text file attachment
     *
     * @param emailAddress The email address of the mail recipient
     * @param subject      The subject of the email
     * @param body         The body of the email
     * @return String. The id of the email added to the draft folder
     * @version 1.0
     */
    public String addDraftMail(
            final String emailAddress
            , final String subject
            , final String body) throws ExecutionException, InterruptedException {
        // Prepare the message.
        List<Recipient> recipientList = new ArrayList<>();

        Recipient recipient = new Recipient();
        EmailAddress email = new EmailAddress();
        email.setAddress(emailAddress);
        recipient.setEmailAddress(email);
        recipientList.add(recipient);

        Message messageToSend = new Message();
        messageToSend.setToRecipients(recipientList);

        ItemBody bodyItem = new ItemBody();
        bodyItem.setContentType(BodyType.HTML);
        bodyItem.setContent(body);
        messageToSend.setBody(bodyItem);
        messageToSend.setSubject(subject);

        // Contact the Office 365 service and try to add the message to
        // the draft folder.
        Message draft = mOutlookClient
                .getMe()
                .getMessages()
                .add(messageToSend)
                .get();

        return draft.getId();
    }

    /**
     * Sends the Exchange server copy of a new mail message
     *
     * @param mailId The email to be sent from the draft folder
     * @return Boolean. The result of the operation
     * @version 1.0
     */
    public Boolean sendMail(String mailId) throws ExecutionException, InterruptedException {
        mOutlookClient
                .getMe()
                .getMessages()
                .getById(mailId)
                .getOperations()
                .send();
        return true;
    }

    /**
     * Gets a message out of the user's draft folder by id and adds a text file attachment
     *
     * @param emailAddress The email address of the mail recipient
     * @param subject      The subject of the email
     * @param body         The body of the email
     * @return String. The id of the sent email
     * @version 1.0
     */
    public String createAndSendMail(
            final String emailAddress
            , final String subject
            , final String body) throws ExecutionException, InterruptedException {

        // Prepare the message.
        List<Recipient> recipientList = new ArrayList<>();

        Recipient recipient = new Recipient();
        EmailAddress email = new EmailAddress();
        email.setAddress(emailAddress);
        recipient.setEmailAddress(email);
        recipientList.add(recipient);

        Message messageToSend = new Message();
        messageToSend.setToRecipients(recipientList);

        ItemBody bodyItem = new ItemBody();
        bodyItem.setContentType(BodyType.HTML);
        bodyItem.setContent(body);
        messageToSend.setBody(bodyItem);
        messageToSend.setSubject(subject);

        // Contact the Office 365 service and try to deliver the message.
        Message draft = mOutlookClient
                .getMe()
                .getMessages()
                .select("ID")
                .add(messageToSend)
                .get();
        mOutlookClient.getMe()
                .getOperations()
                .sendMail(draft, false)
                .get();
        return draft.getId();
    }

    /**
     * Forwards a message out of the user's Inbox folder by id
     *
     * @param emailId The id of the mail to be forwarded
     * @return String. The id of the sent email
     * @version 1.0
     */
    public String forwardMail(String emailId) throws ExecutionException, InterruptedException {
        Message forwardMessage = mOutlookClient
                .getMe()
                .getMessages()
                .getById(emailId)
                .getOperations()
                .createForward()
                .get();
        Message message = getDraftMessageMap().get(forwardMessage.getConversationId());
        if (message == null) {
            return "";
        }
        return message.getId();
    }

    /**
     * Deletes a message out of the user's Sent folder by id
     *
     * @param emailID The id of the mail to be deleted
     * @return Boolean. The result of the operation
     * @version 1.0
     */
    public Boolean deleteMail(String emailID) throws ExecutionException, InterruptedException {
        mOutlookClient
                .getMe()
                .getMessages()
                .getById(emailID)
                .delete()
                .get();

        return true;
    }

    /**
     * Generates a hash table whose key is a mail Id and value is corresponding
     * mail message
     *
     * @return Map of type String, Message. The result of the operation
     * @version 1.0
     */
    public Map<String, Message> getDraftMessageMap() throws ExecutionException, InterruptedException {
        Map<String, Message> draftMessageMap = new HashMap<>();
        for (Message draftMessage : getDraftMessages()) {
            draftMessageMap.put(draftMessage.getId(), draftMessage);
        }
        return draftMessageMap;
    }

    /**
     * Gets a List of type Message representing the contents of
     * the user's email drafts folder
     *
     * @return List. The result of the operation
     * @version 1.0
     */
    public List<Message> getDraftMessages() throws ExecutionException, InterruptedException {
        return mOutlookClient
                .getMe()
                .getFolder("Drafts")
                .getMessages()
                .select("ID")
                .read()
                .get();
    }

    /**
     * Forwards a message out of the user's Inbox folder by id
     *
     * @param emailId     The id of the mail to be forwarded
     * @param messageBody The body of the message as a string
     * @return String. The id of the sent email
     * @version 1.0
     */
    public String replyToEmailMessage(String emailId, String messageBody)
            throws
            ExecutionException
            , InterruptedException {


        //Create a new message in the user draft items folder
        Message replyEmail = mOutlookClient
                .getMe()
                .getFolder("Draft")
                .getMessages()
                .getById(emailId)
                .getOperations()
                .createReply()
                .get();

        if (replyEmail != null) {
            //Create a message subject body and set in the reply message
            ItemBody bodyItem = new ItemBody();
            bodyItem.setContentType(BodyType.HTML);
            bodyItem.setContent(messageBody);
            replyEmail.setBody(bodyItem);

            // Send the email reply
            mOutlookClient
                    .getMe()
                    .getOperations()
                    .sendMail(replyEmail, false)
                    .get();

            return replyEmail.getId();
        } else {
            return "";
        }

    }

}
// *********************************************************
//
// O365-Android-Snippets, https://github.com/OfficeDev/O365-Android-Snippets
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
