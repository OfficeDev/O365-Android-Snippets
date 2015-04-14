/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

package com.microsoft.office365.snippetapp.Snippets;

import com.microsoft.outlookservices.BodyType;
import com.microsoft.outlookservices.EmailAddress;
import com.microsoft.outlookservices.FileAttachment;
import com.microsoft.outlookservices.ItemBody;
import com.microsoft.outlookservices.Message;
import com.microsoft.outlookservices.Recipient;
import com.microsoft.outlookservices.odata.OutlookClient;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutionException;

public class EmailSnippets {
    private final static int pageSize = 11;
    OutlookClient mMailClient;

    public EmailSnippets(OutlookClient mailClient) {
        mMailClient = mailClient;
    }

    public List<Message> getMailMessages() throws ExecutionException, InterruptedException {
        List<Message> messages = mMailClient
                .getMe()
                .getFolders().getById("Inbox")
                .getMessages()
                .orderBy("DateTimeReceived desc")
                .top(10)
                .read().get();
        return messages;

    }

    public List<String> GetInboxMessagesBySubject(String subjectLine) throws ExecutionException, InterruptedException {
        List<Message> inboxMessages = mMailClient
                .getMe()
                .getFolders()
                .getById("Inbox")
                .getMessages()
                .filter("Subject eq '" + subjectLine + "'")
                .read()
                .get();

        ArrayList<String> mailIds = new ArrayList<>();
        for (Message message : inboxMessages) {
            mailIds.add(message.getId());
        }
        return mailIds;
    }

    //Creates a FileAttachment object with given contents and file name
    private FileAttachment getTextFileAttachment(String textContent, String fileName)
    {
        FileAttachment fileAttachment = new FileAttachment();
        fileAttachment.setContentBytes(textContent.getBytes());
        fileAttachment.setName(fileName);
        fileAttachment.setContentType("Text");
        fileAttachment.setSize(textContent.getBytes().length);
        return fileAttachment;
    }

    //Gets a message out of the user's draft folder by id and adds a text file attachment
    public Boolean addAttachmentToDraft(String mailId, String fileContents, String fileName) throws ExecutionException, InterruptedException {

        mMailClient
                .getMe()
                .getMessages()
                .getById(mailId)
                .getAttachments()
                .add(getTextFileAttachment(fileContents,fileName))
                .get();
        return true;
    }

    //Adds a new mail message to the user's draft folder and returns the id of
    //the new message
    public String addDraftMail (final String emailAddress
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
        Message draft = mMailClient
                .getMe()
                .getFolders()
                .getById("Drafts")
                .getMessages()
                .add(messageToSend)
                .get();

        return draft.getId();
    }

    public Boolean sendDraftMail(String mailID) throws ExecutionException, InterruptedException
    {
        //Get a message out of user's draft folder by mail Id
        Message draft = mMailClient.getMe()
                .getFolders()
                .getById("Drafts")
                .getMessages()
                .getById(mailID)
                .read().get();

        //Send the draft message
        mMailClient.getMe()
                .getOperations()
                .sendMail(draft, false)
                .get();
        return true;
    }

    public String sendMail(
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
        Message draft = mMailClient
                .getMe()
                .getMessages()
                .add(messageToSend)
                .get();
        mMailClient.getMe()
                .getOperations()
                .sendMail(draft, false)
                .get();
        return draft.getId();
    }

    public String forwardMail(String emailId) throws ExecutionException, InterruptedException {
        Message forwardMessage = mMailClient
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

    public Boolean deleteMail(String emailID) throws ExecutionException, InterruptedException {
        mMailClient
                .getMe()
                .getMessages()
                .getById(emailID)
                .delete()
                .get();

        return true;
    }

    public Map<String, Message> getDraftMessageMap() throws ExecutionException, InterruptedException {
        Map<String, Message> draftMessageMap = new HashMap<>();
        for (Message draftMessage : getDraftMessages()) {
            draftMessageMap.put(draftMessage.getId(), draftMessage);
        }
        return draftMessageMap;
    }

    public List<Message> getDraftMessages() throws ExecutionException, InterruptedException {
        return mMailClient
                .getMe()
                .getFolder("Drafts")
                .getMessages()
                .read()
                .get();
    }

    public String replyToEmailMessage(String emailId, String messageBody)
            throws
            ExecutionException
            , InterruptedException {


        //Create a new message in the user draft items folder
        Message replyEmail = mMailClient
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
            mMailClient
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
