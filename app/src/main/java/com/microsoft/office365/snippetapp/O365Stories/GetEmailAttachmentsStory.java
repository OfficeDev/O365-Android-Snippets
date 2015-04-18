package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Attachment;
import com.microsoft.outlookservices.FileAttachment;

import java.util.Date;
import java.util.List;

/**
 * Created by johnaustin on 4/15/15.
 */
public class GetEmailAttachmentsStory extends BaseUserStory{
    public static final String SENT_NOTICE = "Attachment email sent with subject line:";
    public static final int MAX_TRY_COUNT = 20;

    @Override
    public String execute() {
        String returnResult = "";
        try {
            AuthenticationController
                    .getInstance()
                    .setResourceId(
                            getO365MailResourceId());

            EmailSnippets emailSnippets = new EmailSnippets(
                    getO365MailClient());
            //1. Send an email and store the ID
            String uniqueGUID = java.util.UUID.randomUUID().toString();
            String mailSubject = getStringResource(R.string.mail_subject_text) + uniqueGUID;

            //Add a new email to the user's draft folder
            String emailID = emailSnippets.addDraftMail(GlobalValues.USER_EMAIL,
                    mailSubject,
                    getStringResource(R.string.mail_body_text));

            //Add a text file attachment to the mail added to the draft folder
            emailSnippets.addAttachmentToMessage(emailID
                    , getStringResource(R.string.text_attachment_contents)
                    , getStringResource(R.string.text_attachment_filename));

            String draftMessageID = emailSnippets.getMailMessageById(emailID).getId();

            //UTC time Immediately before message is sent
            Date sendDate = new Date();
            //Send the draft email to the recipient
            emailSnippets.sendMail(draftMessageID);

            //Get the new message
            String emailId = "";
            int tryCount = 0;

            //Try to get the newly sent email from user's inbox at least once.
            //continue trying to get the email while the email is not found
            //and the loop has tried less than 50 times.
            do {
                List<String> mailIds = emailSnippets
                        .GetInboxMessagesBySubject_DateTimeReceived(
                                mailSubject
                                , sendDate
                                ,getStringResource(R.string.Email_Folder_Inbox));
                if (mailIds.size() > 0) {
                    emailId = mailIds.get(0);
                }
                tryCount++;


                //Stay in loop while these conditions are true.
                //If either condition becomes false, break
            } while (emailId.length() == 0 && tryCount < MAX_TRY_COUNT);

            StringBuilder sb = new StringBuilder();
            sb.append(SENT_NOTICE);
            sb.append(getStringResource(R.string.mail_subject_text) + uniqueGUID);
            if (emailId.length() > 0)
            {
                List<Attachment> attachments = emailSnippets.getAttachmentsFromEmailMessage(emailId);
                //Send the mail with attachments
                //build string for test results on UI
                for (Attachment attachment : attachments)
                {
                    if (attachment.getClass().getSimpleName() == "FileAttachment") {
                        FileAttachment fileAttachment = (FileAttachment) attachment;
                        sb.append(fileAttachment.getContentBytes().toString());
                        sb.append("/n");
                    }
                }
                returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);
            }
            else
                returnResult = StoryResultFormatter.wrapResult(sb.toString(), false);


            //3. Delete the email using the ID
            // Boolean result = emailSnippets.deleteMail(emailID);

        } catch (Exception ex) {
            String formattedException = APIErrorMessageHelper.getErrorMessage(ex.getMessage());
            Log.e("Send email story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Send mail exception: "
                            + formattedException
                    , false
            );

        }
        return returnResult;

    }

    @Override
    public String getDescription() {
        return "Gets the attachments from an email message";
    }
}
