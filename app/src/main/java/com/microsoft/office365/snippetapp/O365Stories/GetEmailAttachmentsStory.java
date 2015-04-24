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
import com.microsoft.outlookservices.Message;

import java.util.Date;
import java.util.List;
import java.util.concurrent.ExecutionException;


public class GetEmailAttachmentsStory extends BaseEmailUserStory {
    public static final String SENT_NOTICE = "Attachment email sent with subject line:";
    public static final boolean IS_INLINE = false;

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
            emailSnippets.addTextFileAttachmentToMessage(emailID
                    , getStringResource(R.string.text_attachment_contents)
                    , getStringResource(R.string.text_attachment_filename)
                    , IS_INLINE);

            String draftMessageID = emailSnippets.getMailMessageById(emailID).getId();

            //UTC time Immediately before message is sent
            Date sendDate = new Date();
            //Send the draft email to the recipient
            emailSnippets.sendMail(draftMessageID);

            //Get the new message
            Message sentMessage = GetAMessageFromInBox(emailSnippets,
                    getStringResource(R.string.mail_subject_text)
                            + uniqueGUID);

            StringBuilder sb = new StringBuilder();
            sb.append(SENT_NOTICE);
            sb.append(getStringResource(R.string.mail_subject_text) + uniqueGUID);
            if (sentMessage.getId().length() > 0) {
                List<Attachment> attachments = emailSnippets.getAttachmentsFromEmailMessage(
                        sentMessage.getId());
                //Send the mail with attachments
                //build string for test results on UI
                for (Attachment attachment : attachments) {
                    if (attachment.getClass().getSimpleName() == "FileAttachment") {
                        FileAttachment fileAttachment = (FileAttachment) attachment;
                        sb.append(fileAttachment.getContentBytes().toString());
                        sb.append("/n");
                    }
                }
                returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);
            }
            else {
                returnResult = StoryResultFormatter.wrapResult(sb.toString(), false);
            }


            //3. Delete the email using the ID
            // Boolean result = emailSnippets.deleteMail(emailID);

        }
        catch (ExecutionException | InterruptedException ex) {
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
