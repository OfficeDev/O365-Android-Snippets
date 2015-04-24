package com.microsoft.office365.snippetapp.O365Stories;

import android.util.Log;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;


public class SendEmailWithTextFileAttachmentStory extends BaseEmailUserStory {

    public static final String STORY_DESCRIPTION = "Sends an email message with a text file attachment";
    public static final String SENT_NOTICE = "Email sent with subject line:";
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

            //Add a new email to the user's draft folder
            String emailID = emailSnippets.addDraftMail(GlobalValues.USER_EMAIL,
                    getStringResource(R.string.mail_subject_text) + uniqueGUID,
                    getStringResource(R.string.mail_body_text));

            //Add a text file attachment to the mail added to the draft folder
            emailSnippets.addTextFileAttachmentToMessage(emailID
                    , getStringResource(R.string.text_attachment_contents)
                    , getStringResource(R.string.text_attachment_filename)
                    , IS_INLINE);

            String draftMessageID = emailSnippets.getMailMessageById(emailID).getId();

            //Send the mail with attachments
            //build string for test results on UI
            StringBuilder sb = new StringBuilder();
            sb.append(SENT_NOTICE);
            sb.append(getStringResource(R.string.mail_subject_text) + uniqueGUID);
            returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);

            //Send the draft email to the recipient
            emailSnippets.sendMail(draftMessageID);

        }
        catch (Exception ex) {
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
        return STORY_DESCRIPTION;
    }
}
