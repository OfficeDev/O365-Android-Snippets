package com.microsoft.office365.snippetapp.O365Stories;

import android.content.Context;
import android.util.Log;

import com.microsoft.office365.snippetapp.AndroidSnippetsApplication;
import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.EmailSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.GlobalValues;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

/**
 * Created by johnaustin on 4/14/15.
 */
public class GetEmailAttachments extends  BaseUserStory {

    public static final String STORY_DESCRIPTION = "Gets the attachments for an email message";
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
            String emailID = emailSnippets.sendMail(GlobalValues.USER_EMAIL,
                    AndroidSnippetsApplication
                            .getApplication()
                            .getApplicationContext()
                            .getString(R.string.mail_subject_text) + uniqueGUID,
                    AndroidSnippetsApplication
                            .getApplication()
                            .getApplicationContext()
                            .getString(R.string.mail_body_text));

            //3. Delete the email using the ID
            Boolean result = emailSnippets.deleteMail(emailID);

            //build string for test results on UI
            StringBuilder sb = new StringBuilder();
            sb.append("Email is added");
            returnResult = StoryResultFormatter.wrapResult(sb.toString(), true);
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
        return STORY_DESCRIPTION;
    }
}
