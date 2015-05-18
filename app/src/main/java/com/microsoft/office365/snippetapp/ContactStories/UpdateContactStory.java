/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.ContactStories;

import android.content.Context;
import android.util.Log;

import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.ContactsSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Contact;

import java.util.concurrent.ExecutionException;

public class UpdateContactStory extends BaseUserStory {
    private Context mContext;

    public UpdateContactStory(Context context) {
        mContext = context;
    }

    @Override
    public String execute() {
        String returnValue = "CreateContactStory failed";
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        ContactsSnippets contactsSnippets = new ContactsSnippets(getO365MailClient());
        try {
            //Create a new contact to test update API with
            String contactId = contactsSnippets.createContact(
                    mContext.getString(R.string.contacts_email),
                    mContext.getString(R.string.contacts_business_phone),
                    mContext.getString(R.string.contacts_home_phone),
                    mContext.getString(R.string.contacts_first_name),
                    mContext.getString(R.string.contacts_last_name)
            );
            if (contactId.length() > 0) {
                //Set a new surname value on contact to update
                //and update the contact
                String newSurname = "Curtis_2";

                contactsSnippets.updateContact(
                        contactId
                        , mContext.getString(R.string.contacts_first_name)
                        , newSurname
                );

                //Get the updated contact from server and compare surname with
                //the intended update surname value
                Contact updatedContact = contactsSnippets.getAContact(contactId);
                if (updatedContact != null) {
                    String testSurName = updatedContact.getSurname();
                    if (testSurName.length() > 0 && testSurName.equals(newSurname)) {
                        returnValue = StoryResultFormatter.wrapResult(
                                "UpdateContactStory: Contact "
                                        + " updated."
                                , true
                        );
                    } else {
                        returnValue = StoryResultFormatter.wrapResult(
                                "UpdateContactStory: Contact "
                                , false
                        );

                    }
                } else {
                    returnValue = StoryResultFormatter.wrapResult(
                            "UpdateContactStory: Contact "
                            , false
                    );

                }
                //Clean up state
                contactsSnippets.deleteContact(contactId);
                return returnValue;
            }
        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Update Contact story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Update contacts exception: "
                            + formattedException
                    , false
            );

        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Update Contact story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Update contacts exception: "
                            + formattedException
                    , false
            );
        }
        return returnValue;
    }

    @Override
    public String getDescription() {
        return "Updates the surname of a contact";
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
