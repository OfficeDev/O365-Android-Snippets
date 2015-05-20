/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.ContactStories;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.ContactsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryAction;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;

import java.util.concurrent.ExecutionException;

//This story handles both of the following stories that appear in the UI list
// based on strings passed in the constructor...
//- Create a contact (which is then deleted for cleanup)
//- Delete a contact (which is created first and then deleted)
public class CreateOrDeleteContactStory extends BaseUserStory {
    private final String CREATE_DESCRIPTION = "Creates a new contact";
    private final String CREATE_TAG = "Create contact story";
    private final String CREATE_SUCCESS = "CreateContactStory: Contact created.";
    private final String CREATE_ERROR = "Create contact exception: ";
    private final String DELETE_DESCRIPTION = "Deletes a contact";
    private final String DELETE_TAG = "Delete contact story";
    private final String DELETE_SUCCESS = "DeleteContactStory: Contact deleted.";
    private final String DELETE_ERROR = "Delete contact exception: ";

    private String mDescription;
    private String mLogTag;
    private String mSuccessDescription;
    private String mErrorDescription;

    public CreateOrDeleteContactStory(StoryAction action) {
        switch (action) {
            case CREATE: {
                mDescription = CREATE_DESCRIPTION;
                mLogTag = CREATE_TAG;
                mSuccessDescription = CREATE_SUCCESS;
                mErrorDescription = CREATE_ERROR;
                break;
            }
            case DELETE: {
                mDescription = DELETE_DESCRIPTION;
                mLogTag = DELETE_TAG;
                mSuccessDescription = DELETE_SUCCESS;
                mErrorDescription = DELETE_ERROR;
                break;
            }
        }
    }

    @Override
    public String execute() {
        String returnValue = StoryResultFormatter.wrapResult(
                mLogTag, false
        );
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        ContactsSnippets contactsSnippets = new ContactsSnippets(getO365MailClient());
        try {
            String contactId = contactsSnippets.createContact(
                    getStringResource(R.string.contacts_email),
                    getStringResource(R.string.contacts_business_phone),
                    getStringResource(R.string.contacts_home_phone),
                    getStringResource(R.string.contacts_first_name),
                    getStringResource(R.string.contacts_last_name));
            if (contactId.length() > 0) {
                contactsSnippets.deleteContact(contactId);
                return StoryResultFormatter.wrapResult(
                        mSuccessDescription, true
                );
            }
        } catch (ExecutionException | InterruptedException e) {
            return FormatException(e, mDescription);
        }
        return returnValue;
    }

    @Override
    public String getDescription() {
        return mDescription;
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
