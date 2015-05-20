/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.ContactStories;

import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.Snippets.ContactsSnippets;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Contact;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetFilteredContactsWithSurnameStory extends BaseUserStory {

    private static final String STORY_DESCRIPTION = "Gets contacts filtered by surname";
    private static final String CONTACT_WITH_SURNAME = "Contact with surname";

    @Override
    public String execute() {
        boolean isStoryComplete = false;
        String storyResultText = "";

        String surname = getStringResource(R.string.contacts_last_name);
        AuthenticationController
                .getInstance()
                .setResourceId(getO365MailResourceId());
        ContactsSnippets contactsSnippets = new ContactsSnippets(getO365MailClient());

        try {
            //Create a contact that we can test against
            String contactId = contactsSnippets.createContact(
                    getStringResource(R.string.contacts_email),
                    getStringResource(R.string.contacts_business_phone),
                    getStringResource(R.string.contacts_home_phone),
                    getStringResource(R.string.contacts_first_name),
                    surname);

            //Find the new test contact
            List<Contact> contacts = contactsSnippets.getContactsWithSurname(surname);
            for (Contact contact : contacts) {
                if (contact.getSurname().equals(surname)) {
                    break;
                }
            }
            //Delete the test contact from tenant
            contactsSnippets.deleteContact(contactId);

            //Story is completed
            storyResultText = StoryResultFormatter.wrapResult(CONTACT_WITH_SURNAME
                    + " " + surname
                    + " found.", true);
        } catch (ExecutionException | InterruptedException e) {
            storyResultText = FormatException(e, STORY_DESCRIPTION);
        }
        return storyResultText;
    }

    @Override
    public String getDescription() {
        return STORY_DESCRIPTION;
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
