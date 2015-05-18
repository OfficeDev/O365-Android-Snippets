/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.ContactStories;

import android.util.Log;

import com.microsoft.office365.snippetapp.helpers.BaseUserStory;
import com.microsoft.office365.snippetapp.Snippets.ContactsSnippets;
import com.microsoft.office365.snippetapp.helpers.APIErrorMessageHelper;
import com.microsoft.office365.snippetapp.helpers.AuthenticationController;
import com.microsoft.office365.snippetapp.helpers.StoryResultFormatter;
import com.microsoft.outlookservices.Contact;

import java.util.List;
import java.util.concurrent.ExecutionException;

public class GetContactsStory extends BaseUserStory {

    @Override
    public String execute() {
        AuthenticationController
                .getInstance()
                .setResourceId(
                        getO365MailResourceId());

        ContactsSnippets contactsSnippets = new ContactsSnippets(
                getO365MailClient());
        try {
            List<Contact> contacts = contactsSnippets.getContacts(11);
            //build string for test results on UI
            StringBuilder sb = new StringBuilder();
            sb.append("Gets contacts: "
                    + contacts.size()
                    + " contacts returned");
            sb.append("\n");
            for (Contact contact : contacts) {
                sb.append("\t\t");
                sb.append(contact.getDisplayName());
                sb.append("\n");
            }
            return StoryResultFormatter.wrapResult(sb.toString(), true);

        } catch (ExecutionException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Get Contacts story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Get contacts exception: "
                            + formattedException, false
            );
        } catch (InterruptedException e) {
            e.printStackTrace();
            String formattedException = APIErrorMessageHelper.getErrorMessage(e.getMessage());
            Log.e("Get Contacts story", formattedException);
            return StoryResultFormatter.wrapResult(
                    "Get contacts exception: "
                            + formattedException, false
            );
        }
    }

    @Override
    public String getDescription() {

        return "Gets your contacts";
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
