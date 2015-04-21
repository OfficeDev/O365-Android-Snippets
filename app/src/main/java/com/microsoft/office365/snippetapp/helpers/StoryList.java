/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.helpers;

import android.content.Context;

import com.microsoft.office365.snippetapp.O365Stories.AcceptEventInviteStory;
import com.microsoft.office365.snippetapp.O365Stories.BaseUserStory;
import com.microsoft.office365.snippetapp.O365Stories.CreateOrDeleteContactStory;
import com.microsoft.office365.snippetapp.O365Stories.CreateOrDeleteEventStory;
import com.microsoft.office365.snippetapp.O365Stories.CreateOrDeleteFileStory;
import com.microsoft.office365.snippetapp.O365Stories.CreateOrDeleteOneDriveFolder;
import com.microsoft.office365.snippetapp.O365Stories.DeclineEventInviteStory;
import com.microsoft.office365.snippetapp.O365Stories.DownloadFileStory;
import com.microsoft.office365.snippetapp.O365Stories.EventsFetcherStory;
import com.microsoft.office365.snippetapp.O365Stories.ForwardEmailMessageStory;
import com.microsoft.office365.snippetapp.O365Stories.GetContactsStory;
import com.microsoft.office365.snippetapp.O365Stories.GetEmailAttachmentsStory;
import com.microsoft.office365.snippetapp.O365Stories.GetEmailMessagesStory;
import com.microsoft.office365.snippetapp.O365Stories.GetFilesAndFoldersStory;
import com.microsoft.office365.snippetapp.O365Stories.GetFilteredImportantEvents;
import com.microsoft.office365.snippetapp.O365Stories.ReplyToEmailMessageStory;
import com.microsoft.office365.snippetapp.O365Stories.SendEmailMessageStory;
import com.microsoft.office365.snippetapp.O365Stories.SendEmailWithTextFileAttachmentStory;
import com.microsoft.office365.snippetapp.O365Stories.UpdateContactStory;
import com.microsoft.office365.snippetapp.O365Stories.UpdateEventStory;
import com.microsoft.office365.snippetapp.O365Stories.UpdateFileContentsOnServerStory;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class StoryList {
    /**
     * An array of story scenarios.
     */
    static final Map<String, BaseUserStory> ITEM_MAP = new HashMap<>();
    /**
     * A map of story scenarios, by ID.
     */
    public List<BaseUserStory> ITEMS = new ArrayList<>();

    /*
    * Add stories to the public story list and story map.
     */
    public StoryList(Context context) {
        List<BaseUserStory> baseUserStories = Arrays.asList(
                new GetFilteredImportantEvents(),
                new SendEmailMessageStory(),
                new GetEmailMessagesStory(),
                new ReplyToEmailMessageStory(),
                new ForwardEmailMessageStory(),
                new SendEmailWithTextFileAttachmentStory(),
                new GetEmailAttachmentsStory(),
                new GetContactsStory(),
                new CreateOrDeleteContactStory("CREATE"),
                new CreateOrDeleteContactStory("DELETE"),
                new UpdateContactStory(context),
                new CreateOrDeleteEventStory("CREATE"),
                new CreateOrDeleteEventStory("DELETE"),
                new EventsFetcherStory(),
                new UpdateEventStory(),
                new AcceptEventInviteStory(),
                new DeclineEventInviteStory(),
                new GetFilesAndFoldersStory(),
                new CreateOrDeleteFileStory("CREATE"),
                new UpdateFileContentsOnServerStory(),
                new CreateOrDeleteFileStory("DELETE"),
                new DownloadFileStory(),
                new CreateOrDeleteOneDriveFolder("CREATE"),
                new CreateOrDeleteOneDriveFolder("DELETE")
        );
        for (BaseUserStory object : baseUserStories) {
            ITEM_MAP.put(object.getId(), object);
        }
        ITEMS = baseUserStories;
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
