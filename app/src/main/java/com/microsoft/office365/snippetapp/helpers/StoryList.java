/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */
package com.microsoft.office365.snippetapp.helpers;

import android.content.Context;

import com.microsoft.office365.snippetapp.CalendarStories.CreateOrDeleteEventStory;
import com.microsoft.office365.snippetapp.CalendarStories.CreateRecurringEventStory;
import com.microsoft.office365.snippetapp.CalendarStories.EventsFetcherStory;
import com.microsoft.office365.snippetapp.CalendarStories.GetFilteredImportantEvents;
import com.microsoft.office365.snippetapp.CalendarStories.RespondToCalendarEventInviteStory;
import com.microsoft.office365.snippetapp.CalendarStories.UpdateEventStory;
import com.microsoft.office365.snippetapp.ContactStories.CreateOrDeleteContactStory;
import com.microsoft.office365.snippetapp.ContactStories.GetContactsStory;
import com.microsoft.office365.snippetapp.ContactStories.GetFilteredContactsWithSurnameStory;
import com.microsoft.office365.snippetapp.ContactStories.UpdateContactStory;
import com.microsoft.office365.snippetapp.EmailStories.ForwardEmailMessageStory;
import com.microsoft.office365.snippetapp.EmailStories.GetEmailAttachmentsStory;
import com.microsoft.office365.snippetapp.EmailStories.GetEmailMessagesStory;
import com.microsoft.office365.snippetapp.EmailStories.ReplyToEmailMessageStory;
import com.microsoft.office365.snippetapp.EmailStories.SendEmailMessageStory;
import com.microsoft.office365.snippetapp.EmailStories.SendEmailWithMessageAttachStory;
import com.microsoft.office365.snippetapp.EmailStories.SendEmailWithTextFileAttachmentStory;
import com.microsoft.office365.snippetapp.FileFolderStories.CreateOrDeleteFileStory;
import com.microsoft.office365.snippetapp.FileFolderStories.CreateOrDeleteOneDriveFolder;
import com.microsoft.office365.snippetapp.FileFolderStories.DownloadFileStory;
import com.microsoft.office365.snippetapp.FileFolderStories.GetFilesAndFoldersStory;
import com.microsoft.office365.snippetapp.FileFolderStories.UpdateFileContentsOnServerStory;
import com.microsoft.office365.snippetapp.R;
import com.microsoft.office365.snippetapp.UserGroupStories.GetADGroupsStory;
import com.microsoft.office365.snippetapp.UserGroupStories.GetADUsersStory;
import com.microsoft.office365.snippetapp.UserGroupStories.GetTenantDetailsStory;

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
                //Active Directory Stories
                new StoryGroupPlaceholder(context.getString(R.string.active_directory_group_placeholder)),
                new GetADUsersStory(),
                new GetTenantDetailsStory(),
                new GetADGroupsStory(),

                //Email Stories
                new StoryGroupPlaceholder(context.getString(R.string.email_group_placeholder)),
                new SendEmailMessageStory(),
                new GetEmailMessagesStory(),
                new ReplyToEmailMessageStory(),
                new ForwardEmailMessageStory(),
                new SendEmailWithTextFileAttachmentStory(),
                new SendEmailWithMessageAttachStory(),
                new GetEmailAttachmentsStory(),

                //Contact Stories
                new StoryGroupPlaceholder(context.getString(R.string.contact_group_placeholder)),
                new GetContactsStory(),
                new CreateOrDeleteContactStory(StoryAction.CREATE),
                new CreateOrDeleteContactStory(StoryAction.DELETE),
                new UpdateContactStory(context),
                new GetFilteredContactsWithSurnameStory(),

                //Calendar Stories
                new StoryGroupPlaceholder(context.getString(R.string.calendar_group_placeholder)),
                new CreateOrDeleteEventStory(StoryAction.CREATE),
                new CreateOrDeleteEventStory(StoryAction.DELETE),
                new EventsFetcherStory(),
                new UpdateEventStory(),
                new RespondToCalendarEventInviteStory(),
                new GetFilteredImportantEvents(),
                new CreateRecurringEventStory(),

                //Files and Folders Stories
                new StoryGroupPlaceholder(context.getString(R.string.file_folder_group_placeholder)),
                new GetFilesAndFoldersStory(),
                new CreateOrDeleteFileStory(StoryAction.CREATE),
                new UpdateFileContentsOnServerStory(),
                new CreateOrDeleteFileStory(StoryAction.DELETE),
                new DownloadFileStory(),
                new CreateOrDeleteOneDriveFolder(StoryAction.CREATE),
                new CreateOrDeleteOneDriveFolder(StoryAction.DELETE)
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
