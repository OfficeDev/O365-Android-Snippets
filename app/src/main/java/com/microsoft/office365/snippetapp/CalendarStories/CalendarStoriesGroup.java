package com.microsoft.office365.snippetapp.CalendarStories;

import com.microsoft.office365.snippetapp.helpers.BaseUserStory;

/**
 * Created by johnau on 5/18/2015.
 */
public class CalendarStoriesGroup extends BaseUserStory{
    @Override
    public String execute() {
        return "";
    }

    @Override
    public String getDescription() {

        //Mark this story as a story group list item for UI list
        setGroupingFlag(true);
        return "Calendar stories";
    }
}
