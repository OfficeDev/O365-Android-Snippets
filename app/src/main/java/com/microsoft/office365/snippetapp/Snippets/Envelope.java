package com.microsoft.office365.snippetapp.Snippets;

import com.google.gson.annotations.SerializedName;

public class Envelope <T> {
    @SerializedName("@odata.context")
    public String context;
    public T [] value;
}
