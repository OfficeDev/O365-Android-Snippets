package com.microsoft.office365.snippetapp.helpers;

import com.google.gson.GsonBuilder;
import com.google.gson.JsonDeserializationContext;
import com.google.gson.JsonElement;
import com.google.gson.JsonParseException;
import com.google.gson.JsonPrimitive;
import com.google.gson.JsonSerializationContext;

import org.joda.time.DateTime;

import java.lang.reflect.Type;

public class GsonDateTime {
    public  static GsonBuilder getDirectoryServiceBuilder() {
        GsonBuilder gsonBuilder = new GsonBuilder();
        gsonBuilder.registerTypeAdapter(DateTime.class, new DateTimeSerializer());
        gsonBuilder.registerTypeAdapter(DateTime.class, new DateTimeDeSerializer());
        return gsonBuilder;
    }
}
