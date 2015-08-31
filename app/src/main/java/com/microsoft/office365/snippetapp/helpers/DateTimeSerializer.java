package com.microsoft.office365.snippetapp.helpers;

import com.google.gson.JsonElement;
import com.google.gson.JsonPrimitive;
import com.google.gson.JsonSerializationContext;
import com.google.gson.JsonSerializer;

import org.joda.time.DateTime;

import java.lang.reflect.Type;

/**
 * Created by johnau on 8/31/2015.
 */
public class DateTimeSerializer implements JsonSerializer<DateTime> {
    public JsonElement serialize(DateTime src, Type typeOfSrc, JsonSerializationContext context) {
        return new JsonPrimitive(src.toString());
    }

}
