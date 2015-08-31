package com.microsoft.office365.snippetapp.helpers;

import com.google.gson.JsonDeserializationContext;
import com.google.gson.JsonDeserializer;
import com.google.gson.JsonElement;
import com.google.gson.JsonParseException;

import org.joda.time.DateTime;

import java.lang.reflect.Type;

/**
 * Created by johnau on 8/31/2015.
 */
public class DateTimeDeSerializer implements JsonDeserializer<DateTime> {
    public DateTime deserialize(JsonElement json, Type typeOfT, JsonDeserializationContext context)
            throws JsonParseException {
        return new DateTime(json.getAsJsonPrimitive().getAsString());
    }
}
