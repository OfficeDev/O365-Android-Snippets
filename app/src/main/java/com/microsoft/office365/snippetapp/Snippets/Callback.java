package com.microsoft.office365.snippetapp.Snippets;

import java.util.Map;

public interface Callback<T> extends retrofit.Callback<T> {

    Map<String, String> getParams();
}
