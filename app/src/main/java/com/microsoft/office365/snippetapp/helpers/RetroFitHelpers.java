package com.microsoft.office365.snippetapp.helpers;

import retrofit.RequestInterceptor;
import retrofit.RestAdapter;
import retrofit.converter.GsonConverter;

/**
 * Created by johnau on 8/31/2015.
 */
public  class RetroFitHelpers {

    private GsonConverter mGsonConverter;
    private  String mAccessToken;
    private RequestInterceptor mRequestInterceptor;

    public  RetroFitHelpers (String AccessToken) {
        mAccessToken = AccessToken;
    }
    public  RestAdapter getRestAdapter() {
        mGsonConverter = new GsonConverter(GsonDateTime.getDirectoryServiceBuilder()
                .create());


        RequestInterceptor requestInterceptor =  new RequestInterceptor() {
            @Override
            public void intercept(RequestFacade request) {
                final String token = mAccessToken;
                if (null != token) {
                    request.addHeader("Authorization", "Bearer " + token);
                }
            }
        };

        return new RestAdapter.Builder()
                .setEndpoint(Constants.DIRECTORY_RESOURCE_URL)
                .setLogLevel(RestAdapter.LogLevel.FULL)
                .setConverter(mGsonConverter)
                .setRequestInterceptor(requestInterceptor)
                .build();
    }

}
