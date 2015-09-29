package com.microsoft.office365.snippetapp.Snippets;

import retrofit.Callback;
import retrofit.http.GET;
import retrofit.http.Header;
import retrofit.http.Query;

public interface UsersAndGroupsService {

    @GET("/users?api-version=2013-04-05")
    public void getUsers(
            @Query("orderBy") String orderBy,
            @Query("select") String select,
            @Query("top") Integer top,
            @Query("skip") Integer skip,
            @Query("search") String search,
            @Header("Content-type") String contentTypeHeader,
            Callback<Envelope<UserValue>> callback);
}
