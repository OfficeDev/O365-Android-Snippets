/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

package com.microsoft.office365.snippetapp.helpers;

import android.os.Build;
import android.util.Log;

import com.microsoft.aad.adal.AuthenticationSettings;
import com.microsoft.office365.snippetapp.OperationListActivity;

import java.io.UnsupportedEncodingException;
import java.security.NoSuchAlgorithmException;
import java.security.spec.InvalidKeySpecException;

import javax.crypto.SecretKey;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;
import javax.crypto.spec.SecretKeySpec;

public class AuthUtil {

    public static final int MIN_SDK_VERSION_FOR_ENCRYPT = 18;
    private static final String TAG = "AuthUtil";

    public static void setupEncryptionKey(OperationListActivity activity) {
        // Devices with API level lower than 18 must setSecretKey an encryption key.
        if (Build.VERSION.SDK_INT >= MIN_SDK_VERSION_FOR_ENCRYPT) {
            return;
        }
        try {
            setSecretKey();
        } catch (NoSuchAlgorithmException | InvalidKeySpecException | UnsupportedEncodingException ex) {
            Log.e(TAG, "setup encryption key failed!");
            activity.showEncryptionKeyErrorUI();
        }
        AuthenticationSettings.INSTANCE.setSkipBroker(true);
    }

    private static void setSecretKey() throws NoSuchAlgorithmException,
            InvalidKeySpecException, UnsupportedEncodingException {
        if (AuthenticationSettings.INSTANCE.getSecretKeyData() == null) {
            SecretKeyFactory keyFactory = SecretKeyFactory.getInstance(
                    "PBEWithSHA256And256BitAES-CBC-BC");
            PBEKeySpec keySpec = new PBEKeySpec("O365_password".toCharArray(),
                    "O365_salt".getBytes("UTF-8"),
                    100,
                    256);
            SecretKey tempKey = keyFactory.generateSecret(keySpec);
            SecretKey secretKey = new SecretKeySpec(tempKey.getEncoded(), "AES");
            AuthenticationSettings.INSTANCE.setSecretKey(secretKey.getEncoded());
        }
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
