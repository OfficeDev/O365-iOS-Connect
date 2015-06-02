/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "Office365ClientFetcher.h"
#import "AuthenticationManager.h"

@implementation Office365ClientFetcher

// Gets the Outlook Services client. This will authenticate a user with the service
// and get the application an access and refresh token to act on behalf of the user.
// The access and refresh token will be cached. The next time a user attempts
// to access the service, the access token will be used. If the access token
// has expired, the client will use the refresh token to get a new access token.
// If the refresh token has expired, then ADAL will get the authorization code
// from the cookie cache and use that to get a new access and refresh token.
- (void)fetchOutlookClient:(void (^)(MSOutlookClient *outlookClient))callback
{
    // Get an instance of the authentication controller.
    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
    
    // The first time this application is run, the authentication manager will send a request
    // to the authority which will redirect you to a login page. You will provide your credentials
    // and the response will contain your refresh and access tokens. The second time this
    // application is run, and assuming you didn't clear your token cache, the authentication
    // manager will use the access or refresh token in the cache to authenticate client requests.
    // This will result in a call to the service if you need to get an access token.
    [authenticationManager acquireAuthTokenWithResourceId:@"https://outlook.office365.com/"
                                           completionHandler:^(BOOL authenticated) {

        if(authenticated){
            NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];

            NSDictionary *serviceEndpoints = [userDefaults objectForKey:@"O365ServiceEndpoints"];

            // Gets the MSOutlookClient with the URL for the Mail service.
            callback([[MSOutlookClient alloc] initWithUrl:serviceEndpoints[@"Mail"]
                                       dependencyResolver:authenticationManager.dependencyResolver]);
            
        }
        else{
            //Display an alert in case of an error
            dispatch_async(dispatch_get_main_queue(), ^{
                NSLog(@"Error in the authentication");
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                message:@"Authentication failed. Check the log for errors."
                                                               delegate:self
                                                      cancelButtonTitle:@"OK"
                                                      otherButtonTitles:nil];
                [alert show];
                
            });
        }
    }];
}


//Gets the DiscoveryClient which is used to discover the service endpoints
- (void)fetchDiscoveryClient:(void (^)(MSDiscoveryClient *discoveryClient))callback
{

    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
    
    // The first time this application is run, the authentication manager will send a request
    // to the authority which will redirect you to a login page. You will provide your credentials
    // and the response will contain your refresh and access tokens. The second time this
    // application is run, and assuming you didn't clear your token cache, the authentication
    // manager will use the access or refresh token in the cache to authenticate client requests.
    // This will result in a call to the service if you need to get an access token.
    [authenticationManager acquireAuthTokenWithResourceId:@"https://api.office.com/discovery/"
                                        completionHandler:^(BOOL authenticated) {
        if (authenticated) {
            callback([[MSDiscoveryClient alloc] initWithUrl:@"https://api.office.com/discovery/v1.0/me/"
                                         dependencyResolver:authenticationManager.dependencyResolver]);

        }
        else {
            dispatch_async(dispatch_get_main_queue(), ^{
                NSLog(@"Error in the authentication");
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                message:@"Authentication failed. This may be because the Internet connection is offline  or perhaps the credentials are incorrect. Check the log for errors and try again."
                                                               delegate:self
                                                      cancelButtonTitle:@"OK"
                                                      otherButtonTitles:nil];
                [alert show];

            });
        }
    }];
}

@end
// *********************************************************
//
// O365-iOS-Connect, https://github.com/OfficeDev/O365-iOS-Connect
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