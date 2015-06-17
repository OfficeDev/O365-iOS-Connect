/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "SendMailViewController.h"
#import "Office365ClientFetcher.h"
#import "AuthenticationManager.h"
#import "MSDiscoveryServiceInfoCollectionFetcher.h"

@interface SendMailViewController ()

@property (weak, nonatomic) IBOutlet UILabel *headerLabel;
@property (weak, nonatomic) IBOutlet UITextView *mainContentTextView;
@property (weak, nonatomic) IBOutlet UITextView *statusTextView;
@property (weak, nonatomic) IBOutlet UITextField *emailTextField;
@property (weak, nonatomic) IBOutlet UIBarButtonItem *disconnectBarButtonItem;
@property (weak, nonatomic) IBOutlet UIButton *sendMailButton;
@property (weak, nonatomic) IBOutlet UIActivityIndicatorView *activityIndicator;

@property (strong, nonatomic) Office365ClientFetcher *baseController;
@property (strong, nonatomic) NSMutableDictionary *serviceEndpointLookup;

- (IBAction)sendMailTapped:(id)sender;
- (IBAction)disconnectTapped:(id)sender;

@end


@implementation SendMailViewController

#pragma mark - Properties
- (Office365ClientFetcher *)baseController
{
    if (!_baseController) {
        _baseController = [[Office365ClientFetcher alloc] init];
    }

    return _baseController;
}

#pragma mark - Lifecycle Methods
- (void)viewDidLoad
{
    [super viewDidLoad];

    self.disconnectBarButtonItem.enabled = NO;
    self.sendMailButton.hidden = YES;
    self.emailTextField.hidden = YES;
    self.mainContentTextView.hidden = YES;
    self.headerLabel.hidden = YES;

    [self connectToOffice365];
}

#pragma mark - IBActions
//Send a mail message when the Send button is clicked
- (IBAction)sendMailTapped:(id)sender
{
    [self sendMailMessage];
}

// Clear the token cache and update the UI when the Disconnect button is tapped
- (IBAction)disconnectTapped:(id)sender
{
    self.disconnectBarButtonItem.enabled = NO;
    self.sendMailButton.hidden = YES;
    self.mainContentTextView.text = @"You're no longer connected to Office 365.";
    self.headerLabel.hidden = YES;
    self.emailTextField.hidden = YES;
    self.statusTextView.hidden = YES;

    // Clear the access and refresh tokens from the credential cache. You need to clear cookies
    // since ADAL uses information stored in the cookies to get a new access token.
    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
    [authenticationManager clearCredentials];
}

#pragma mark - Helper Methods
- (void)connectToOffice365
{
    // Connect to the service by discovering the service endpoints and authorizing
    // the application to access the user's email. This will store the user's
    // service URLs in a property list to be accessed when calls are made to the
    // service. This results in two calls: one to authenticate, and one to get the
    // URLs. ADAL will cache the access and refresh tokens so you won't need to
    // provide credentials unless you sign out.

    // Get the discovery client. First time this is ran you will be prompted
    // to provide your credentials which will authenticate you with the service.
    // The application will get an access token in the response.
    [self.baseController fetchDiscoveryClient:^(MSDiscoveryClient *discoveryClient) {
        MSDiscoveryServiceInfoCollectionFetcher *servicesInfoFetcher = [discoveryClient getservices];

        // Call the Discovery Service and get back an array of service endpoint information
        NSURLSessionTask *servicesTask = [servicesInfoFetcher readWithCallback:^(NSArray *serviceEndpoints, MSODataException *error) {
            if (serviceEndpoints) {
                
                // Here is where we cache the service URLs returned by the Discovery Service. You may not
                // need to call the Discovery Service again until either this cache is removed, or you
                // get an error that indicates that the endpoint is no longer valid.
                self.serviceEndpointLookup = [[NSMutableDictionary alloc] init];

                for(MSDiscoveryServiceInfo *serviceEndpoint in serviceEndpoints) {
                    self.serviceEndpointLookup[serviceEndpoint.capability] = serviceEndpoint.serviceEndpointUri;
                }

                // Keep track of the service endpoints in the user defaults
                NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];

                [userDefaults setObject:self.serviceEndpointLookup
                                 forKey:@"O365ServiceEndpoints"];

                [userDefaults synchronize];

                dispatch_async(dispatch_get_main_queue(), ^{
                   
                    NSString *userEmail = [userDefaults stringForKey:@"LogInUser"];
                    NSArray *parts = [userEmail componentsSeparatedByString: @"@"];

                    self.headerLabel.text = [NSString stringWithFormat:@"Hi %@!", parts[0]];
                    self.headerLabel.hidden = NO;
                    self.mainContentTextView.hidden = NO;
                    self.emailTextField.text = userEmail;
                    self.statusTextView.text = @"";
                    self.disconnectBarButtonItem.enabled = YES;
                    self.sendMailButton.hidden = NO;
                    self.emailTextField.hidden = NO;
                });
            }
            else {
                dispatch_async(dispatch_get_main_queue(), ^{
                    NSLog(@"Error in the authentication: %@", error);

                    UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                    message:@"Authentication failed. This may be because the Internet connection is offline  or perhaps the credentials are incorrect. Check the log for errors and try again."
                                                                   delegate:nil
                                                          cancelButtonTitle:@"OK"
                                                          otherButtonTitles:nil];
                    [alert show];
                });
            }
        }];
        
        [servicesTask resume];
    }];
}

// This method creates a new mail message and sends it to the specified address
// by using Office 365.
- (void)sendMailMessage
{
    MSOutlookMessage *message = [self buildMessage];

    // Get the MSOutlookClient. A call will be made to Azure AD and you will be prompted for credentials if you don't
    // have an access or refresh token in your token cache.
    [self.baseController fetchOutlookClient:^(MSOutlookClient *outlookClient) {
        dispatch_async(dispatch_get_main_queue(), ^{
            // Show the activity indicator
            [self.activityIndicator startAnimating];
        });

        MSOutlookUserFetcher *userFetcher = [outlookClient getMe];
        MSOutlookUserOperations *userOperations = [userFetcher operations];

        // Send the mail message. This results in a call to the service.
        NSURLSessionTask *task = [userOperations sendMailWithMessage:message
                                                     saveToSentItems:YES
                                                            callback:^(int returnValue, MSODataException *error) {
            NSString *statusText;

            if (error == nil) {
                statusText = @"Check your inbox, you have a new message. :)";
            }
            else {
                statusText = @"The email could not be sent. Check the log for errors.";
                NSLog(@"%@",[error localizedDescription]);
            }

            // Update the UI.
            dispatch_async(dispatch_get_main_queue(), ^{
                self.statusTextView.text = statusText;
                [self.activityIndicator stopAnimating];
            });
        }];

        [task resume];
    }];
}

//Compose the mail message
- (MSOutlookMessage *)buildMessage
{
    // Create a new message. Set properties on the message.
    MSOutlookMessage *message = [[MSOutlookMessage alloc] init];
    message.Subject = @"Welcome to Office 365 development on iOS with the Office 365 Connect sample";

    // Get the recipient's email address.
    // The ToRecipients property is an array of MSOulookRecipient objects.
    // See the helper method getRecipients to understand the usage.
    NSString *toEmail = self.emailTextField.text;

    MSOutlookRecipient *recipient = [[MSOutlookRecipient alloc] init];

    recipient.EmailAddress = [[MSOutlookEmailAddress alloc] init];
    recipient.EmailAddress.Address = [toEmail stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]];

    // The cast here is required to maintain compatibility with the API.
    message.ToRecipients = (NSMutableArray<MSOutlookRecipient> *)[[NSMutableArray alloc] initWithObjects:recipient, nil];

    // Get the email text and put in the email body.
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"EmailBody" ofType:@"html" ];
    NSString *body = [[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]
                      stringByReplacingOccurrencesOfString:@"\"" withString:@"\\\""];
    message.Body = [[MSOutlookItemBody alloc] init];
    message.Body.ContentType = MSOutlook_BodyType_HTML;
    message.Body.Content = body;

    return message;
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
