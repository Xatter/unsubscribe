# About

This project uses the Gmail API to automatically unsubscribe from any emails that have the `List-Unsusbscribe` header.

This is the same thing Gmail uses when they prompt you if you'd like to unsubscribe to a single email, or use the unsubscribe button. Why they don't offer to automatically unsubscribe if they can detect this is probably because they dropped the motto "Don't be evil".

Gmail's API REQUIRES OAUTH, which means you need to run this from your laptop so it can open a browser window to authenticate. You might be able to run it from a headless server once you've authenticated but I haven't tried that yet.

# Requirements

* [dotnet core 8.0+](https://dotnet.microsoft.com/en-us/download)
* [Gmail Oauth](https://developers.google.com/gmail/api/auth/web-server) API setup on gCloud

# How do I run this?

    dotnet run

