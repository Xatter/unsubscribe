# About

This project uses the Gmail API to automatically unsubscribe from any emails that have the `List-Unsusbscribe` header.

This is the same thing Gmail uses when they prompt you if you'd like to unsubscribe to a single email, or use the unsubscribe button. Why they don't offer to automatically unsubscribe if they can detect this is probably because they dropped the motto "Don't be evil".

Gmail's API REQUIRES OAUTH, which means you need to run this from your laptop so it can open a browser window to authenticate. You might be able to run it from a headless server once you've authenticated but I haven't tried that yet.

# Requirements

* [dotnet core 8.0+](https://dotnet.microsoft.com/en-us/download)
* [Gmail Oauth](https://developers.google.com/gmail/api/auth/web-server) API setup on gCloud

# How do I run this?

    dotnet run

If this code looks like "That's a pretty strange way to do this" I agree! This code is optimized around Gmail API call limits and not necessarily the best possible parallel execution.

Therefore, we group based on the sender of an email and only send 1 unsubscribe per sender, not per-email.

# What does it do?

As written, this will find all of the messages that have the `List-Unsubscribe` header and then it will parse that header, following any links or sending unsubscribe emails to that sender.

It will do this only once per-sender, but it will then delete all the messages from that sender.

You can optionally include a folder/label and it will do the above for only that folder.


# WHY IS THERE CODE???

Because I'm a smelly nerd