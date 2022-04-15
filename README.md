# GraphTaskFromCalendar
 Creates Tasks in Planner from Calendar Events from a User

## ➤ Longer Description
 This was created with C# and the Graph API to pull from a shared calendar each month and create tasks in Planner so we can assign them to people manually afterward. In our use case we are pulling the subject and preview of the body as the task title and description, as well as the end date of said calendar event as the due date. 

## ➤ Resources
 Things I used: 
 Visual Studio Code, C#, .NET 6.0, Graph API, Azure App Registrations, Postman, Restsharp.
 
 I created this using a few basic resources, starting with [this guide](https://github.com/microsoftgraph/dotnetcore-console-sample) which is the basic foundation of the app. Secondly I used [Postman](https://www.postman.com) and [this guide](https://dzone.com/articles/getting-access-token-for-microsoft-graph-using-oau) to create an access token which I will get into further down.

## ➤ Quick Setup

Since most of this is basically pulled from the two other guides, I will refer to those in how to [set up](https://github.com/microsoftgraph/dotnetcore-console-sample/tree/main/base-console-app) the app registration in Azure (I will also show what access I have granted below this paragraph), and how to do your own pull request for the [access token](https://dzone.com/articles/getting-access-token-for-microsoft-graph-using-oau) if you don't want to prompt for login every time, hint: use flow 2. What you will need to change before you think, "It doesn't work!" I will explain below.

![image](https://user-images.githubusercontent.com/1349908/161835376-e5a1991d-765f-490f-a17d-1e3f1e3b95cd.png)


Starting with the most important part, appsettings.json, we're going to need your app registration and your company/user information put in here. Replace the variables with your own information, this will be called by the other programs. The first three will be long strings, the final two are simply your domain. Do not lose the secret, once you create the App, the first time you get it, will be the only time you see it, save it somewhere safe or else you'll have to create a new App, and no one likes doing that a second time with all the permissions needed.

    {
    
    "applicationId": "appid",
    
    "applicationSecret": "yoursecret",
    
    "tenantId": "yourdomainid",
    
    "redirectUri": "https://uri.yourdomain.com",
    
    "domain": "yourdomain.com"
    
    }

Now that is set up, lets move into the Helpers folder and work with PlannerHelper.cs. This is where your inner 12yo hacker self can come out as we'll be doing some inspect element fun. Lines 20, 25, and 27 have some variables for you to change. You can find the first two in your planner url at https://tasks.office.com/

Get yourself into the group and plan you want to be in, then copy out your url bar such as this (obviously not my real one) where you'll find your groupId and planId waiting for you to copy over. Why didn't I put these into the appsettings.json? I don't know, maybe I will in the future (probably not) I didn't think this far ahead when I started.

``https://tasks.office.com/yourdomain.com/en-US/Home/Planner/#/plantaskboard?groupId=3439sdad-ce16-4eaa-73b1-9f6c236854155&planId=yi7xKSBr054DseEFM89wi2UAHg9l``

Next we need the bucketId and this is where you get to have a little old school fun. Right Click and Inspect Element on your bucket that you want the tasks to be dropped into. From here you should end up at the h3 tag, go up a couple divs until you reach the li tag, then find your id and copy that out, that's your bucketId. 

Congratulations, we're done! Unless you want to automate it and not prompt for a user to log in all the time. Then, we'll need to do a couple different things. Open back up that appsettings.json file we got some extra work to do.

Running Flow 2 from [the second guide](https://dzone.com/articles/getting-access-token-for-microsoft-graph-using-oau) using Postman, or whatever you prefer, I pulled out the C# Rest code and suffered as I found out it hasn't been updated with how RestSharp works anymore and solved all the problems for you. This is the real automation. All the things you need to update are in appsettings.json. The first two you don't need to touch, get your username and password in there, and for the final microsoftLogin, change "yourdomainid" to the same tenantId from earlier in the file.

    {
        
    "grant_type": "password",
    
    "resource": "https://graph.microsoft.com",
    
    "username": "username@yourdomain.com",
    
    "password": "actually the password",
    
    "microsoftLogin": "https://login.microsoftonline.com/yourdomainid/oauth2/token"

    }

"But wait, you're storing the password in plain text!" I can hear you screaming, well this is on a secure server for an account specifically for this purpose with very limited access to who can even see it, so it doesn't matter in my case, if you need to encrypt it, go for it.

Now most of the changes if you'd like to change the way it works will be made in PlannerHelper.cs. This script by default checks the next month of the logged in user's calendar and pulls specific info, you can change the date range on lines 29 and 30, and the info/if you want to change what it searches, below that.

## ➤ License
 Licensed under [MIT](https://opensource.org/licenses/MIT).
