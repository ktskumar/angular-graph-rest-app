# Microsoft Graph + Angular JS Application (Completed App)#

This Repo is a completed app of my another Repo - [angular-graph-rest-preapp](https://github.com/ktskumar/angular-graph-rest-preapp). This Repo is generated on top of Angular QuickStart from Microsoft Graph.

Follow the below steps and use the code snippets to fetch the details from Microsoft Cloud using Microsoft Graph REST API

## How to prepare the Repo

### Prepare the App
- Download or clone the Repo  [angular-graph-rest-preapp](https://github.com/ktskumar/angular-graph-rest-preapp)
- Extract the downloaded file
- Navigate to the project folder, where you have extracted or forked
- Open the project in your favourite editor. I am using **Visual Studio Code**
- Meanwhile, Create the Application Client ID by registring in [App Registration Portal](https://apps.dev.microsoft.com/)
- During the registration of App. Ensure,
  - Platform as **Web**
  - Enable **Allow Implicit Flow** checkbox
  - Enter the Redirect URLs as http://localhost:8080/
- Copy the **Client ID** which is generated in App Registration.
![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/configClientID.png)
- In the editor navigate to *Scripts/config.js*
  - Replace **&#60;Client ID&#62;** with your **Client ID** value
  - Update the GraphScopes in config.json
  ```javascript
  graphScopes: ["user.read mail.send Sites.Read.All Files.Read Files.ReadWrite user.read.all Group.Read.All Calendars.Read Calendars.ReadWrite"]
  ```
- Run the below commands in terminal
  -  **npm install**
  -  **npm start**
- After successfull build, nvaigate to the location http://localhost:8080
- Click **Connect** button to authenticate the App
![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/connect.png)
- After connected, the view looks like below. Now you free to add your own code to connect Graph API
![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/outlook.png)

### Use the Code Snippets
The file **Public/Views/Main.html** contains required html snippets for running the complete application. So you have to concentrate on calling the GRAPH API and pass that method to the view.

If you want to add more properties and functionalities, you can modify the **Main.html** file.

#### Connect SharePoint

![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/sharepoint.png)

The below Code snippets calls the Microsoft Graph API to fetch all the sites based on the searchquery and lists from the site based on the siteid

Navigate to the file **public/scripts/graphhelper.js** and enter the below snippet under the **`//SharePoint Methods`**

```javascript
        getAllSites: function getAllSites(searchQuery){
            return $http.get('https://graph.microsoft.com/v1.0/sites?search='+searchQuery);
        },
        getSite: function getLists(siteid){
            return $http.get('https://graph.microsoft.com/v1.0/sites/'+siteid+'/lists');
        },
```

Navigate to the file **public/controllers/maincontroller.js** and enter the below snippets

Under **`//SharePoint Methods`**

```javascript
    vm.getAllSites = function(){
      var searchQuery = "*";
      GraphHelper.getAllSites(searchQuery).then(function (response){
        vm.spsites = response.data.value;
        console.log(vm.spsites);
      });
    }

    vm.getSiteInfo = function(siteid){
      GraphHelper.getLists(siteid).then(function(response){
        console.log(response);
        var strLists ="";
        response.data.value.forEach(function(lst){
          strLists +=lst.displayName+"\n";
        });
        alert(strLists);
      });
    }

```

#### Connect OneDrive

![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/onedrive.png)

The below code snippets calls the Microsoft Graph API to fetch the folder, files and create a folder

Navigate to the file **public/scripts/graphhelper.js** and enter the below snippet under the **`//OneDrive Methods`**

```javascript
        getDriveItems: function(){
            return $http.get('https://graph.microsoft.com/v1.0/me/drive/root/children');;
        },
        getFolderItems: function(folderId){
            return $http.get('https://graph.microsoft.com/v1.0/me/drive/items/'+folderId+'/children');            
        },
        createFolder: function(folderInfo){
            return $http.post('https://graph.microsoft.com/v1.0/me/drive/root/children',folderInfo);
        },
```

Navigate to the file **public/controllers/maincontroller.js** and enter the below snippets

Under **`//OneDrive Methods`**

```javascript
    vm.getDriveItems = function(){
      GraphHelper.getDriveItems().then(function(response){
        vm.oditems = response.data.value;
      })

    }

    vm.getFolderItems = function(folderid){
      GraphHelper.getFolderItems(folderid).then(function(response){
        console.log(response);
        var strItems = "";
        response.data.value.forEach(function(item){
          strItems += item.name+'\n';
        });
        alert(strItems);
      });
    }

    vm.createFolder = function(){
      var folderInfo ={
        "name": vm.odFolderName,
        "folder": {}
      };
      GraphHelper.createFolder(folderInfo).then(function (response) {
        console.log(response);
        $log.debug('HTTP request to the Microsoft Graph API returned successfully.', response);
        response.status === 201 ? vm.folderSuccess = true : vm.folderSuccess = false;
        vm.folderFinished= true;
      }, function (error) {
        $log.error('HTTP request to the Microsoft Graph API failed.');
        vm.folderSuccess = false;
        vm.folderFinished = true;
      });
    }
```


#### Connect Users and Groups

![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/usersgroups.png)

The below code snippets calls the Microsoft Graph API to fetch the users and groups

Navigate to the file **public/scripts/graphhelper.js** and enter the below snippet under the **`//Users and Groups Methods`**

```javascript
        getAllUsers: function getAllUsers(){
            return $http.get('https://graph.microsoft.com/v1.0/users');
        },
        getAllGroups: function getAllGroups(){
            return $http.get('https://graph.microsoft.com/v1.0/groups');
        },
        getUserGroups: function getUserGroups(usrid){
            return $http.get('https://graph.microsoft.com/v1.0/users/'+usrid+'/memberOf');
        },
        getGroupUsers: function getGroupUsers(grpid){
            return $http.get('https://graph.microsoft.com/v1.0/groups/'+grpid+'/members');
        },

```

Navigate to the file **public/controllers/maincontroller.js** and enter the below snippets

Under **`//Users and Groups Methods`**

```javascript
    vm.getAllUsers = function(){
      GraphHelper.getAllUsers().then(function(response){
        vm.allusers = response.data.value;
      });
    }

    vm.getAllGroups = function(){
      GraphHelper.getAllGroups().then(function(response){
        vm.allgroups = response.data.value;
      });
    }

    vm.getUserGroups = function (usrid){
      GraphHelper.getUserGroups(usrid).then(function (response) {
        var str = "Total Groups: " + response.data.value.length + "\r\n";
        response.data.value.forEach(function (item) {
          str += item.displayName + "\r\n";
        });
        alert(str);
        console.log(response.data.value);
      });
    }

    vm.getGroupUsers = function(grpid){
      GraphHelper.getGroupUsers(grpid).then(function (response) {
        var str = "Total Users: " + response.data.value.length + "\r\n";
        response.data.value.forEach(function (item) {
          str += item.displayName + "\r\n";
        });
        alert(str);
        console.log(response.data.value);
      });
    }

```

#### Connect Calendar Events

![](https://raw.githubusercontent.com/ktskumar/angular-graph-rest-app/master/README_assets/events.png)

The below code snippets calls the Microsoft Graph API to fetch the events for the user and create a new event

Navigate to the file **public/scripts/graphhelper.js** and enter the below snippet under the **`//Calendar Event methods`**

```javascript
        getEvents: function getEvents(){
            return $http.get('https://graph.microsoft.com/v1.0/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location');            
        },
        createEvent: function createEvent(meetingInfo){
            return $http.post('https://graph.microsoft.com/v1.0/me/events',meetingInfo);
        } 

```

Navigate to the file **public/controllers/maincontroller.js** and enter the below snippets

Under **`//Calendar Event methods`**

```javascript
    vm.getEvents = function () {
      GraphHelper.getEvents().then(function (response) {
        console.log(response);
        vm.allevents = response.data.value;
      });
    }

    vm.createEvent = function(){
      var meetingInfo ={
        "subject": vm.meetSubject,
        "start": {
          "dateTime": vm.meetStart,
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": vm.meetEnd,
          "timeZone": "UTC"
        }
      };

      GraphHelper.createEvent(meetingInfo).then(function (response) {
        $log.debug('HTTP request to the Microsoft Graph API returned successfully.', response);
        response.status === 201 ? vm.eventSuccess = true : vm.eventSuccess = false;
        vm.eventFinished = true;
      }, function (error) {
        $log.error('HTTP request to the Microsoft Graph API failed.');
        vm.eventSuccess = false;
        vm.eventFinished = true;
      });
    }
```

After updating all the above codes. Run **npm start** from the terminal to get the final output. The repo for final output is available here [angular-graph-rest-app](https://github.com/ktskumar/angular-graph-rest-app).


*Cheers!*
- **Shantha Kumar T** ( **@ktskumar** )
- [ktskumar.com](http://www.ktskumar.com)