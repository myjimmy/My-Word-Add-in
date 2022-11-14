# My-Word-Add-in
[Build your first Word task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator) using Office Add-in.

---
## Prerequisites
For more information, please refer to the following link: <br/>
https://learn.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment

### Node.js
Go to the [Node.js website](https://nodejs.org/en/) and download the latest recommended version.

**npm** usually installed automatically when you install Node.js.

Please run the following commands to check the version of **Node.js** and **npm**:
```powershell
$ node -v
v18.12.0

$ npm -v
8.19.2
```

### Install a code editor
You can use any code editor or IDE that supports client-side development to build your web part, such as:

- [Visual Studio Code](https://code.visualstudio.com/) (recommended)
- [Atom](https://atom.io/)
- [Webstorm](https://www.jetbrains.com/webstorm)

### Install the Yeoman generator — Yo Office
The project creation and scaffolding tool is [Yeoman generator for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview), affectionately known as Yo Office. You need to install the latest version of [Yeoman](https://github.com/yeoman/yo) and Yo Office. To install these tools globally, run the following command via the command prompt.
```powershell
$ npm install -g yo generator-office
```

### Install and use the Office JavaScript linter
Microsoft provides a JavaScript linter to help you catch common errors when using the Office JavaScript library. To install the linter, run the following two commands:
```powershell
$ npm install office-addin-lint --save-dev
$ npm install eslint-plugin-office-addins --save-dev
```

Run the linter with the following command either in the terminal of an editor, such as Visual Studio Code, or in a command prompt.
```powershell
$ npm run lint
```

## Create the add-in project
For more information, please refer to the following link: <br/>
https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator#create-the-add-in

Run the following command to create an add-in project using the **Yeoman generator**.
```powershell
$ yo office
```

When prompted, provide the following information to create your add-in project.

* **Choose a project type:** `Office Add-in Task Pane project`
* **Choose a script type:** `Javascript`
* **What do you want to name your add-in?** `My Word Add-in`
* **Which Office client application would you like to support?** `Word`

![Alt text](help/yo-office-word.png?raw=true)

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Explore the project
The add-in project that you've created with the Yeoman generator contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.
* The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.<br/>To learn more about the manifest.xml file, see [Office Add-ins XML manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).
* The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
* The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
* The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.

## Try it out
1. Navigate to the root folder of the project.
   ```powershell
   $ cd "My-Word-Add-in"
   ```
2. Complete the following steps to start the local web server and sideload your add-in.
   ```
   Note
   Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
   ```
  * To test your add-in in Word, run the following command in the root directory of your project.<br/>This starts the local web server and opens ExWordcel with your add-in loaded.
    ```powershell
    $ npm start
    ```
  * To test your add-in in Word on a browser, run the following command in the root directory of your project.<br/>When you run this command, the local web server starts.<br/>Replace "{url}" with the URL of an Word document on your OneDrive or a SharePoint library to which you have permissions.
    ```powershell
    $ npm run start:web -- --document {url}
    ```
    The following are examples.
    * npm run start:web -- --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3VYhYQ?e=F4QM1R
    * npm run start:web -- --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
    * npm run start:web -- --document https://contoso-my.sharepoint-df.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP

    If your add-in doesn't sideload in the document, manually sideload it by following the instructions in [Manually sideload add-ins to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web).

3. In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.<br/>
   ![Alt text](help/word-quickstart-addin-2b.png?raw=true)

4. At the bottom of the task pane, choose the **Run** link to add the text "Hello World" to the document in blue font.
   ![Alt text](help/word-quickstart-addin-1c.png?raw=true)

## Develop new features
We develop the folowing new features using "[Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial)".
1. Insert a paragraph.<br/>
   Please follow the [Insert a range of text](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial#insert-a-range-of-text) to develop this feature. (See the below image)
   ![Alt text](help/word-insert-paragraph.png?raw=true)
