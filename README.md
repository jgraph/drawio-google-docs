# drawio-google-docs
draw.io add-ons for Google Docs/Sheets/Slides

You can install these add-ons using your own account by following the instructions below:

1. Open Google Docs/Sheets/Slides, then from "Extensions" menu, select "Apps Script".
2. In the open editor, copy the code from Code.gs and paste it into the editor. Then, add a new HTML file named "Picker.html" and copy the code from Picker.html and paste it into the editor.
3. Now, save the project and give it a name (e.g, My draw.io).
4. Click the Deploy button and select "Test deployments". Select "Editor Add-on" from the deployment type gear icon.
5. Click "Add test", and select a document to test on. Then, click "Save Test" button.
6. Finally, select the newly created test and click "Execute" button. The document will open and you can access the add-on from the "Extensions" menu -> "You project name (e.g, My draw.io)".

## Publish the add-on privately

Follow the steps in https://developers.google.com/apps-script/add-ons/how-tos/publish-add-on-overview and https://developers.google.com/workspace/marketplace/how-to-publish to publish the add-on privately. When you publish the add-on privately, you can share it with your organization or a specific group of users. In addition, it won't require a review from Google.