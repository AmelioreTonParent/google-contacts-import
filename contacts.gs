/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Import...', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * defines a custom server-side include() function in this back-end script file 
 * to import the xxx-stylesheet.html and xxx-JavaScript.html file content into
 * the sidebar.html file. When called using printing scriptlets, this function
 * imports the specified file content into the current file. 
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  //var ui = HtmlService.createHtmlOutputFromFile('sidebar')
  //    .setTitle('Contacts import');
  var extendedUi = HtmlService.createTemplateFromFile('sidebar')
      .evaluate().setTitle('Contacts import');

  DocumentApp.getUi().showSidebar(extendedUi);
}

/**
 * Gets a list of people in the user's contacts.
 * https://developers.google.com/people/api/rest/v1/people.connections/list
 */
function getConnections() {
  // should be extracted to simulate a cache
  var contacts = People.People.Connections.list('people/me', {
    personFields: 'names,emailAddresses,addresses',
    sortOrder: 'LAST_NAME_ASCENDING'
  });
  let totalItems = contacts.totalItems;

  let nextSyncToken = contacts.nextSyncToken;
  var contactInformationArray = contacts.connections;
  return contactInformationArray;
}

/**
 * Search for dedicated templates elements and replace them by provided values.
 */
function searchAndReplace(familyName, givenName, displayName, email, address) {
  var body = DocumentApp.getActiveDocument().getBody();
  
  if (familyName) {
    body.replaceText('{contact.familyName}', familyName);
  }
  if (givenName) {
    body.replaceText('{contact.givenName}', givenName);
  }
  if (displayName) {
    body.replaceText('{contact.displayName}', displayName);
  }
  if (email) {
    body.replaceText('{contact.email}', email);
  }
  if (address) {
    body.replaceText('{contact.address}', address);
  }
}


