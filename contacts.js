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
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Import...", "showSidebar")
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
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  //var ui = HtmlService.createHtmlOutputFromFile('sidebar')
  //    .setTitle('Contacts import');
  var extendedUi = HtmlService.createTemplateFromFile("sidebar")
    .evaluate()
    .setTitle("Contacts import");

  DocumentApp.getUi().showSidebar(extendedUi);
}

/**
 * Gets a list of people in the user's contacts.
 * https://developers.google.com/people/api/rest/v1/people.connections/list
 */
function getConnections() {
  // there is a pagination mechanism
  // and we want to retrieve every contacts right now !
  const defaultPageSize = 100; // see api definition
  // should be extracted to simulate a cache
  let contacts = People.People.Connections.list("people/me", {
    personFields: "names,emailAddresses,addresses",
    pageSize: defaultPageSize,
  });

  // store initial connections array
  let contactInformationArray = contacts.connections;

  // check pagination !
  let totalItems = contacts.totalItems;
  while (contactInformationArray.length < totalItems) {
    // it means we have to retrieve more connections !
    contacts = People.People.Connections.list("people/me", {
      personFields: "names,emailAddresses,addresses",
      pageSize: defaultPageSize,
      pageToken: contacts.nextPageToken,
    });
    // add more elements in existing array
    contactInformationArray = contactInformationArray.concat(
      contacts.connections
    );
  }

  // for a later synchronization ...
  let nextSyncToken = contacts.nextSyncToken;
  return contactInformationArray;
}

/**
 * Computes document name using provided information,
 * also to be able to retrieve this document easily,
 * add a date information.
 */
function computeDocumentName(familyName, givenName, displayName) {
  let documentName;
  if (displayName) {
    documentName = displayName;
  } else {
    documentName = familyName + " " + givenName;
  }
  // add date information
  let date = new Date();
  documentName = documentName + " " + date.toISOString();
  return documentName;
}

/**
 * Creates a new document using current one as template.
 *
 * In new document, search for dedicated templates elements and replace them by provided values.
 */
function createDocumentFromTemplateWith(
  familyName,
  givenName,
  displayName,
  email,
  address
) {
  let documentName = computeDocumentName(familyName, givenName, displayName);
  let currentDocument = DocumentApp.getActiveDocument();

  const options = {
    fields: "id", // properties sent back to you from the API
    //supportsTeamDrives: true, // needed for Team Drives
  };
  const metadata = {
    title: documentName,
    // other possible fields you can supply:
    // https://developers.google.com/drive/api/v2/reference/files/copy#request-body
  };

  let newDocumentId = Drive.Files.copy(
    metadata,
    currentDocument.getId(),
    options
  );
  let newDocument = DocumentApp.openById(newDocumentId.id);

  let body = newDocument.getBody();
  if (familyName) {
    body.replaceText("{contact.familyName}", familyName);
  }
  if (givenName) {
    body.replaceText("{contact.givenName}", givenName);
  }
  if (displayName) {
    body.replaceText("{contact.displayName}", displayName);
  }
  if (email) {
    body.replaceText("{contact.email}", email);
  }
  if (address) {
    body.replaceText("{contact.address}", address);
  }

  let newDocumentUrl = newDocument.getUrl();
  // save modifications to current document
  newDocument.saveAndClose();
  return newDocumentUrl;
}

/**
 * Create a new contact in People management.
 * Need the full scope to call this api.
 *
 * @param {string} familyName
 * @param {string} givenName
 * @param {string} email
 * @param {string} address
 */
function createContact(familyName, givenName, email, address) {
  let newContact = People.People.createContact({
    names: [
      {
        familyName: familyName,
        givenName: givenName,
      },
    ],
    addresses: [
      {
        formattedValue: address,
      },
    ],
    emailAddresses: [
      {
        value: email,
      },
    ],
  });
  // should we had the contact to current list ?
}

/**
 * Import contact information into google docs.
 * Several steps are included:
 * - replace templates by corresponding values,
 * - create a contact if this is a manual edition,
 * - and later, create a new document (in order to avoid modifications of this template).
 *
 * @param {string} familyName
 * @param {string} givenName
 * @param {string} displayName
 * @param {string} email
 * @param {string} address
 * @param {boolean} manualEditFlag
 *
 * @returns url of created document.
 */
function importContactInformation(
  familyName,
  givenName,
  displayName,
  email,
  address,
  manualEditFlag
) {
  let newDocumentUrl = createDocumentFromTemplateWith(
    familyName,
    givenName,
    displayName,
    email,
    address
  );
  if (manualEditFlag) {
    createContact(familyName, givenName, email, address);
  }
  return newDocumentUrl;
}
