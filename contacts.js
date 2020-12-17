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
function populateAllConnectionsWithoutSyncToken() {
  // there is a pagination mechanism
  // and we want to retrieve every contacts right now !
  const defaultPageSize = 100; // see api definition
  // should be extracted to simulate a cache
  let contacts = People.People.Connections.list("people/me", {
    personFields: "names,emailAddresses,addresses",
    pageSize: defaultPageSize,
    requestSyncToken: true,
  });

  // store initial connections array
  contactInformationArray = contacts.connections;

  // check pagination !
  let totalItems = contacts.totalItems;
  while (contactInformationArray.length < totalItems) {
    // it means we have to retrieve more connections !
    contacts = People.People.Connections.list("people/me", {
      personFields: "names,emailAddresses,addresses",
      pageSize: defaultPageSize,
      pageToken: contacts.nextPageToken,
      requestSyncToken: true,
    });
    // add more elements in existing array
    contactInformationArray = contactInformationArray.concat(
      contacts.connections
    );
  }

  // for a later synchronization ...
  // and finally update current array of contacts
  return {
    synchronizationFlag: false,
    nextSyncToken: contacts.nextSyncToken,
    contactInformationArray: contactInformationArray,
  };
}

/**
 * We have a synchronization token for contacts, but it could be expired:
 *
 * The request returns a 410 error if syncToken is specified and is expired.
 * Sync tokens expire after 7 days to prevent data drift between clients and the server.
 * To handle a sync token expired error, a request should be sent without syncToken to get all contacts.
 */
function synchronizeConnections(nextSyncToken) {
  // store synchronization token and connections array (try to cache them)
  let contactInformation;
  try {
    console.log("start synchronization.");
    // there is a pagination mechanism
    // and we want to retrieve every contacts right now !
    const defaultPageSize = 100; // see api definition
    // should be extracted to simulate a cache
    let contacts = People.People.Connections.list("people/me", {
      personFields: "names,emailAddresses,addresses",
      pageSize: defaultPageSize,
      requestSyncToken: true,
      syncToken: nextSyncToken,
    });
    console.log("current token = " + nextSyncToken);
    console.log("next token = " + contacts.nextSyncToken);
    // store initial connections array
    let contactInformationArray = contacts.connections;

    // check pagination !
    console.log("total number of items = " + contacts.totalItems);
    if (contactInformationArray) {
      let retrievedItems = contactInformationArray.length;
      console.log("retrieved contacts = " + retrievedItems);
      for (contact of contactInformationArray) {
        console.log("retrieved contact = " + contact.toString());
      }
      while (retrievedItems == defaultPageSize) {
        // it means we have to retrieve more connections !
        contacts = People.People.Connections.list("people/me", {
          personFields: "names,emailAddresses,addresses",
          pageSize: defaultPageSize,
          pageToken: contacts.nextPageToken,
          requestSyncToken: true,
          syncToken: nextSyncToken,
        });
        if (contacts.connections) {
          retrievedItems = contacts.connections.length;
          // add more elements in existing array
          contactInformationArray = contactInformationArray.concat(
            contacts.connections
          );
        } else {
          retrievedItems = 0;
        }
        console.log("iterate !!!!");
        console.log("current token = " + nextSyncToken);
        console.log("next token = " + contacts.nextSyncToken);
        console.log("total number of items = " + contacts.totalItems);
        console.log("retrieved contacts = " + retrievedItems);
      }
    } else {
      // empty array, because contacts.connections is undefined.
      contactInformationArray = [];
    }

    // for a later synchronization ...
    // and finally update current array of contacts
    contactInformation = {
      synchronizationFlag: true,
      nextSyncToken: contacts.nextSyncToken,
      contactInformationArray: contactInformationArray,
    };
  } catch (error) {
    console.log("Synchronization failure: " + error);
    // perform a full retrieval of contacts
    contactInformation = populateAllConnectionsWithoutSyncToken();
  }
  return contactInformation;
}

/**
 * Gets a list of people in the user's contacts.
 * https://developers.google.com/people/api/rest/v1/people.connections/list
 *
 * We will start by using cached data if any.
 *
 * And, if there is no cached data, we will populate our cache with a full retrieval.
 * @param {string} nextSyncToken a value to limit data exhanged with contacts server
 */
function getConnections(nextSyncToken) {
  // store synchronization token and connections array (try to cache them)
  let contactInformation;
  console.log("Retrieve contact informations ...");
  if (nextSyncToken) {
    console.log("try synchronization ...");
    contactInformation = synchronizeConnections(nextSyncToken);
  } else {
    contactInformation = populateAllConnectionsWithoutSyncToken();
  }
  return contactInformation;
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
 *
 * Use google workspace api for Drive:
 * https://developers.google.com/apps-script/reference/drive/drive-app
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

  let currentFile = DriveApp.getFileById(currentDocument.getId());
  let newFile = currentFile.makeCopy(documentName);

  let newDocument = DocumentApp.openById(newFile.getId());

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
