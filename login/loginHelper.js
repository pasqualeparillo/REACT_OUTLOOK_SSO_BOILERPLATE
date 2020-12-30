import axios from "axios";

let loginDialog = Office.Dialog;
const dialogLoginUrl =
  location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + "/login/login.html";

export const signInO365 = async (displayError, setState) => {
  console.log({ authStatus: "loginInProcess" });
  setState({ authStatus: "loginInProcess" });
  await Office.context.ui.displayDialogAsync(dialogLoginUrl, { height: 80, width: 30 }, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      displayError(`${result.error.code} ${result.error.message}`);
    } else {
      loginDialog = result.value;
      loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
      loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
    }
  });

  async function processLoginMessage(arg) {
    let messageFromDialog = JSON.parse(arg.message);
    setState(messageFromDialog.result);
    if (messageFromDialog.status === "success") {
      // We now have a valid access token.
      loginDialog.close();
      let response = getGraphToken("https://graph.microsoft.com/v1.0/me/", messageFromDialog.result);
      console.log(response);
      //const response = await sso.makeGraphApiCall(messageFromDialog.result);
      // console.log(JSON.stringify(response));
    } else {
      // Something went wrong with authentication or the authorization of the web application.
      loginDialog.close();
      sso.showMessage(JSON.stringify(messageFromDialog.error.toString()));
    }
  }

  const processLoginDialogEvent = arg => {
    console.log(JSON.stringify(arg) + " processed");
    processDialogEvent(arg);
  };
};

export const getGraphToken = async (url, accesstoken) => {
  const response = await axios({
    url: url,
    method: "get",
    headers: { Authorization: `Bearer ${accesstoken}` }
  });
  return response;
};

const processDialogEvent = (arg, displayError) => {
  switch (arg.error) {
    case 12002:
      displayError(
        "The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid."
      );
      break;
    case 12003:
      displayError("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");
      break;
    case 12006:
      // 12006 means that the user closed the dialog instead of waiting for it to close.
      // It is not known if the user completed the login or logout, so assume the user is
      // logged out and revert to the app's starting state. It does no harm for a user to
      // press the login button again even if the user is logged in.
      console.log({ authStatus: "notLoggedIn", headerMessage: "Welcome" });
      break;
    default:
      displayError("Unknown error in dialog box.");
      break;
  }
};
