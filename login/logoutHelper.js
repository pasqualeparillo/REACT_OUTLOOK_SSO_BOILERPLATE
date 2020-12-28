let logoutDialog = Office.Dialog;
const dialogLogoutUrl =
  location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + "/logout/logout.html";

export const logoutFromO365 = async displayError => {
  Office.context.ui.displayDialogAsync(dialogLogoutUrl, { height: 80, width: 30 }, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      displayError(`${result.error.code} ${result.error.message}`);
    } else {
      logoutDialog = result.value;
      logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
      logoutDialog.addEventHandler(Office.EventType.DialogEventReceived, processLogoutDialogEvent);
    }
  });

  const processLogoutMessage = () => {
    logoutDialog.close();
    setState({ authStatus: "notLoggedIn", headerMessage: "Welcome" });
  };

  const processLogoutDialogEvent = arg => {
    processDialogEvent(arg, setState, displayError);
  };
};

const processDialogEvent = arg => {
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
