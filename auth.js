/* Sign-in: uses Outlook's own sign-in (nested app auth) so there's no extra login.
   Falls back to a pop-up sign-in on older Outlook builds. */
(function () {
  "use strict";
  var cfg = window.TCC_CONFIG;
  var SCOPES = ["Files.ReadWrite.All", "User.Read"];
  var pca = null, mode = null;

  async function init() {
    if (pca) return;
    var msalConfig = {
      auth: { clientId: cfg.clientId, authority: "https://login.microsoftonline.com/" + cfg.tenantId },
      cache: { cacheLocation: "localStorage" }
    };
    var naa = Office.context.requirements && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
    if (naa) {
      pca = await msal.createNestablePublicClientApplication(msalConfig);
      mode = "naa";
    } else {
      msalConfig.auth.redirectUri = location.origin + location.pathname.replace(/[^/]*$/, "") + "auth.html";
      pca = await msal.createStandardPublicClientApplication(msalConfig);
      mode = "dialog";
    }
  }

  function dialogLogin() {
    return new Promise(function (resolve, reject) {
      var url = location.origin + location.pathname.replace(/[^/]*$/, "") + "auth.html";
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error("Couldn't open sign-in window"));
        var d = res.value;
        d.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          d.close();
          var m = JSON.parse(arg.message);
          m.token ? resolve(m.token) : reject(new Error(m.error || "Sign-in failed"));
        });
        d.addEventHandler(Office.EventType.DialogEventReceived, function () { reject(new Error("Sign-in window closed")); });
      });
    });
  }

  async function getToken() {
    await init();
    var req = { scopes: SCOPES };
    var account = pca.getActiveAccount() || pca.getAllAccounts()[0];
    if (account) req.account = account;
    if (mode === "dialog" && !account) return dialogLogin();
    try {
      var r = await pca.acquireTokenSilent(req);
      return r.accessToken;
    } catch (e) {
      if (mode === "naa") {
        var p = await pca.acquireTokenPopup(req);
        pca.setActiveAccount(p.account);
        return p.accessToken;
      }
      return dialogLogin();
    }
  }

  window.TCCAuth = { getToken: getToken, SCOPES: SCOPES };
})();
