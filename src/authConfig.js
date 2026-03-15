export const msalConfig = {
  auth: {
    clientId: "6556629b-4c4b-4681-b5ae-2ad0a457c994",
    authority: "https://login.microsoftonline.com/b1f151b3-c451-42ce-86d8-4c8002af3c0b",
    redirectUri: window.location.origin,
  }
};

export const loginRequest = {
  scopes: ["User.Read"]
};
