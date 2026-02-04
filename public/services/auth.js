/**
 * Auth service: MSAL-based sign-in and Graph API calls.
 * No backend; SPA only. Redirect URI must match taskpane URL exactly.
 */
const AuthService = {
  msalConfig: {
    auth: {
      clientId: '4279250e-13dd-44da-a98a-63badacdbaf3',
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: typeof window !== 'undefined' ? window.location.origin + '/taskpane.html' : '',
    },
    cache: {
      cacheLocation: 'localStorage',
      storeAuthStateInCookie: false,
    },
  },

  scopes: [
    'User.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Mail.Send.Shared',
    'Mail.Read.Shared',
    'MailboxSettings.ReadWrite',  // Required for creating/managing email categories
  ],

  msalInstance: null,
  currentAccount: null,

  async initialize() {
    if (typeof msal === 'undefined') {
      console.error('MSAL library not loaded');
      return false;
    }
    try {
      this.msalInstance = new msal.PublicClientApplication(this.msalConfig);

      const response = await this.msalInstance.handleRedirectPromise();
      if (response) {
        this.currentAccount = response.account;
        console.log('Logged in via redirect:', this.currentAccount.username);
      } else {
        const accounts = this.msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          this.currentAccount = accounts[0];
          console.log('Found existing account:', this.currentAccount.username);
        }
      }
      return true;
    } catch (error) {
      console.error('MSAL initialization error:', error);
      return false;
    }
  },

  isSignedIn() {
    return !!this.currentAccount;
  },

  getUser() {
    if (!this.currentAccount) return null;
    return {
      email: this.currentAccount.username,
      name: this.currentAccount.name || this.currentAccount.username,
    };
  },

  async signIn() {
    if (!this.msalInstance) throw new Error('MSAL not initialized');
    try {
      const response = await this.msalInstance.loginPopup({
        scopes: this.scopes,
        prompt: 'select_account',
      });
      this.currentAccount = response.account;
      return this.getUser();
    } catch (error) {
      console.error('Sign in error:', error);
      throw error;
    }
  },

  async signOut() {
    if (!this.msalInstance || !this.currentAccount) return;
    try {
      await this.msalInstance.logoutPopup({
        account: this.currentAccount,
        postLogoutRedirectUri: this.msalConfig.auth.redirectUri,
      });
      this.currentAccount = null;
    } catch (error) {
      console.error('Sign out error:', error);
      throw error;
    }
  },

  async getAccessToken() {
    if (!this.msalInstance) throw new Error('MSAL not initialized');
    if (!this.currentAccount) throw new Error('No user signed in');

    const tokenRequest = { scopes: this.scopes, account: this.currentAccount };
    try {
      const response = await this.msalInstance.acquireTokenSilent(tokenRequest);
      return response.accessToken;
    } catch (error) {
      if (error.name === 'InteractionRequiredAuthError' || (error.errorCode && error.errorCode === 'interaction_required')) {
        const response = await this.msalInstance.acquireTokenPopup(tokenRequest);
        return response.accessToken;
      }
      throw error;
    }
  },

  async graphRequest(endpoint, options = {}) {
    const token = await this.getAccessToken();
    const url = endpoint.startsWith('https://')
      ? endpoint
      : 'https://graph.microsoft.com/v1.0' + (endpoint.startsWith('/') ? endpoint : '/' + endpoint);

    const response = await fetch(url, {
      ...options,
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error?.message || response.statusText);
    }

    if (response.status === 204 || response.status === 202) return null;
    const contentType = response.headers.get('content-type');
    if (contentType && contentType.includes('application/json')) {
      const text = await response.text();
      if (text) return JSON.parse(text);
    }
    return null;
  },
};
