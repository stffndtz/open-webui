import * as microsoftTeams from '@microsoft/teams-js';
import { getSessionUser } from '$lib/apis/auths';

export interface TeamsAuthResult {
	success: boolean;
	token?: string;
	error?: string;
	user?: unknown;
}

export interface TeamsAuthOptions {
	enableSilentAuth?: boolean;
	forceInteractive?: boolean;
	redirectUri?: string;
}

class TeamsAuthManager {
	private isInitialized = false;
	private authPromise: Promise<TeamsAuthResult> | null = null;

	async initialize(): Promise<boolean> {
		if (this.isInitialized) {
			return true;
		}

		try {
			// Initialize Teams SDK
			await microsoftTeams.app.initialize();

			this.isInitialized = true;
			console.log('Teams SDK initialized successfully');
			return true;
		} catch (error) {
			console.error('Teams SDK initialization failed:', error);
			return false;
		}
	}

	async isInTeams(): Promise<boolean> {
		try {
			// Check if we're in a browser environment that supports Teams
			if (typeof window === 'undefined' || !window.parent || window.parent === window) {
				console.log('Not in a Teams iframe environment');
				return false;
			}

			await this.initialize();
			const context = await microsoftTeams.app.getContext();
			console.log('Teams context check:', {
				host: context.app.host.name,
				sessionId: context.app.sessionId,
				theme: context.app.theme,
				locale: context.app.locale,
				user: context.user ? 'User available' : 'No user',
				page: context.page
			});

			// Check if we're in Teams and have proper configuration
			const isTeams = context.app.host.name === 'Teams';
			console.log('Is in Teams:', isTeams);

			if (isTeams && !context.user) {
				console.warn(
					'In Teams but no user context available - this might indicate a configuration issue'
				);
			}

			return isTeams;
		} catch (error) {
			console.error('Failed to check Teams environment:', error);
			return false;
		}
	}

	async getContext(): Promise<microsoftTeams.app.Context | null> {
		try {
			await this.initialize();
			return await microsoftTeams.app.getContext();
		} catch (error) {
			console.error('Failed to get Teams context:', error);
			return null;
		}
	}

	async getAuthToken(): Promise<TeamsAuthResult> {
		try {
			// Check if we're actually in Teams
			const isInTeams = await this.isInTeams();
			if (!isInTeams) {
				console.log('Not in Teams environment, skipping getAuthToken');
				return {
					success: false,
					error: 'Not in Teams environment'
				};
			}

			console.log('Getting auth token from Microsoft Teams...');

			// Follow the exact pattern from Microsoft documentation
			return new Promise((resolve) => {
				microsoftTeams.authentication
					.getAuthToken()
					.then((token) => {
						console.log('Auth token received successfully');

						if (token) {
							// Try to get user session with the token
							getSessionUser(token)
								.then((sessionUser) => {
									resolve({
										success: true,
										token: token,
										user: sessionUser
									});
								})
								.catch((error) => {
									console.error('Failed to get session user:', error);
									resolve({
										success: false,
										error: 'Failed to validate authentication token'
									});
								});
						} else {
							resolve({
								success: false,
								error: 'No token received from Teams'
							});
						}
					})
					.catch((error) => {
						console.error('getAuthToken error:', error);
						console.error('Error details:', {
							message: error instanceof Error ? error.message : 'Unknown error',
							stack: error instanceof Error ? error.stack : undefined,
							name: error instanceof Error ? error.name : 'Unknown'
						});
						resolve({
							success: false,
							error: error instanceof Error ? error.message : 'Authentication failed'
						});
					});
			});
		} catch (error) {
			console.error('getAuthToken outer error:', error);
			return {
				success: false,
				error: error instanceof Error ? error.message : 'Authentication failed'
			};
		}
	}

	async authenticateWithSSO(): Promise<TeamsAuthResult> {
		if (this.authPromise) {
			return this.authPromise;
		}

		this.authPromise = this._authenticateWithSSO();
		return this.authPromise;
	}

	private async _authenticateWithSSO(): Promise<TeamsAuthResult> {
		try {
			console.log('Starting SSO authentication flow...');

			// Follow the exact pattern from Microsoft documentation
			return new Promise((resolve) => {
				console.log('Initializing Teams SDK...');
				microsoftTeams.app
					.initialize()
					.then(() => {
						console.log('Teams SDK initialized, getting context...');
						return microsoftTeams.app.getContext();
					})
					.then((context) => {
						console.log('Teams context received:', {
							app: context.app,
							user: context.user ? 'User available' : 'No user',
							page: context.page
						});
						console.log('Getting auth token...');
						return this.getClientSideToken();
					})
					.then((clientSideToken) => {
						console.log('Client-side token received, validating...');
						console.log('Token received:', clientSideToken ? 'Yes' : 'No');
						return this.validateToken(clientSideToken);
					})
					.then((result) => {
						console.log('Token validation complete');
						console.log('Result:', result);
						resolve(result);
					})
					.catch((error) => {
						console.error('SSO authentication error:', error);
						console.error('Error details:', {
							message: error instanceof Error ? error.message : 'Unknown error',
							stack: error instanceof Error ? error.stack : undefined,
							name: error instanceof Error ? error.name : 'Unknown'
						});
						resolve({
							success: false,
							error: error instanceof Error ? error.message : 'SSO authentication failed'
						});
					});
			});
		} catch (error) {
			console.error('SSO authentication outer error:', error);
			return {
				success: false,
				error: error instanceof Error ? error.message : 'SSO authentication failed'
			};
		} finally {
			this.authPromise = null;
		}
	}

	private async getClientSideToken(): Promise<string> {
		return new Promise((resolve, reject) => {
			console.log('Getting auth token from Microsoft Teams...');
			console.log('Teams SDK version:', microsoftTeams.version);
			console.log('Current window location:', window.location.href);
			console.log('Parent window:', window.parent !== window ? 'Has parent' : 'No parent');

			// Try to get auth token with specific parameters
			const authTokenRequest = {
				prompt: 'none', // Try to get token without prompting
				successCallback: (result: string) => {
					console.log('Auth token received successfully via callback');
					console.log('Token length:', result ? result.length : 0);
					resolve(result);
				},
				failureCallback: (reason: string) => {
					console.error('Auth token failed via callback:', reason);

					// Handle specific embedded browser error
					if (
						reason.includes('embedded browser') ||
						reason.includes('URL was unable to be opened')
					) {
						console.warn(
							'Embedded browser error detected, trying alternative authentication method'
						);
						// Try alternative authentication method
						this.tryAlternativeAuth(resolve, (error: string) => reject(error));
					} else {
						reject('Error getting token: ' + reason);
					}
				}
			};

			try {
				microsoftTeams.authentication.getAuthToken(authTokenRequest);
			} catch (error) {
				console.error('Error calling getAuthToken:', error);
				reject('Error calling getAuthToken: ' + error);
			}
		});
	}

	private async tryAlternativeAuth(
		resolve: (value: string) => void,
		_reject: (reason: string) => void
	): Promise<void> {
		try {
			console.log('Trying alternative authentication method...');

			// Try to get token with login prompt instead of silent auth
			const tokenRequest = {
				prompt: 'login', // Force login prompt instead of silent auth
				successCallback: (result: string) => {
					console.log('Alternative auth successful');
					resolve(result);
				},
				failureCallback: (reason: string) => {
					console.error('Alternative auth failed:', reason);
					_reject('Alternative authentication failed: ' + reason);
				}
			};

			microsoftTeams.authentication.getAuthToken(tokenRequest);
		} catch (error) {
			console.error('Alternative auth error:', error);
			_reject('Alternative authentication error: ' + error);
		}
	}

	private async validateToken(token: string): Promise<TeamsAuthResult> {
		try {
			// Try to get user session with the token
			const sessionUser = await getSessionUser(token);
			return {
				success: true,
				token: token,
				user: sessionUser
			};
		} catch (error) {
			console.error('Failed to validate token:', error);
			return {
				success: false,
				error: 'Failed to validate authentication token'
			};
		}
	}

	async signOut(): Promise<void> {
		try {
			await this.initialize();
			await microsoftTeams.authentication.notifySuccess('');
		} catch (error) {
			console.error('Teams sign out error:', error);
		}
	}

	async getCurrentUser(): Promise<unknown> {
		try {
			await this.initialize();
			const context = await microsoftTeams.app.getContext();
			return {
				id: context.user?.id,
				displayName: context.user?.displayName,
				email: context.user?.userPrincipalName,
				tenantId: context.user?.tenant
			};
		} catch (error) {
			console.error('Failed to get current user:', error);
			return null;
		}
	}
}

export const teamsAuth = new TeamsAuthManager();
