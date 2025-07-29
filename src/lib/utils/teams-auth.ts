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
				locale: context.app.locale
			});
			return context.app.host.name === 'Teams';
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
			await this.initialize();

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
			console.log('Teams context:', await this.getContext());

			const token = await microsoftTeams.authentication.getAuthToken();
			console.log('Auth token received successfully');

			if (token) {
				// Try to get user session with the token
				try {
					const sessionUser = await getSessionUser(token);
					return {
						success: true,
						token: token,
						user: sessionUser
					};
				} catch (error) {
					console.error('Failed to get session user:', error);
					return {
						success: false,
						error: 'Failed to validate authentication token'
					};
				}
			} else {
				return {
					success: false,
					error: 'No token received from Teams'
				};
			}
		} catch (error) {
			console.error('getAuthToken error:', error);
			console.error('Error details:', {
				message: error instanceof Error ? error.message : 'Unknown error',
				stack: error instanceof Error ? error.stack : undefined,
				name: error instanceof Error ? error.name : 'Unknown'
			});
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
			await this.initialize();

			// Use the proper SSO pattern as recommended by Microsoft
			return await this.getAuthToken();
		} catch (error) {
			console.error('SSO authentication error:', error);
			return {
				success: false,
				error: error instanceof Error ? error.message : 'SSO authentication failed'
			};
		} finally {
			this.authPromise = null;
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
