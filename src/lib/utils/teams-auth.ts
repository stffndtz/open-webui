import * as microsoftTeams from '@microsoft/teams-js';
import { getSessionUser } from '$lib/apis/auths';
import { WEBUI_BASE_URL } from '$lib/constants';

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

	/**
	 * Initialize the Teams SDK
	 */
	async initialize(): Promise<boolean> {
		if (this.isInitialized) {
			return true;
		}

		try {
			await microsoftTeams.app.initialize();
			this.isInitialized = true;
			console.log('Teams SDK initialized successfully');
			return true;
		} catch {
			console.log('Teams SDK initialization failed');
			return false;
		}
	}

	/**
	 * Check if we're running in a Teams environment
	 */
	async isInTeams(): Promise<boolean> {
		try {
			await this.initialize();
			const context = await microsoftTeams.app.getContext();
			return context.app.host.name === 'Teams';
		} catch {
			return false;
		}
	}

	/**
	 * Get Teams context information
	 */
	async getContext(): Promise<microsoftTeams.app.Context | null> {
		try {
			await this.initialize();
			return await microsoftTeams.app.getContext();
		} catch (error) {
			console.error('Failed to get Teams context:', error);
			return null;
		}
	}

	/**
	 * Attempt silent authentication using Teams SSO
	 */
	async attemptSilentAuth(): Promise<TeamsAuthResult> {
		try {
			await this.initialize();

			// Try to get the user's access token silently
			const authTokenRequest: microsoftTeams.authentication.AuthTokenRequest = {
				successCallback: (result: string) => {
					console.log('Silent auth successful');
					return { success: true, token: result };
				},
				failureCallback: (reason: string) => {
					console.log('Silent auth failed:', reason);
					return { success: false, error: reason };
				},
				resources: [`${WEBUI_BASE_URL}/api`]
			};

			return new Promise((resolve) => {
				microsoftTeams.authentication.getAuthToken(authTokenRequest);
				// Note: The callbacks will handle the resolution
				resolve({ success: false, error: 'Silent auth not implemented in this version' });
			});
		} catch (error) {
			console.error('Silent auth error:', error);
			return { success: false, error: 'Silent authentication failed' };
		}
	}

	/**
	 * Perform interactive authentication using Teams authentication dialog
	 */
	async performInteractiveAuth(): Promise<TeamsAuthResult> {
		if (this.authPromise) {
			return this.authPromise;
		}

		this.authPromise = this._performInteractiveAuth();
		return this.authPromise;
	}

	private async _performInteractiveAuth(): Promise<TeamsAuthResult> {
		try {
			await this.initialize();

			console.log('Starting Teams interactive authentication...');
			console.log('Authentication URL:', `${WEBUI_BASE_URL}/oauth/microsoft/login?teams=true`);

			// Start authentication flow
			const authResult = await microsoftTeams.authentication.authenticate({
				url: `${WEBUI_BASE_URL}/oauth/microsoft/login?teams=true`,
				width: 600,
				height: 535
			});

			console.log('Teams authentication result:', authResult);

			if (authResult) {
				// Try to get user session with the token
				try {
					const sessionUser = await getSessionUser(authResult);
					return {
						success: true,
						token: authResult,
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
					error: 'Authentication was cancelled or failed'
				};
			}
		} catch (error) {
			console.error('Interactive auth error:', error);
			console.error('Error details:', {
				message: error instanceof Error ? error.message : 'Unknown error',
				stack: error instanceof Error ? error.stack : undefined,
				name: error instanceof Error ? error.name : 'Unknown'
			});
			return {
				success: false,
				error: error instanceof Error ? error.message : 'Authentication failed'
			};
		} finally {
			this.authPromise = null;
		}
	}

	/**
	 * Handle authentication with iframe communication fallback
	 */
	async authenticateWithIframeFallback(options: TeamsAuthOptions = {}): Promise<TeamsAuthResult> {
		try {
			// First try silent authentication
			if (options.enableSilentAuth !== false) {
				const silentResult = await this.attemptSilentAuth();
				if (silentResult.success) {
					return silentResult;
				}
			}

			// If silent auth fails or is disabled, try interactive auth
			if (!options.forceInteractive) {
				const interactiveResult = await this.performInteractiveAuth();
				if (interactiveResult.success) {
					return interactiveResult;
				}
			}

			// Fallback to iframe-based authentication
			return await this.authenticateWithIframe(options);
		} catch (error) {
			console.error('Authentication with iframe fallback failed:', error);
			return {
				success: false,
				error: error instanceof Error ? error.message : 'Authentication failed'
			};
		}
	}

	/**
	 * Authenticate using iframe with postMessage communication
	 */
	private async authenticateWithIframe(options: TeamsAuthOptions = {}): Promise<TeamsAuthResult> {
		return new Promise((resolve) => {
			// Create iframe for authentication
			const iframe = document.createElement('iframe');
			iframe.style.position = 'fixed';
			iframe.style.top = '50%';
			iframe.style.left = '50%';
			iframe.style.transform = 'translate(-50%, -50%)';
			iframe.style.width = '600px';
			iframe.style.height = '535px';
			iframe.style.border = 'none';
			iframe.style.borderRadius = '8px';
			iframe.style.boxShadow = '0 4px 20px rgba(0, 0, 0, 0.3)';
			iframe.style.zIndex = '9999';
			iframe.style.backgroundColor = 'white';

			const redirectUri = options.redirectUri || `${WEBUI_BASE_URL}/oauth/microsoft/teams-callback`;
			iframe.src = `${WEBUI_BASE_URL}/oauth/microsoft/login?teams=true&redirect_uri=${encodeURIComponent(redirectUri)}`;

			// Add iframe to page
			document.body.appendChild(iframe);

			// Listen for messages from the iframe
			const messageHandler = (event: MessageEvent) => {
				if (event.data?.type === 'teams-auth-success') {
					// Clean up
					document.body.removeChild(iframe);
					window.removeEventListener('message', messageHandler);

					// Process the token
					getSessionUser(event.data.token)
						.then((sessionUser) => {
							resolve({
								success: true,
								token: event.data.token,
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
				} else if (event.data?.type === 'teams-auth-error') {
					// Clean up
					document.body.removeChild(iframe);
					window.removeEventListener('message', messageHandler);

					resolve({
						success: false,
						error: event.data.error || 'Authentication failed'
					});
				}
			};

			window.addEventListener('message', messageHandler);

			// Add close button
			const closeButton = document.createElement('button');
			closeButton.innerHTML = '×';
			closeButton.style.position = 'fixed';
			closeButton.style.top = 'calc(50% - 267px)';
			closeButton.style.left = 'calc(50% + 300px)';
			closeButton.style.width = '30px';
			closeButton.style.height = '30px';
			closeButton.style.border = 'none';
			closeButton.style.borderRadius = '50%';
			closeButton.style.backgroundColor = '#ff4444';
			closeButton.style.color = 'white';
			closeButton.style.fontSize = '20px';
			closeButton.style.cursor = 'pointer';
			closeButton.style.zIndex = '10000';
			closeButton.onclick = () => {
				document.body.removeChild(iframe);
				document.body.removeChild(closeButton);
				window.removeEventListener('message', messageHandler);
				resolve({
					success: false,
					error: 'Authentication cancelled'
				});
			};

			document.body.appendChild(closeButton);

			// Timeout after 5 minutes
			setTimeout(() => {
				if (document.body.contains(iframe)) {
					document.body.removeChild(iframe);
					document.body.removeChild(closeButton);
					window.removeEventListener('message', messageHandler);
					resolve({
						success: false,
						error: 'Authentication timeout'
					});
				}
			}, 300000);
		});
	}

	/**
	 * Sign out the user from Teams
	 */
	async signOut(): Promise<void> {
		try {
			await this.initialize();
			await microsoftTeams.authentication.notifySuccess('');
		} catch (error) {
			console.error('Teams sign out error:', error);
		}
	}

	/**
	 * Get the current user's Teams information
	 */
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

// Export singleton instance
export const teamsAuth = new TeamsAuthManager();
