// Import MSAL library
import * as msal from '@azure/msal-browser';

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: '3143790b-b33c-4c83-b845-fff06a435a51', // Zastąp swoim Client ID
        authority: 'c906d678-a366-4a55-8042-625ef63569d1', // Zastąp swoim Tenant ID
        redirectUri: 'https://localhost:3000', // URI przekierowania, dostosuj do środowiska produkcyjnego
    }
};

// Initialize MSAL Instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Function to authenticate and fetch user data from Microsoft Graph API
async function getUserData() {
    try {
        // Log in user
        const loginResponse = await msalInstance.loginPopup({
            scopes: ['User.Read']
        });

        // Acquire token silently
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ['User.Read'],
            account: loginResponse.account
        });

        // Fetch user data from Microsoft Graph API
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`
            }
        });

        const userData = await response.json();
        return userData;
    } catch (error) {
        console.error('Error fetching user data:', error);
        throw error;
    }
}

// Function to insert the footer into the email body
async function insertFooter() {
    try {
        // Fetch user data
        const userData = await getUserData();

        // Dynamic HTML footer template
        const footerHtml = `
            <div style="font-family:Calibri, Arial;font-size:9pt;line-height:140%;color:#595959;">
                <span style="font-size:14pt;color:#DF292F;">${userData.displayName}</span><br />
                <span style="font-size:12pt;">${userData.jobTitle || 'Stanowisko nieznane'}</span><br /><br />
                <a href="https://familijna.pl" style="color:#595959;text-decoration: none;">
                    <span style="color:#DF292F;">www.</span>familijna.pl
                </a> <span style="color:#DF292F;">email: </span>
                <a href="mailto:${userData.mail}" style="color:#595959;text-decoration: none;">${userData.mail}</a><br />
                <span style="color:#DF292F;">tel.</span> ${userData.businessPhones ? userData.businessPhones[0] : 'Brak danych'} 
                <span style="color:#DF292F;">kom.</span> ${userData.mobilePhone || 'Brak danych'}
            </div>
        `;

        // Insert footer into email body
        Office.context.mailbox.item.body.setAsync(
            footerHtml,
            { coercionType: Office.CoercionType.Html },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('Stopka została wstawiona.');
                } else {
                    console.error('Błąd podczas wstawiania stopki:', result.error.message);
                }
            }
        );
    } catch (error) {
        console.error('Error inserting footer:', error);
    }
}

// Initialize Office.js
Office.initialize = function (reason) {
    $(document).ready(function () {
        // Bind the insertFooter function to the button in taskpane.html
        $('#addFooter').click(function () {
            insertFooter();
        });
    });
};
