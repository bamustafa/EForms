const clientId = "83a6a7e1-0994-44e9-a357-80e9bc5ef9ba"; // Replace with your Azure AD App ID
const tenantId = "Y4cef5c93-f877-4a8d-9bb9-f02c18536c88"; // Replace with your Azure Tenant ID
const authority = `https://login.microsoftonline.com/${tenantId}`;
const graphEndpoint = "https://graph.microsoft.com/v1.0";
const username = "Bilal.Mustafa@dar.com"; // Replace with the target username

let msalUrl = 'https://alcdn.msauth.net/browser/2.30.0/js/msal-browser.min.js'


var onRender = async function(layout, moduleName, formType) {

    //updateTotal();

    //getUserPhoto();

    await getModulesCount();

}

function updateTotal() {

    let a = parseFloat(fd.field('A').value) || 0;
    let b = parseFloat(fd.field('B').value) || 0;


               let total1 = a * b;


    fd.field('COUNT').value = total1.toFixed(2);
         }

// Authenticate and get token
const getAccessToken = async (msal) => {

    debugger;
    if(!msal) return;
    const msalConfig = {
        auth: {
            clientId: clientId,
            authority: authority,
            redirectUri: window.location.origin
        }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const loginRequest = {
        scopes: ["User.Read"]
    };

    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        const tokenResponse = await msalInstance.acquireTokenSilent(loginRequest);
        return tokenResponse.accessToken;
    } catch (error) {
        console.error("Authentication Error:", error);
        return null;
    }

}

// Fetch user photo
async function getUserPhoto() {


    _spComponentLoader.loadScript(msalUrl).then(getAccessToken);

    // const token = await getAccessToken();
    // if (!token) return;

    const headers = new Headers();
    headers.append("Authorization", `Bearer ${token}`);

    try {
        debugger;
        const response = await fetch(`${graphEndpoint}/users/${username}/photo/$value`, { headers });
        if (response.ok) {
            const blob = await response.blob();
            const imageUrl = URL.createObjectURL(blob);
            document.getElementById("profilePhoto").src = imageUrl;
        } else {
            console.warn("No photo found for user.");
        }
    } catch (error) {
        console.error("Error fetching user photo:", error);
    }
}

let apiUrl = `${_spPageContextInfo.siteAbsoluteUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeImageAsBinary&Emails=ali.hsleiman@dar.com,bilal.mustafa@dar.com`
async function fetchResult(apiUrl, method, data) {
    try {
        const response = await fetch(apiUrl, {
            method: method, // "GET", "POST", "PUT", "PATCH", or "DELETE"
            headers: {
                "Content-Type": "application/json",
                // "Authorization": "Bearer YOUR_ACCESS_TOKEN", // Uncomment if authentication is needed
            },
            body: data ? JSON.stringify(data) : null,
        });

        const result = await response.json(); // Attempt to parse JSON response

        if (!response.ok) {
            throw new Error(`Error ${response.status}: ${result.message || "Request failed"}`);
        }

        console.log("Request successful:", result);
        return result;
    } catch (error) {
        console.error("Fetch error:", error.message);
        return { success: false, error: error.message };
    }
}


let getModulesCount = async function () {
    //let apiUrl = `${_spPageContextInfo.siteAbsoluteUrl}/_layouts/15/PCW/APICalls/SPUtils.aspx?command=GetTotalOpenPerModule`

    debugger;
    //const user = await pnp.sp.web.currentUser.get();
    //let apiUrl = `${_spPageContextInfo.siteAbsoluteUrl}/_layouts/15/PCW/APICalls/SPUtils.aspx?Command=GetConstTasksByUser&username=${user.Title}`

    //let apiUrl = `${_spPageContextInfo.siteAbsoluteUrl}/_layouts/15/PCW/APICalls/SPUtils.aspx?command=GetTopContributors`

    let apiUrl = `${_spPageContextInfo.siteAbsoluteUrl}/_layouts/15/PCW/APICalls/SPUtils.aspx?command=GetRecentActivities`

    await fetch(apiUrl)
        .then(async response => {
            return response.text()
         })
         .then(async data => {
            let result = JSON.parse(data);
            console.log(result);
         });
}


