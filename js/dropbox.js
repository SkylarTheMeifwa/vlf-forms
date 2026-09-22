// =========================================================
// DROPBOX CONNECTION
// =========================================================

const DROPBOX_APP_KEY = "6gxrjei4t4jm04m";


// =========================================================
// REDIRECT URI
// =========================================================

function getDropboxRedirectUri() {
    return window.location.origin + window.location.pathname;
}


// =========================================================
// PKCE
// =========================================================

function generateRandomString(length) {

    const characters =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    const randomValues =
        new Uint8Array(length);

    crypto.getRandomValues(randomValues);

    let result = "";

    for (let i = 0; i < length; i++) {
        result += characters[randomValues[i] % characters.length];
    }

    return result;
}


async function createCodeChallenge(codeVerifier) {

    const encoder = new TextEncoder();

    const data = encoder.encode(codeVerifier);

    const digest =
        await crypto.subtle.digest("SHA-256", data);

    return btoa(
        String.fromCharCode(...new Uint8Array(digest))
    )
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");
}


// =========================================================
// CONNECT TO DROPBOX
// =========================================================

async function connectDropbox() {

    const redirectUri =
        getDropboxRedirectUri();

    const codeVerifier =
        generateRandomString(64);

    const codeChallenge =
        await createCodeChallenge(codeVerifier);

    sessionStorage.setItem(
        "dropbox_code_verifier",
        codeVerifier
    );

    const authUrl =
        "https://www.dropbox.com/oauth2/authorize" +
        "?client_id=" +
        encodeURIComponent(DROPBOX_APP_KEY) +
        "&response_type=code" +
        "&code_challenge=" +
        encodeURIComponent(codeChallenge) +
        "&code_challenge_method=S256" +
        "&redirect_uri=" +
        encodeURIComponent(redirectUri);

    window.location.href = authUrl;
}


// =========================================================
// HANDLE DROPBOX CALLBACK
// =========================================================

async function handleDropboxRedirect() {

    const params =
        new URLSearchParams(window.location.search);

    const code =
        params.get("code");

    if (!code) {
        return;
    }

    const codeVerifier =
        sessionStorage.getItem(
            "dropbox_code_verifier"
        );

    if (!codeVerifier) {

        console.error(
            "Dropbox code verifier was not found."
        );

        return;
    }

    try {

        const response =
            await fetch(
                "https://api.dropboxapi.com/oauth2/token",
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/x-www-form-urlencoded"
                    },

                    body:
                        new URLSearchParams({

                            code: code,

                            grant_type:
                                "authorization_code",

                            client_id:
                                DROPBOX_APP_KEY,

                            code_verifier:
                                codeVerifier,

                            redirect_uri:
                                getDropboxRedirectUri()
                        })
                }
            );

        if (!response.ok) {

            const errorText =
                await response.text();

            throw new Error(errorText);
        }

        const tokenData =
            await response.json();

        sessionStorage.setItem(
            "dropbox_access_token",
            tokenData.access_token
        );

        sessionStorage.removeItem(
            "dropbox_code_verifier"
        );

        window.history.replaceState(
            {},
            document.title,
            window.location.pathname
        );

        console.log(
            "Dropbox connected successfully!"
        );

    } catch (error) {

        console.error(
            "Dropbox authorization failed:",
            error
        );
    }
}


// =========================================================
// CHECK CONNECTION
// =========================================================

function isDropboxConnected() {

    return Boolean(
        sessionStorage.getItem(
            "dropbox_access_token"
        )
    );
}


// ========================================================
// TEST SHARED FOLDER ACCESS
// ========================================================

async function getVLFFormsFiles() {

    const accessToken =
        sessionStorage.getItem("dropbox_access_token");

    if (!accessToken) {
        return null;
    }

    try {

        const response =
            await fetch(
                "https://api.dropboxapi.com/2/files/list_folder",
                {
                    method: "POST",

                    headers: {
                        "Authorization":
                            `Bearer ${accessToken}`,

                        "Content-Type":
                            "application/json"
                    },

                    body: JSON.stringify({
                        path: "/VLF Forms"
                    })
                }
            );

        if (!response.ok) {

            const errorText =
                await response.text();

            throw new Error(errorText);
        }

        const data =
            await response.json();

        return data.entries;

    } catch (error) {

        console.error(
            "Could not access VLF Forms:",
            error
        );

        return null;
    }
}


// =========================================================
// TEST DROPBOX
// =========================================================

async function testDropboxConnection() {

    const files =
        await getVLFFormsFiles();

    if (files === null) {

        alert(
            "Could not access the VLF Forms folder."
        );

        return;
    }

    const accessibleFiles =
        files
            .filter(
                entry => entry[".tag"] === "file"
            )
            .map(
                entry => entry.name
            );

    if (accessibleFiles.length === 0) {

        alert(
            "The VLF Forms folder is accessible, " +
            "but it contains no files."
        );

        return;
    }

    alert(
        "Accessible VLF Forms files:\n\n" +
        accessibleFiles.join("\n")
    );

    console.log(
        "Accessible VLF Forms files:",
        accessibleFiles
    );
}


// =========================================================
// CHECK TEMPLATE ACCESS
// =========================================================

async function isTemplateAccessible(templateName) {

    const files =
        await getVLFFormsFiles();

    if (files === null) {
        return false;
    }

    return files.some(
        entry =>
            entry[".tag"] === "file" &&
            entry.name === templateName
    );
}


// =========================================================
// OPEN FORM IF TEMPLATE IS ACCESSIBLE
// =========================================================

async function openFormIfAccessible(
    templateName,
    formURL
) {

    const accessToken =
        sessionStorage.getItem("dropbox_access_token");

    if (!accessToken) {

        alert(
            "Please connect your Dropbox account " +
            "before opening this form."
        );

        return;
    }

    const accessible =
        await isTemplateAccessible(templateName);

    if (!accessible) {

        alert(
            "The template required for this form " +
            "is not accessible in your Dropbox account."
        );

        return;
    }

    window.location.href = formURL;
}

async function downloadDropboxTemplate(templateName) {

    const accessToken =
        sessionStorage.getItem("dropbox_access_token");

    if (!accessToken) {
        throw new Error(
            "Dropbox is not connected."
        );
    }

    const response =
        await fetch(
            "https://content.dropboxapi.com/2/files/download",
            {
                method: "POST",

                headers: {
                    "Authorization":
                        `Bearer ${accessToken}`,

                    "Dropbox-API-Arg":
                        JSON.stringify({
                            path:
                                `/VLF Forms/${templateName}`
                        })
                }
            }
        );

    if (!response.ok) {

        const errorText =
            await response.text();

        throw new Error(errorText);
    }

    return await response.arrayBuffer();
}

// =========================================================
// STARTUP
// =========================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        handleDropboxRedirect();

    }
);
