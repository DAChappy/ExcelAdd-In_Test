(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $("#getIDToken").on("click", getIDToken);
        });
    };

    async function getIDToken() {
        try {
            let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
                allowSignInPrompt: true,
            });
            let userToken = jwt_decode(userTokenEncoded);
            document.getElementById("userInfo").innerHTML =
                "name: " +
                userToken.name +
                "<br>email: " +
                userToken.preferred_username +
                "<br>id: " +
                userToken.oid;
            console.log(userToken);
        } catch (error) {
            document.getElementById("userInfo").innerHTML =
                "An error occurred. <br>Name: " +
                error.name +
                "<br>Code: " +
                error.code +
                "<br>Message: " +
                error.message;
            console.log(error);
        }
    }
})();