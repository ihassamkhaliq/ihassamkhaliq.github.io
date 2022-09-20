
(async function () {
    
    
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    var messageBanner;
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync(); 
    
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");

                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $('#login').click(login);


        });
    };



    


    async function login() {
        try {
            const email = document.getElementById("email").value;
            const password = document.getElementById("password").value;
            const getCall = await fetch("https://localhost:7018/tenant/Authenticate", {
                "method": "POST",
                body: JSON.stringify({
                    email: email,
                    password: password
                }),
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            })
            const token = await getCall.json();
            if(token.status === 200){
            localStorage.setItem("JWT", token.token);
            const userDetails = token.user.id;
            localStorage.setItem("userDetail", userDetails);
            location.assign('/users.html')
        }else{
            showNotification("Invalid Email","Your Email is incorrect please Sign up First Please!")
        }
        } catch (error) {
            showNotification("Invalid Email","Your Email is incorrect please Sign up First Please!")
        }
    }

    const token = localStorage.getItem("JWT")
    if (token === null) {
        return
    }else if(token === "null"){
        return
    }
    else if(token === "undefined"){
        return
    } else{
        location.assign('/users.html')
    }

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }



    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
