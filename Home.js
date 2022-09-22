
(async function () {
    
    
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
            Office.context.document.settings.saveAsync(); 

            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                return;
            }

            $('#login').click(login);


        });
    };
    

    async function login() {
        try {
            const email = document.getElementById("email").value;
            const password = document.getElementById("password").value;
            const getCall = await fetch("https://localhost:7018/users/Authenticate", {
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
            localStorage.setItem("JWT", token.token);
            const tenantId = token.user.tenantid;
            const userId = token.user.id;
            localStorage.setItem("userId", userId);
            localStorage.setItem("tenantId", tenantId);
            location.assign('/users.html')
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



    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
