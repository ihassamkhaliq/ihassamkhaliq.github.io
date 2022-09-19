


(async function () {

    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();

    var messageBanner;
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');

                return;
            }

            // Add a click event handler for the highlight button.
            $('#logout').click(logout);

        });
    };
    // The initialize function must be run each time a new page is loaded.





    setInterval(function () {
        functionality() 
    }, 5000);

    // Getting User Data

    const token = localStorage.getItem("JWT")
    console.log("users side tokens ====>>>>",token)
    console.log("This is Token ====>>>",typeof(token))
    const userDetail = localStorage.getItem("userDetail")

    if(token === "undefined"){
        console.log("I'm at the undefined of users side")
        location.assign("/Home.html")
    }else if(token === null){
        console.log("I'm null from users side")
        location.assign("/Home.html")
    }

    //Will bring Data for Get Calls

    async function getData(url) {
        return await fetch(url, {
            "method": "GET",
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Authorization': `Bearer ${token}`
            }
        })
    }

    // Get the Subscription of User
    const subscriptionDetail = await getData(`https://localhost:7018/tenant/${userDetail}/subscription/all`)
    const userSub = await subscriptionDetail.json();




        // Get all the Rules

        const rulesDetail = await getData(`https://localhost:7018/rules`)
        const rulesInfo = await rulesDetail.json();

        // Get all the Detectors

        const detectorsDetail = await getData(`https://localhost:7018/detectors`)
        const detectorsInfo = await detectorsDetail.json();


        // Get all dictionaries

        const dictionaryDetail = await getData(`https://localhost:7018/dictionaries`)
        const dictionaryInfo = await dictionaryDetail.json();

        // Get all keywords

        const keywordsDetail = await getData(`https://localhost:7018/keywords`)
        const keywordsInfo = await keywordsDetail.json();


    async function functionality() {

        userSub.forEach(element => {

            // This condition checks if the user is not on Trial
            if (element.isTrialSub === false) {
                    checks()
            }

            // This Condition will execute when user will be onTrial 

            else {
                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false) {
                        checks()
                    }
                })
            }
        })
    }


    function logout() {
        console.log("I'm in the Logout Function")
        localStorage.setItem("JWT", null)
        location.assign('/Home.html')
    }

    async function checks() {
        document.getElementById("message").innerHTML = "";
        rulesInfo.forEach(rule => {
            detectorsInfo.forEach(detector => {
                if (rule.id === detector.rulesid) {
                    dictionaryInfo.forEach(dictionary => {
                        if (detector.id === dictionary.detectorsid) {
                            let score = 0;
                            keywordsInfo.forEach(keyword => {
                                if (dictionary.id === keyword.dictionaryid) {
                                    Word.run((context) => {
                                        // Queue a command to get the current selection and then
                                        // create a proxy range object with the results.
                                        let range = context.document.body;



                                        // This variable will keep the search results for the longest word.
                                        // Queue a command to load the range selection result.
                                        context.load(range, 'text');

                                        // Synchronize the document state by executing the queued commands
                                        // and return a promise to indicate task completion.
                                        return context.sync()
                                            .then(async function () {
                                                // Get the longest word from the selection.
                                                function countOccurences(string, word) {
                                                    return string.split(word).length - 1;
                                                }
                                                const string = range.text.toLowerCase();
                                                const word = keyword.name
                                                const count = countOccurences(string, word.toLowerCase());  // will give the total number of counts of a word which occurs in Document


                                                score += (keyword.weight * count);

                                                if (detector.threshold > score) {
                                                    await Word.run(async (context) => {

                                                        document.getElementById("message").innerHTML += `<div class="success-msg">
                                                    <i class="fa fa-check"></i>
                                                    Rule :"${rule.name}" threshold is not breached for keyword "${keyword.name}"
                                                    </div>`

                                                        // Queue a command to search the document and ignore punctuation.
                                                        const searchResults = context.document.body.search(keyword.name, { ignorePunct: true });

                                                        // Queue a command to load the font property values.
                                                        searchResults.load('font');

                                                        // Synchronize the document state.
                                                        await context.sync();

                                                        // Queue a set of commands to change the font for each found item.
                                                        for (let i = 0; i < searchResults.items.length; i++) {
                                                            searchResults.items[i].font.color = 'black';
                                                            searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                            searchResults.items[i].font.bold = false;
                                                        }


                                                        // Synchronize the document state.
                                                        await context.sync();
                                                    });
                                                } else {
                                                    await Word.run(async (context) => {

                                                        document.getElementById("message").innerHTML += `<div class="error-msg">
                                                    <i class="fa fa-times-circle"></i>
                                                    Rule :"${rule.name}" threshold is breached for keyword "${keyword.name}"
                                                    </div>`

                                                        // Queue a command to search the document and ignore punctuation.
                                                        const searchResults = context.document.body.search(keyword.name, { ignorePunct: true });

                                                        // Queue a command to load the font property values.
                                                        searchResults.load('font');

                                                        // Synchronize the document state.
                                                        await context.sync();

                                                        // Queue a set of commands to change the font for each found item.
                                                        for (let i = 0; i < searchResults.items.length; i++) {
                                                            searchResults.items[i].font.color = 'purple';
                                                            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                            searchResults.items[i].font.bold = true;
                                                        }


                                                        // Synchronize the document state.
                                                        await context.sync();
                                                    });
                                                }
                                                // Queue a search command.

                                                // Queue a commmand to load the font property of the results.
                                            })
                                            .then(context.sync)
                                    })
                                        .catch(errorHandler);
                                }
                            })
                        }
                    })

                }
            })
        })
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
}
    ());
