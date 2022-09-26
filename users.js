﻿
(async function () {
    var messageBanner;
    Office.initialize = function (reason) {
        $(document).ready(function () {
            Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
            Office.context.document.settings.saveAsync();
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
            $('#info').click(info);

        });
    };
    // The initialize function must be run each time a new page is loaded.

    function info(){
        location.assign('/info.html')
    }

    setInterval(function () {
        functionality() 
    }, 5000);
 

    // Getting User Data

    const token = localStorage.getItem("JWT")
    const tenantId = localStorage.getItem("tenantId")

    if(token === "undefined"){
        location.assign("/Home.html")
    }else if(token === null){
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


   
        const subscriptionDetail = await getData(`https://localhost:7018/tenant/${tenantId}/subscription/all`)
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
                    document.getElementById("message").innerHTML = "";
                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false) {
                        document.getElementById("message").innerHTML += `<div id="rules">
                                                        <i id="icon"></i>
                                                        Rule :"${rule.name}" threshold is not breached for keyword "<b id="keywords"></b>"
                                                        </div>`
                        detectorsInfo.forEach(detector => {
                            if (rule.id === detector.rulesid) {
                                dictionaryInfo.forEach(dictionary => {
                                    if (detector.id === dictionary.detectorsid) {
                                        let score = 0;
                                        keywordsInfo.forEach((keyword, occurence) => {
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
                                                                    console.log("I'm at level 11")
                                                                    let element = document.getElementById("rules");
                                                                    console.log("I'm at level 12")
                                                                    element.classList.add("success-msg");
                                                                    console.log("I'm at level 13")
                                                                    let ruleIcon = document.getElementById("icon");
                                                                    console.log("I'm at level 14")
                                                                    ruleIcon.classList.add("fa fa-check")
                                                                    document.getElementById("keywords").innerHTML += `${occurence}`
                                                                    
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
                                                                    console.log("I'm at level 17")
                                                                    let element = document.getElementById("rules");
                                                                    console.log("I'm at level 18")
                                                                    element.classList.add("error-msg");
                                                                    console.log("I'm at level 19")
                                                                    let ruleIcon = document.getElementById("icon");
                                                                    console.log("I'm at level 20")
                                                                    ruleIcon.classList.add("fa fa-times-circle")
                                                                    document.getElementById("keywords").innerHTML += `${occurence}`

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
                    }
                })
                }
            })
        }
    
        async function checks() {
            document.getElementById("message").innerHTML = "";
            rulesInfo.forEach(rule => {
                `<div id="rules">
                <i id="icon"></i>
                Rule :"${rule.name}" threshold is not breached for keyword "<b id="keywords"></b>"
                </div>`
                detectorsInfo.forEach(detector => {
                    if (rule.id === detector.rulesid) {
                        dictionaryInfo.forEach(dictionary => {
                            if (detector.id === dictionary.detectorsid) {
                                let score = 0;
                                keywordsInfo.forEach((keyword,occurence) => {
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
                                                            console.log("I'm here at level 1")
                                                            let element = document.getElementById("rules");
                                                            console.log("I'm here at level 2")
                                                            // element.classList.add("success-msg");
                                                            console.log("I'm here at level 3")
                                                            let ruleIcon = document.getElementById("icon");
                                                            console.log("I'm here at level 4")
                                                            ruleIcon.classList.add("fa fa-check")
                                                            document.getElementById("keywords").innerHTML += `${occurence}`
    
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
                                                            console.log("I'm here at level 5")
                                                            let element = document.getElementById("rules");
                                                            console.log("I'm here at level 6")
                                                            element.classList.add("error-msg");
                                                            console.log("I'm here at level 7")
                                                            let ruleIcon = document.getElementById("icon");
                                                            console.log("I'm here at level 8")
                                                            ruleIcon.classList.add("fa fa-times-circle")
                                                                    document.getElementById("keywords").innerHTML += `${occurence}`
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
   

    // Get the Subscription of User



    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    function logout() {
        console.log("I'm in the Logout Function")
        localStorage.setItem("JWT", null)
        location.assign('/Home.html')
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    } } 
    ());
