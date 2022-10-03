(async function () {
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
            $('#info').click(info);

        });
    };
    // The initialize function must be run each time a new page is loaded.




    function info() {
        location.assign('/info.html')
    }

    setInterval(function () {
        functionality()
    }, 5000);


    // Getting User Data

    const token = localStorage.getItem("JWT")
    const tenantId = localStorage.getItem("tenantId")

    if (token === "undefined") {
        console.log("I'm here")
        console.log("I'm here")
        location.assign("/Home.html")
    } else if (token === null) {
        console.log("I'm here null")
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



    const tenantDetail = await getData(`https://localhost:7018/tenant/${tenantId}`)
    const usertenant = await tenantDetail.json();
    document.getElementById("companyName").innerHTML = `${usertenant.companyName}`

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

    let autoOpenChecker = 0;

    async function functionality() {

        userSub.forEach(element => {

            document.getElementById("message").innerHTML = "";
            // This condition checks if the user is not on Trial
            if (element.isTrialSub === false) {
                rulesInfo.forEach(rule => {
                    checks(rule.id, rule.name)
                })
            }

            // This Condition will execute when user will be onTrial 
            else {
                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false) {
                        checks(rule.id, rule.name)
                    }
                })
            }
        })
    }

   

    let regex = /\b\w{9}\b/g
    let regex1 = /[a-zA-Z0-9]{2}[0-9]{6,}/g

    async function checks(ruleId, ruleName) {
        let valid = false;
        detectorsInfo.forEach(detector => {
            if (ruleId === detector.rulesid) {
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

                                            const string = range.text.toLowerCase();
                                            const word = keyword.name
                                            //const count = countOccurences(string, word.toLowerCase());  // will give the total number of counts of a word which occurs in Document
                                            let switcher = document.getElementById('toggler').checked

                                            if (switcher) {
                                                await Word.run(async (context) => {

                                                    // Queue a command to search the document and ignore punctuation.
                                                    let body = context.document.body;
                                                    let searchResults = context.document.body.search(keyword.name, { matchCase: true });

                                                    searchResults.load('font');
                                                    context.load(body, 'text');


                                                    return context.sync().then(async () => {

                                                        const count = searchResults.items.length  // will give the total number of counts of a word which occurs in Document
                                                        await regEx(body.text, regex)
                                                        await regEx(body.text, regex1)
                                                        score += (keyword.weight * count);
                                                        if (detector.threshold > score) {

                                                            if (valid) {
                                                                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                console.log("Auto Open Feature Active")
                                                                
                                                                Office.context.document.settings.saveAsync();


                                                            }else{
                                                                Office.context.document.settings.remove("Office.AutoShowTaskpaneWithDocument");
                                                                Office.context.document.settings.saveAsync();
                                                            }
                                                      


                                                            // Queue a set of commands to change the font for each found item.
                                                            document.getElementById("message").innerHTML += `<div class="success-msg">
                                                            <i class="fa fa-check"></i>
                                                            Rule :"${ruleName}" threshold is not breached for keyword "${keyword.name}"
                                                            </div>`

                                                            for (let i = 0; i < searchResults.items.length; i++) {
                                                                searchResults.items[i].font.color = 'black';
                                                                searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                searchResults.items[i].font.bold = false;
                                                            }
                                                        } else {

                                                            valid = true;

                                                            if (valid) {
                                                                console.log("Auto Open Feature Active")
                                                                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                Office.context.document.settings.saveAsync();
                                                            }else{
                                                                Office.context.document.settings.remove("Office.AutoShowTaskpaneWithDocument");
                                                                Office.context.document.settings.saveAsync();
                                                            }
                                                            document.getElementById("message").innerHTML += `<div class="error-msg">
                                                        <i class="fa fa-times-circle"></i>
                                                        Rule :"${ruleName}" threshold is breached for keyword "${keyword.name}"
                                                        </div>`
                                                            for (let i = 0; i < searchResults.items.length; i++) {
                                                                searchResults.items[i].font.color = 'purple';
                                                                searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                                searchResults.items[i].font.bold = true;
                                                            }
                                                        }
                                                    }).then(context.sync)

                                                }).catch(errorHandler)
                                            }
                                            else {
                                                await Word.run(async (context) => {

                                                    // Queue a command to search the document and ignore punctuation.
                                                    let body = context.document.body;
                                                    let searchResults = context.document.body.search(keyword.name);

                                                    searchResults.load('font');
                                                    context.load(body, 'text');


                                                    return context.sync().then(async () => {

                                                        const count = searchResults.items.length  // will give the total number of counts of a word which occurs in Document
                                                        await regEx(body.text, regex)
                                                        await regEx(body.text, regex1)
                                                        score += (keyword.weight * count);
                                                        if (detector.threshold > score) {

                                                            if (valid) {
                                                                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                console.log("Auto Open Feature Active")
                                                                Office.context.document.settings.saveAsync();
                                                            }else{
                                                                Office.context.document.settings.remove("Office.AutoShowTaskpaneWithDocument");
                                                                Office.context.document.settings.saveAsync();
                                                            }


                                                            // Queue a set of commands to change the font for each found item.
                                                            document.getElementById("message").innerHTML += `<div class="success-msg">
                                                            <i class="fa fa-check"></i>
                                                            Rule :"${ruleName}" threshold is not breached for keyword "${keyword.name}"
                                                            </div>`

                                                            for (let i = 0; i < searchResults.items.length; i++) {
                                                                searchResults.items[i].font.color = 'black';
                                                                searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                searchResults.items[i].font.bold = false;
                                                            }
                                                        } else {

                                                            valid = true;

                                                            if (valid) {
                                                                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                console.log("Auto Open Feature Active")
                                                                Office.context.document.settings.saveAsync();
                                                            }else{
                                                                Office.context.document.settings.remove("Office.AutoShowTaskpaneWithDocument");
                                                                Office.context.document.settings.saveAsync();
                                                            }
                                                            

                                                            document.getElementById("message").innerHTML += `<div class="error-msg">
                                                        <i class="fa fa-times-circle"></i>
                                                        Rule :"${ruleName}" threshold is breached for keyword "${keyword.name}"
                                                        </div>`
                                                            for (let i = 0; i < searchResults.items.length; i++) {
                                                                searchResults.items[i].font.color = 'purple';
                                                                searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                                searchResults.items[i].font.bold = true;
                                                            }
                                                        }
                                                    }).then(context.sync)

                                                }).catch(errorHandler)
                                            }
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




    // Get the Subscription of User
    function countOccurences(string, word) {
        return string.split(word).length - 1;
    }

    async function regEx(docBody, regex) {
        let string = docBody;
        let result = string.match(regex)
        if (result) {
            result.forEach(async (regexLiterals) => {
                Word.run(async (context) => {
                    const searchResults = context.document.body.search(regexLiterals, { ignorePunct: true });
                    searchResults.load('font');
                    // Synchronize the document state.
                    return await context.sync().then(() => {


                        for (let i = 0; i < searchResults.items.length; i++) {
                            searchResults.items[i].font.color = 'purple';
                            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                            searchResults.items[i].font.bold = true;
                        }

                    })

                    // Queue a set of commands to change the font for each found item.

                    // Synchronize the document state.
                }).catch(() => {
                    console.log("Exception Here")
                })


            })
        }

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
    }
}
    ());