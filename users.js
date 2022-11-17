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

        });
    };
    // The initialize function must be run each time a new page is loaded.


    setInterval(functionality, 5000);
    setTimeout(displayRules, 5000)

    // Getting User Data

    const token = localStorage.getItem("JWT")
    const tenantId = localStorage.getItem("tenantId")

    if (token === "undefined") {
        location.assign("/Home.html")
    } else if (token === null) {
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
    document.getElementById("logo").src = `https://localhost:7018/tenant/image/${tenantId}`


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

    function displayRules() {
        userSub.sub.forEach(subscription => {

            document.getElementById("message").innerHTML = "";

            // This condition checks if the user is not on Trial
            if (subscription.isTrialSub === false) {

                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false || rule.subscriptionid === userSub.sub[0].id) {
                        document.getElementById("message").innerHTML += `<div class="paidRule">
                                                        <i class="paidRuleIcon fa"></i>
                                                        Rule :"${rule.name}"
                                                        </div>`
                    }
                })
            }

            // This Condition will execute when user will be onTrial 

            else {
                rulesInfo.forEach((rule, index) => {
                    if (rule.isPaidRule === false) {
                        document.getElementById("message").innerHTML += `<div id="unpaid${index}" class="unpaidRule">
                                                        <i id="unpaidIcon${index}" class="unpaidRuleIcon fa"></i>
                                                        Rule :"${rule.name}"
                                                        </div>`
                    }
                })
            }
        })
    }





    async function functionality() {

        userSub.sub.forEach(element => {
            let valid = false;
            Office.context.document.settings.remove("Office.AutoShowTaskpaneWithDocument");
            Office.context.document.settings.saveAsync();
            // This condition checks if the user is not on Trial
            if (element.isTrialSub === false) {
                rulesInfo.forEach((rule, index) => {
                    if (rule.isPaidRule === false || rule.subscriptionid === userSub.sub[0].id) {
                        let rules = document.getElementsByClassName("paidRule")[index];
                        rules.classList.remove("error-msg")
                        rules.classList.add("success-msg")
                        let ruleIcon = document.getElementsByClassName("paidRuleIcon")[index];
                        ruleIcon.classList.remove("fa-times-circle");
                        ruleIcon.classList.add("fa-check");
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

                                                            //In Paid user - This check will see if the Keyword is case Sensitive 
                                                            if (keyword.isCaseSensitive) {

                                                                // Here Case Sensitive is true

                                                                // Here code see if the it is a REGEX

                                                                if (keyword.isRegularExpression) {
                                                                    const regex = new RegExp(keyword.name, "g");
                                                                    await Word.run(async (context) => {
                                                                        let body = context.document.body;

                                                                        context.load(body, 'text');

                                                                        return context.sync().then(async () => {

                                                                            let docBody = body.text;
                                                                            let result = docBody.match(regex)

                                                                            if (result) {
                                                                                let count = result.length;
                                                                                score += (keyword.weight * count);

                                                                                // checks if threshold is not braeched

                                                                                if (detector.threshold > score) {
                                                                                    if (valid) {
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
                                                                                    result.forEach(async (regexLiterals) => {
                                                                                        Word.run(async (context) => {
                                                                                            const searchResults = context.document.body.search(regexLiterals, { ignorePunct: true });
                                                                                            searchResults.load('font');
                                                                                            // Synchronize the document state.
                                                                                            return await context.sync().then(() => {


                                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                                    searchResults.items[i].font.color = 'black';
                                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                                    searchResults.items[i].font.bold = false;
                                                                                                }

                                                                                            })

                                                                                            // Queue a set of commands to change the font for each found item.

                                                                                            // Synchronize the document state.
                                                                                        }).catch(() => {
                                                                                            console.log("Exception Here")
                                                                                        })


                                                                                    })
                                                                                }
                                                                                //if threshold breaches this will run
                                                                                else {

                                                                                    let rules = document.getElementsByClassName("paidRule")[index];
                                                                                    rules.classList.remove("success-msg")
                                                                                    rules.classList.add("error-msg")
                                                                                    let ruleIcon = document.getElementsByClassName("paidRuleIcon")[index];
                                                                                    ruleIcon.classList.remove("fa-check");
                                                                                    ruleIcon.classList.add("fa-times-circle");
                                                                                    if (valid) {
                                                                                        valid = true;
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
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
                                                                            }

                                                                        })
                                                                    })
                                                                }
                                                                // if Keyword was not a Regex
                                                                else {
                                                                    await Word.run(async (context) => {

                                                                        // Queue a command to search the document and ignore punctuation.
                                                                        let body = context.document.body;
                                                                        let searchResults = context.document.body.search(keyword.name, { matchCase: true });

                                                                        searchResults.load('font');
                                                                        context.load(body, 'text');


                                                                        return context.sync().then(async () => {

                                                                            const count = searchResults.items.length  // will give the total number of counts of a word which occurs in Document
                                                                            score += (keyword.weight * count);
                                                                            if (detector.threshold > score) {
                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }
                                                                                // Queue a set of commands to change the font for each found item.

                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'black';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                    searchResults.items[i].font.bold = false;
                                                                                }
                                                                            } else {

                                                                                let rules = document.getElementsByClassName("paidRule")[index];
                                                                                rules.classList.remove("success-msg")
                                                                                rules.classList.add("error-msg")
                                                                                let ruleIcon = document.getElementsByClassName("paidRuleIcon")[index];
                                                                                ruleIcon.classList.remove("fa-check");
                                                                                ruleIcon.classList.add("fa-times-circle");

                                                                                valid = true;

                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }
                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'purple';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                                                    searchResults.items[i].font.bold = true;
                                                                                }
                                                                            }
                                                                        }).then(context.sync)

                                                                    })
                                                                    //.catch(errorHandler)
                                                                }
                                                                // This Part will run when Case is Insensitive
                                                            }
                                                            else {
                                                                // Runs if Keyword is a REGEX & Case Insensitive

                                                                if (keyword.isRegularExpression) {
                                                                    const regex = new RegExp(keyword.name, "g");
                                                                    await Word.run(async (context) => {
                                                                        let body = context.document.body;

                                                                        context.load(body, 'text');

                                                                        return context.sync().then(async () => {

                                                                            let docBody = body.text;
                                                                            let result = docBody.match(regex)

                                                                            if (result) {
                                                                                let count = result.length;
                                                                                score += (keyword.weight * count);

                                                                                // Runs if threshold is not breached

                                                                                if (detector.threshold > score) {
                                                                                    if (valid) {
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
                                                                                    result.forEach(async (regexLiterals) => {
                                                                                        Word.run(async (context) => {
                                                                                            const searchResults = context.document.body.search(regexLiterals, { ignorePunct: true });
                                                                                            searchResults.load('font');
                                                                                            // Synchronize the document state.
                                                                                            return await context.sync().then(() => {


                                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                                    searchResults.items[i].font.color = 'black';
                                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                                    searchResults.items[i].font.bold = false;
                                                                                                }

                                                                                            })

                                                                                            // Queue a set of commands to change the font for each found item.

                                                                                            // Synchronize the document state.
                                                                                        }).catch(() => {
                                                                                            console.log("Exception Here1")
                                                                                        })


                                                                                    })
                                                                                }

                                                                                // runs when threshold is breached

                                                                                else {
                                                                                    valid = true;

                                                                                    let rules = document.getElementsByClassName("paidRule")[index];
                                                                                    rules.classList.remove("success-msg")
                                                                                    rules.classList.add("error-msg")
                                                                                    let ruleIcon = document.getElementsByClassName("paidRuleIcon")[index];
                                                                                    ruleIcon.classList.remove("fa-check");
                                                                                    ruleIcon.classList.add("fa-times-circle");
                                                                                    if (valid) {
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        console.log("Auto Open Feature Active")

                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
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
                                                                                                console.log("Exception Here2")
                                                                                            })


                                                                                        })
                                                                                    }
                                                                                }
                                                                            }
                                                                        })
                                                                    })
                                                                }

                                                                // When Keyword is not REGEX & Case Insensitive

                                                                else {

                                                                    await Word.run(async (context) => {

                                                                        // Queue a command to search the document and ignore punctuation.
                                                                        let body = context.document.body;
                                                                        let searchResults = context.document.body.search(keyword.name);

                                                                        searchResults.load('font');
                                                                        context.load(body, 'text');


                                                                        return context.sync().then(async () => {

                                                                            const count = searchResults.items.length  // will give the total number of counts of a word which occurs in Document

                                                                            score += (keyword.weight * count);
                                                                            if (detector.threshold > score) {

                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    console.log("Auto Open Feature Active")
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }
                                                                                // Queue a set of commands to change the font for each found item.

                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'black';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                    searchResults.items[i].font.bold = false;
                                                                                }
                                                                            } else {

                                                                                let rules = document.getElementsByClassName("paidRule")[index];
                                                                                rules.classList.remove("success-msg")
                                                                                rules.classList.add("error-msg")
                                                                                let ruleIcon = document.getElementsByClassName("paidRuleIcon")[index];
                                                                                ruleIcon.classList.remove("fa-check");
                                                                                ruleIcon.classList.add("fa-times-circle");

                                                                                valid = true;

                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    console.log("Auto Open Feature Active")
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }

                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'purple';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                                                    searchResults.items[i].font.bold = true;
                                                                                }
                                                                            }
                                                                        }).then(context.sync)

                                                                    })
                                                                    .catch(errorHandler)
                                                                }
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
                })
            }

            // This Condition will execute when user will be onTrial 
            else {
                rulesInfo.forEach((rule, index = 0) => {
                    if (rule.isPaidRule === false) {
                        let rules = document.getElementById(`unpaid${index}`);
                        rules.classList.remove("error-msg")
                        rules.classList.add("success-msg")
                        let ruleIcon = document.getElementById(`unpaidIcon${index}`);
                        ruleIcon.classList.remove("fa-times-circle");
                        ruleIcon.classList.add("fa-check");
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

                                                            //In Paid user - This check will see if the Keyword is case Sensitive 
                                                            if (keyword.isCaseSensitive) {

                                                                // Here Case Sensitive is true

                                                                // Here code see if the it is a REGEX

                                                                if (keyword.isRegularExpression) {
                                                                    const regex = new RegExp(keyword.name, "g");
                                                                    await Word.run(async (context) => {
                                                                        let body = context.document.body;

                                                                        context.load(body, 'text');

                                                                        return context.sync().then(async () => {

                                                                            let docBody = body.text;
                                                                            let result = docBody.match(regex)

                                                                            if (result) {
                                                                                let count = result.length;
                                                                                score += (keyword.weight * count);

                                                                                // checks if threshold is not braeched

                                                                                if (detector.threshold > score) {
                                                                                    if (valid) {
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
                                                                                    result.forEach(async (regexLiterals) => {
                                                                                        Word.run(async (context) => {
                                                                                            const searchResults = context.document.body.search(regexLiterals, { ignorePunct: true });
                                                                                            searchResults.load('font');
                                                                                            // Synchronize the document state.
                                                                                            return await context.sync().then(() => {


                                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                                    searchResults.items[i].font.color = 'black';
                                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                                    searchResults.items[i].font.bold = false;
                                                                                                }

                                                                                            })

                                                                                            // Queue a set of commands to change the font for each found item.

                                                                                            // Synchronize the document state.
                                                                                        }).catch(() => {
                                                                                            console.log("Exception Here")
                                                                                        })


                                                                                    })
                                                                                }
                                                                                //if threshold breaches this will run
                                                                                else {
                                                                                    let rules = document.getElementById(`unpaid${index}`);
                                                                                    rules.classList.remove("success-msg")
                                                                                    rules.classList.add("error-msg")
                                                                                    let ruleIcon = document.getElementById(`unpaidIcon${index}`);
                                                                                    ruleIcon.classList.remove("fa-check");
                                                                                    ruleIcon.classList.add("fa-times-circle");


                                                                                    if (valid) {
                                                                                        valid = true;
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
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
                                                                            }

                                                                        })
                                                                    })
                                                                }
                                                                // if Keyword was not a Regex
                                                                else {
                                                                    await Word.run(async (context) => {

                                                                        // Queue a command to search the document and ignore punctuation.
                                                                        let body = context.document.body;
                                                                        let searchResults = context.document.body.search(keyword.name, { matchCase: true });

                                                                        searchResults.load('font');
                                                                        context.load(body, 'text');


                                                                        return context.sync().then(async () => {

                                                                            const count = searchResults.items.length  // will give the total number of counts of a word which occurs in Document
                                                                            score += (keyword.weight * count);
                                                                            if (detector.threshold > score) {
                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }
                                                                                // Queue a set of commands to change the font for each found item.

                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'black';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                    searchResults.items[i].font.bold = false;
                                                                                }
                                                                            } else {

                                                                                let rules = document.getElementById(`unpaid${index}`);
                                                                                rules.classList.remove("success-msg")
                                                                                rules.classList.add("error-msg")
                                                                                let ruleIcon = document.getElementById(`unpaidIcon${index}`);
                                                                                ruleIcon.classList.remove("fa-check");
                                                                                ruleIcon.classList.add("fa-times-circle");

                                                                                valid = true;

                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }
                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'purple';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                                                    searchResults.items[i].font.bold = true;
                                                                                }
                                                                            }
                                                                        }).then(context.sync)

                                                                    })
                                                                    //.catch(errorHandler)
                                                                }
                                                                // This Part will run when Case is Insensitive
                                                            }
                                                            else {
                                                                // Runs if Keyword is a REGEX & Case Insensitive

                                                                if (keyword.isRegularExpression) {
                                                                    const regex = new RegExp(keyword.name, "g");
                                                                    await Word.run(async (context) => {
                                                                        let body = context.document.body;

                                                                        context.load(body, 'text');

                                                                        return context.sync().then(async () => {

                                                                            let docBody = body.text;
                                                                            let result = docBody.match(regex)

                                                                            if (result) {
                                                                                let count = result.length;
                                                                                score += (keyword.weight * count);

                                                                                // Runs if threshold is not breached

                                                                                if (detector.threshold > score) {
                                                                                    if (valid) {
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
                                                                                    result.forEach(async (regexLiterals) => {
                                                                                        Word.run(async (context) => {
                                                                                            const searchResults = context.document.body.search(regexLiterals, { ignorePunct: true });
                                                                                            searchResults.load('font');
                                                                                            // Synchronize the document state.
                                                                                            return await context.sync().then(() => {


                                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                                    searchResults.items[i].font.color = 'black';
                                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                                    searchResults.items[i].font.bold = false;
                                                                                                }

                                                                                            })

                                                                                            // Queue a set of commands to change the font for each found item.

                                                                                            // Synchronize the document state.
                                                                                        }).catch(() => {
                                                                                            console.log("Exception Here1")
                                                                                        })


                                                                                    })
                                                                                }

                                                                                // runs when threshold is breached

                                                                                else {
                                                                                    valid = true;

                                                                                    let rules = document.getElementById(`unpaid${index}`);
                                                                                    rules.classList.remove("success-msg")
                                                                                    rules.classList.add("error-msg")
                                                                                    let ruleIcon = document.getElementById(`unpaidIcon${index}`);
                                                                                    ruleIcon.classList.remove("fa-check");
                                                                                    ruleIcon.classList.add("fa-times-circle");

                                                                                    if (valid) {
                                                                                        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                        Office.context.document.settings.saveAsync();
                                                                                    }
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
                                                                                                console.log("Exception Here2")
                                                                                            })


                                                                                        })
                                                                                    }
                                                                                }
                                                                            }
                                                                        })
                                                                    })
                                                                }

                                                                // When Keyword is not REGEX & Case Insensitive

                                                                else {

                                                                    await Word.run(async (context) => {

                                                                        // Queue a command to search the document and ignore punctuation.
                                                                        let body = context.document.body;
                                                                        let searchResults = context.document.body.search(keyword.name);

                                                                        searchResults.load('font');
                                                                        context.load(body, 'text');


                                                                        return context.sync().then(async () => {

                                                                            const count = searchResults.items.length  // will give the total number of counts of a word which occurs in Document

                                                                            score += (keyword.weight * count);
                                                                            if (detector.threshold > score) {

                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }
                                                                                // Queue a set of commands to change the font for each found item.

                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'black';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFFFF'; //white
                                                                                    searchResults.items[i].font.bold = false;
                                                                                }
                                                                            } else {

                                                                                let rules = document.getElementById(`unpaid${index}`);
                                                                                rules.classList.remove("success-msg")
                                                                                rules.classList.add("error-msg")
                                                                                let ruleIcon = document.getElementById(`unpaidIcon${index}`);
                                                                                ruleIcon.classList.remove("fa-check");
                                                                                ruleIcon.classList.add("fa-times-circle");

                                                                                valid = true;

                                                                                if (valid) {
                                                                                    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                                                                                    Office.context.document.settings.saveAsync();
                                                                                }

                                                                                for (let i = 0; i < searchResults.items.length; i++) {
                                                                                    searchResults.items[i].font.color = 'purple';
                                                                                    searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
                                                                                    searchResults.items[i].font.bold = true;
                                                                                }
                                                                            }
                                                                        }).then(context.sync)

                                                                    })
                                                                    //.catch(errorHandler)
                                                                }
                                                            }
                                                        })
                                                        .then(context.sync)
                                                })
                                                //.catch(errorHandler);
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