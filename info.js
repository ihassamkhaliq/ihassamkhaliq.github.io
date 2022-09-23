
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
            $('#goBack').click(goback);

        });
    };
    // The initialize function must be run each time a new page is loaded.

    function goback(){
        location.assign('/users.html')
    }
 

    // Getting User Data

    const token = localStorage.getItem("JWT")
    const tenantId = localStorage.getItem("tenantId")

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
    const  usertenant = await tenantDetail.json();
    document.getElementById("tenantRows").innerHTML += `<tr>
    <td  class="active-row">${usertenant.pocName}</td>
    <td>${usertenant.companyName}</td>
</tr>`


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
    
            getInfo()
            function getInfo() {
    
            userSub.forEach(subscription => {
    
                // This condition checks if the user is not on Trial
                if (subscription.isTrialSub === false) {
                    document.getElementById("subscriptionRows").innerHTML += `<tr>
                    <td  class="active-row">${subscription.name}</td>
                    <td>${subscription.subOwner}</td>
                    <td> On Trial </td>
                </tr>`
                        checks()
                }
    
                // This Condition will execute when user will be onTrial 
    
                else {
                   
                    document.getElementById("subscriptionRows").innerHTML += `<tr>
                    <td  class="active-row">${subscription.name}</td>
                    <td>${subscription.subOwner}</td>
                    <td> On Trial </td>
                </tr>`
                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false) {
                        detectorsInfo.forEach(detector => {
                            if (rule.id === detector.rulesid) {
                                dictionaryInfo.forEach(dictionary => {
                                    if (detector.id === dictionary.detectorsid) {
                                        keywordsInfo.forEach(keyword => {
                                            if (dictionary.id === keyword.dictionaryid) {
                                                document.getElementById("rulesRow").innerHTML += `<tr>
                                                <td  class="active-row">${rule.name}</td>
                                            </tr>`
                                            document.getElementById("detectorsRow").innerHTML += `                <tr>
                                            <td  class="active-row">${detector.name}</td>
                                            <td  class="active-row">${detector.threshold}</td>
                                        </tr>`
                                        document.getElementById("dictionary").innerHTML += `<tr>
                                        <td  class="active-row">${dictionary.name}</td>
                                    </tr>`
                                        document.getElementById("keywordsRow").innerHTML += `<tr>
                                        <td  class="active-row">${dictionary.name}</td>
                                        <td  class="active-row">${dictionary.weight}</td>
                                    </tr>`
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
            rulesInfo.forEach(rule => {
                detectorsInfo.forEach(detector => {
                    if (rule.id === detector.rulesid) {
                        dictionaryInfo.forEach(dictionary => {
                            if (detector.id === dictionary.detectorsid) {
                                keywordsInfo.forEach(keyword => {
                                    if (dictionary.id === keyword.dictionaryid) {
                                        document.getElementById("rulesRow").innerHTML += `<tr>
                                        <td  class="active-row">${rule.name}</td>
                                    </tr>`
                                    document.getElementById("detectorsRow").innerHTML += ` <tr>
                                    <td  class="active-row">${detector.name}</td>
                                    <td  class="active-row">${detector.threshold}</td>
                                </tr>`
                                document.getElementById("dictionary").innerHTML += `<tr>
                                <td  class="active-row">${dictionary.name}</td>
                            </tr>`
                                document.getElementById("keywordsRow").innerHTML += `<tr>
                                <td  class="active-row">${dictionary.name}</td>
                                <td  class="active-row">${dictionary.weight}</td>
                            </tr>`
                                    }
                                })
                            }
                        })
    
                    }
                })
            })
        }
   

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    } } 
    ());
