
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

    if (!token) {
        location.assign('./Home.html')
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
    const  usertenant = await tenantDetail.json();
    document.getElementById("tenantRows").innerHTML += `<tr>
    <td  class="active-row">${usertenant.id}</td>
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
    
        userSub.sub.forEach(subscription => {
            document.getElementById("subscriptionRows").innerHTML += `<tr>
                    <td  class="active-row">${subscription.name}</td>
                    <td>${subscription.subOwner}</td>
                    <td> Paid </td>
                    <td>${subscription.subStartDate} </td>
                    <td>${subscription.subEndDate} </td>
                </tr>`
    
                // This condition checks if the user is not on Trial
            if (subscription.isTrialSub === false) {

                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false || rule.subscriptionid === userSub.sub[0].id)
                        document.getElementById("rulesRow").innerHTML += `<tr>
                                        <td  class="active-row">${rule.name}</td>
                                    </tr>`
                        checks(rule.id)
                    })
                }
    
                // This Condition will execute when user will be onTrial 
    
                else {
                rulesInfo.forEach(rule => {
                    if (rule.isPaidRule === false) {
                        document.getElementById("rulesRow").innerHTML += `<tr>
                        <td  class="active-row">${rule.name}</td>
                        </tr>`
                        checks(rule.id)
                    }
                })
                }
            })
        }
    
        async function checks(ruleId) {
           
                detectorsInfo.forEach(detector => {
                    if (ruleId === detector.rulesid) {
                        document.getElementById("detectorsRow").innerHTML += ` <tr>
                                    <td  class="active-row">${detector.name}</td>
                                    <td  class="active-row">${detector.threshold}</td>
                                </tr>`
                        dictionaryInfo.forEach(dictionary => {
                            if (detector.id === dictionary.detectorsid) {
                                document.getElementById("dictionaryRow").innerHTML += `<tr>
                                <td  class="active-row">${dictionary.name}</td>
                            </tr>`
                                keywordsInfo.forEach(keyword => {
                                    if (dictionary.id === keyword.dictionaryid) {
                                document.getElementById("keywordsRow").innerHTML += `<tr>
                                <td  class="active-row">${keyword.name}</td>
                                <td  class="active-row">${keyword.weight}</td>
                            </tr>`
                                    }
                                })
                            }
                        })
    
                    }
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
