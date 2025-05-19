var restUrl;

var GetDictionaries = async function (ProjectNo){
	try {
        const startTime = performance.now();
        restUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx`
        await fetchProjectTeamMethod(ProjectNo);


        await fetchResult();
        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time fetch Project Team Method: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);
	}
}

var fetchProjectTeamMethod = async function (ProjectNo, dynamicsPhase){
	try {
        (async () => {
            try {

                let serviceUrl = `${restUrl}?command=GetProjectMembers&ProjectNo=${ProjectNo}`;
                let ProjectTeam = await fetchProjectTeam('GET', serviceUrl, true);

                const GetProjectTeamNodes = ProjectTeam.getElementsByTagName("Table1");
                let projectTeamByDepartmentAndRole = {};

                for (let i = 0; i < GetProjectTeamNodes.length; i++) {

                    const projectNode = GetProjectTeamNodes[i];

                    // Extracting data from each projectNode
                    let projectCode = projectNode.getElementsByTagName("ProjectCode")[0]?.textContent || '';
                    let projectName = projectNode.getElementsByTagName("ProjectName")[0]?.textContent || '';
                    let employeeId = projectNode.getElementsByTagName("EmployeeId")[0]?.textContent || '';
                    let fullName = projectNode.getElementsByTagName("FullName")[0]?.textContent || '';
                    let loginName = projectNode.getElementsByTagName("LoginName")[0]?.textContent || '';
                    let roleId = projectNode.getElementsByTagName("RoleId")[0]?.textContent || '';
                    let role = projectNode.getElementsByTagName("Role")[0]?.textContent || '';
                    let roleRemarks = projectNode.getElementsByTagName("RoleRemarks")[0]?.textContent || '';
                    let departmentName = projectNode.getElementsByTagName("DepartmentName")[0]?.textContent || '';
                    let sectionDesc = projectNode.getElementsByTagName("SectionDesc")[0]?.textContent || '';
                    let email = projectNode.getElementsByTagName("Email")[0]?.textContent || '';
                    let grade = projectNode.getElementsByTagName("Grade")[0]?.textContent || '';
                    let dptCode = projectNode.getElementsByTagName("DptCode")[0]?.textContent || '';
                    let dptAcr = projectNode.getElementsByTagName("DptAcr")[0]?.textContent || '';
                    let projectStatus = projectNode.getElementsByTagName("ProjectStatus")[0]?.textContent || '';
                    let projectStatusId = projectNode.getElementsByTagName("ProjectStatusId")[0]?.textContent || '';
                    let adArea = projectNode.getElementsByTagName("ADArea")[0]?.textContent || '';
                    let isLead = projectNode.getElementsByTagName("IsLead")[0]?.textContent || '';

                    if(isLead === 'true' || isLead === '1' || isLead === 'on' || isLead === 'yes')
                        isLead = true
                    else isLead = false;

                    // if(departmentName !== 'Area Operations'){

                        let teamMember = {
                            projectCode,
                            projectName,
                            employeeId,
                            fullName,
                            loginName,
                            roleId,
                            role,
                            roleRemarks,
                            departmentName,
                            sectionDesc,
                            email,
                            grade,
                            dptCode,
                            dptAcr,
                            projectStatus,
                            projectStatusId,
                            adArea,
                            isLead
                        };

                        // Initialize department if not present
                        if (!projectTeamByDepartmentAndRole[departmentName]) {
                            projectTeamByDepartmentAndRole[departmentName] = {};
                        }

                        // Initialize role if not present in the department
                        if (!projectTeamByDepartmentAndRole[departmentName][role]) {
                            projectTeamByDepartmentAndRole[departmentName][role] = [];
                        }

                        // Add the team member to the appropriate department and role
                        projectTeamByDepartmentAndRole[departmentName][role].push(teamMember);
                    // }
                }

                //console.log(projectTeamByDepartmentAndRole);

                //let phases = ['Phase 3', '']; //['Phase 3','Phase 4', '']; // for test integration part

                await createPhases(ProjectNo);
                if(dynamicsPhase){
                    let isFound = await ensureItemsExist(projectTeamByDepartmentAndRole, dynamicsPhase, false);

                    if(isFound){
                        let pmMetaInfo = {
                            Title: 'PM',
                            Phases: phase,
                            IsProjectComp: phase === undefined || phase === null || phase === '' ? true : false,
                            IsPMTask: true,
                            MasterIDId: _itemId
                        };

                        await ensureItemsExist(pmMetaInfo, phase, true); // CREATE PM TASK
                    }

                    checkForLabelAndAppend(dynamicsPhase);
                    //await onCompletionReportRender();
                    hidePreloader();
                }

            } catch (error) {
                console.log("Error in getting filtered items:", error);
            }
        })();
	}
	catch (e) {
		console.log(e);
	}
}

var fetchResult = async function (){
	try {
        (async () => {
            try {
                let serviceUrl = `${restUrl}?command=get-mtds`;
                let items = await fetchRequestUrl('GET', serviceUrl, true);
                await setDictionaries(items.MTDs, MTDs);

                let modifiedDate = getToday()
                serviceUrl = `${restUrl}?command=get-firms&ModifiedDate=${modifiedDate}`;
                items = await fetchRequestUrl('GET', serviceUrl, true);
                await setDictionaries(items.Firms, Firms);

                serviceUrl = `${restUrl}?command=get-QM`;
                items = await fetchRequestUrl('GET', serviceUrl, true);
                await addUsers(items, 'QM');

            } catch (error){
                console.log("Error in- getting items:", error);
            }
        })();
	}
	catch (e) {
		console.log(e);
	}
}

//#region General Functions
async function ensureItemsExist(projectTeamByDepartmentAndRole, phase, isPM) {
    let itemValue = '';
    try {

        if(phase){
            if(!phase.startsWith("Phase"))
                phase = `Phase ${phase}`
        }

        let ListNames = isPM ? [CompReports] : [RevDepartments, CompReports];
        let selectedColumns = ['Title', 'MasterID/Id', 'GLMain/Id', 'GLMain/Title', 'GL/Id', 'GL/Title', 'TeamMembers/Id', 'TeamMembers/Title', 'ID'];
        let expandColumns = ['GLMain', 'GL', 'TeamMembers', 'MasterID'];
        let filterColumns = `MasterID/Id eq ${_itemId}`;

        ListNames.forEach(async listName => {
            // Fetch existing items with necessary fields expanded

            if(listName === 'Completion Reports'){
                filterColumns += phase === '' ? ` and (Phases eq null or Phases eq '${phase}') ` : ` and Phases eq '${phase}' `;
                selectedColumns = ['Title', 'Phases', 'MasterID/Id', 'ID'];
                expandColumns = ['MasterID'];

                if(isPM)
                    filterColumns += ` and Title eq 'PM'`
            }

            const existingItems = await pnp.sp.web.lists.getByTitle(listName).items
            .select(selectedColumns)
            .expand(expandColumns)
            .filter(filterColumns)
            .getAll();

            if(!isPM && listName === 'Completion Reports' && existingItems.length > 0)
                return true

            let itemsToUpdate = [];

            if(isPM){
                if(existingItems.length === 0)
                   itemsToUpdate.push(projectTeamByDepartmentAndRole);
                else return;
            }

            else{

                let existingItemsMap;
                if (listName === 'Completion Reports') {
                    existingItemsMap = new Map(
                        existingItems.map(item =>
                            [`${item.Title.trim()}|${item.Phases === undefined || item.Phases === null ? '' : item.Phases}`, item])
                    );
                } else {
                    existingItemsMap = new Map(
                        existingItems.map(item => [item.Title.trim(), item])
                    );
                }

                const GLsgroupId = await getGroupIdByName('GLs');
                const PMgroupId = await getGroupIdByName('PM');
                const PDgroupId = await getGroupIdByName('PD');
                const ADgroupId = await getGroupIdByName('AD');

                for (const departmentName in projectTeamByDepartmentAndRole) {

                    let isLead = false;
                    let glMainUsers = [];
                    let glUsers = [];
                    let teamMembers = [];

                    if(listName !== 'Completion Reports'){
                        const roles = projectTeamByDepartmentAndRole[departmentName];

                        const userPromises = [];

                        for (const role in roles) {
                            const members = roles[role];

                            members.forEach(member => {
                                if(!isLead && member.isLead)
                                isLead = true;
                                if (role === 'GL(PMIS)') {
                                    userPromises.push(
                                        (async () => {
                                            try {
                                                const user = await _web.siteUsers.getByEmail(member.email).get(); // Await the user retrieval
                                                glMainUsers.push(user.Id);

                                                if (GLsgroupId) {
                                                    // Check if the user already exists in the group
                                                    const groupUsers = await _web.siteGroups.getById(GLsgroupId).users.get();
                                                    const userExists = groupUsers.some(groupUser => groupUser.LoginName === user.LoginName);

                                                    if (!userExists) {
                                                        await _web.siteGroups.getById(GLsgroupId).users.add(user.LoginName); // Add the user to the group if not already in it
                                                    }
                                                }
                                            } catch (error) {
                                                //console.error(`Error adding user ${member.email}:`, error); // Log error if any
                                            }
                                        })()
                                    );

                                } else if (role === 'GL') {
                                    userPromises.push(
                                        (async () => {
                                            try {
                                                const user = await _web.siteUsers.getByEmail(member.email).get(); // Await the user retrieval
                                                glUsers.push(user.Id); // Push user ID to the array
                                            } catch (error) {
                                                //console.error(`Error adding user ${member.email}:`, error); // Log error if any
                                            }
                                        })()
                                    );
                                }
                                else if (role.includes("PD") || role.includes("PM") || role.includes("AD")){
                                    (async () => {
                                        try {
                                            const user = await _web.siteUsers.getByEmail(member.email).get(); // Get the user once
                                            let groupId;

                                            if (role.includes("PD")) {
                                                groupId = PDgroupId;
                                            } else if (role.includes("PM")) {
                                                groupId = PMgroupId;
                                            } else if (role.includes("AD")) {
                                                groupId = ADgroupId;
                                            }

                                            if (groupId) {
                                                const groupUsers = await _web.siteGroups.getById(groupId).users.get();
                                                const userExists = groupUsers.some(groupUser => groupUser.LoginName === user.LoginName);

                                                if (!userExists) {
                                                    await _web.siteGroups.getById(groupId).users.add(user.LoginName); // Add the user to the group if not already in it
                                                }
                                            }
                                        } catch (error) {
                                            //console.error(`Error adding user ${member.email}:`, error); // Log error if any
                                        }
                                    })()
                                }
                                else {
                                    userPromises.push(
                                        (async () => {
                                            try {
                                                const user = await _web.siteUsers.getByEmail(member.email).get(); // Await the user retrieval
                                                teamMembers.push(user.Id); // Push user ID to the array
                                            } catch (error) {
                                                //console.error(`Error adding user ${member.email}:`, error); // Log error if any
                                            }
                                        })() // Immediately invoke the async function
                                    );
                                }
                            });
                        }
                        await Promise.all(userPromises);
                    }

                    let itemData = {
                        Title: departmentName,
                        GLMainId: { results: glMainUsers || []},
                        GLId: { results: glUsers || []},
                        TeamMembersId: { results: teamMembers || []},
                        MasterIDId: _itemId,
                        ISLead: isLead
                    };

                    if(listName === 'Completion Reports'){
                        delete itemData.GLMainId;
                        delete itemData.GLId;
                        delete itemData.TeamMembersId;
                        delete itemData.ISLead;

                        itemData.Phases = phase;
                        itemData.IsProjectComp = phase === undefined || phase === null || phase === '' ? true : false;

                        if (existingItemsMap.has(`${departmentName}|${phase}`)){

                            // let existingItem = existingItemsMap.get(`${departmentName}|${phase}`);
                            // itemsToUpdate.push({
                            //     Id: existingItem.Id,
                            //     Status: existingItem.Status,
                            //     ExistGLMain: existingItem.GLMain,
                            //     ...itemData
                            // });

                            // if(listName === 'Completion Reports'){
                            //     delete itemsToUpdate.ExistGLMain;
                            // }

                        } else {
                            // Add new item
                            itemsToUpdate.push(itemData);
                        }
                    }
                    else{
                        if (existingItemsMap.has(departmentName)) {

                            let existingItem = existingItemsMap.get(departmentName.trim());
                            itemsToUpdate.push({
                                Id: existingItem.Id,
                                Status: existingItem.Status,
                                ExistGLMain: existingItem.GLMain,
                                ...itemData
                            });

                        } else {
                            // Add new item
                            itemsToUpdate.push(itemData);
                        }
                    }
                }
            }

            if (itemsToUpdate.length > 0) {
                const batch = pnp.sp.web.createBatch();

                itemsToUpdate.forEach(item => {

                    if(listName !== 'Completion Reports'){
                        let existGLMain = item.ExistGLMain ? item.ExistGLMain.length : 0;
                        let ExistItemStatus = item.Status;
                        let currentGLMainCount = item.GLMainId.results.length;

                        if (currentGLMainCount === 0 && (ExistItemStatus === 'In Progress' || ExistItemStatus === undefined))
                            item.Status = 'Approve';
                        else if (currentGLMainCount > 0 && ExistItemStatus === 'Approve' && existGLMain === 0)
                            item.Status = 'In Progress';

                        delete item.ExistGLMain;
                    }

                    if (item.Id) {
                        // Update existing item
                        pnp.sp.web.lists.getByTitle(listName).items.getById(item.Id).inBatch(batch).update(item)
                            //.then(() => console.log(`Item for Department "${item.Title}" updated.`));
                    } else {
                        // Add new item
                        pnp.sp.web.lists.getByTitle(listName).items.inBatch(batch).add(item)
                            //.then(() => console.log(`Item for Department "${item.Title}" created.`));
                    }
                });

                await batch.execute();
            } else {
                console.log("No items to update.");
            }

        });

    } catch (error) {
        console.log(`Item: ${JSON.stringify(itemValue)} Error ensuring items exist:`, error);
    }
}

async function setDictionaries(items, listname){

    const batch = _web.createBatch();

    const cols = getMetaInfo(listname).cols;
    for(const item of items){
        let filterQuery = getMetaInfo(listname, item).filterQuery;
        if(filterQuery === '' || filterQuery === undefined)
            continue;

        const result = await _web.lists.getByTitle(listname).items.select(cols).filter(filterQuery).get();
        if(result.length === 0){
            let itemResult = getMetaInfo(listname, item, true).itemResult;
            // _web.lists.getByTitle(listname).items.inBatch(batch).add(itemResult)
            //   .then(() => console.log(`Item for ${listname} "${filterQuery}" created.`));

            await _web.lists.getByTitle(listname).items.add(itemResult);
        }
        else{
            let listItem = result[0];
            let obj = getMetaInfo(listname, item, true, listItem)

            if(obj.doUpdate)
              await pnp.sp.web.lists.getByTitle(listname).items.getById(listItem.Id).update(obj.itemResult);
        }
    }
    if(batch._index > -1)
      await batch.execute();
}

function getMetaInfo(listname, item, isResult, result){
  let cols = 'Id,Title';
  let filterQuery = '';
  let itemResult = {};
  let desc;
  let doUpdate = false;

  switch(listname){
    case 'MTDs':
      cols = 'Id,Title,Code,FullDesc,Superseded'

      if(item !== undefined){
        desc = item.Description.trim();
        if(desc === '' || desc === undefined || desc === null){
            filterQuery = "";
            return {
                cols: cols,
                filterQuery: filterQuery,
                itemResult: itemResult
              }
        }
        filterQuery = `Title eq '${desc}'`
      }

      if(isResult){
          let code = item.MTDCode.trim()
          itemResult['Title'] = desc;
          itemResult['Code'] = code;
          itemResult['FullDesc'] = `${code} : ${desc}`;
          itemResult['Superseded'] = item.Superseded;
      }
      if(result !== undefined){
        if(item.Superseded !== result.Superseded)
            doUpdate = true
      }
      break;
    case 'Firms':
      cols = 'Id,Title,FullDesc,IsActive,FirmId,Relationships,FirmType'

      if(item !== undefined){
        desc = item.FirmName.trim();
        if(desc === '' || desc === undefined || desc === null){
            filterQuery = "";
            return {
                cols: cols,
                filterQuery: filterQuery,
                itemResult: itemResult
              }
        }
        filterQuery = `FirmId eq '${item.FirmId}'`
      }

      if(isResult){
        let isActive = item.Status === 'Active' ? true: false;
        itemResult['Title'] = desc.substring(0, 255);
        itemResult['FullDesc'] = desc;
        itemResult['IsActive'] = isActive;
        itemResult['FirmId'] = item.FirmId;
        itemResult['Relationships'] = item.Relationships;
        itemResult['FirmType'] = item.Type;
      }

      if(result !== undefined){
        let isActive = item.Inactive ? false: true;
        if(isActive !== result.IsActive)
            doUpdate = true
      }
      break;
  }
  return {
    cols: cols,
    filterQuery: filterQuery,
    itemResult: itemResult,
    doUpdate: doUpdate
  }
}

function getToday(){
   let today = new Date();
   let pastDate = new Date(today);
   pastDate.setDate(today.getDate() - 30);

  let yyyy = pastDate.getFullYear();
  let mm = String(pastDate.getMonth() + 1).padStart(2, '0'); // Months are zero-based
  let dd = String(pastDate.getDate()).padStart(2, '0');

  return `${yyyy}-${mm}-${dd}`;
}

async function addUsers(items, groupName){
  let groupmembers = await getSharePointGroupMembers(groupName);

  const difference = items.filter(item1 => !groupmembers.some(item2 => item2.Email === item1.Email));

  difference.forEach(async (item) => {
    let email = item.Email;
    try {
        //_web.siteUsers.add(item.LoginName);
        await _web.siteGroups.getByName(groupName).users.add(item.LoginName);
        //console.log(`User ${email} added successfully.`);
    } catch (error) {
        console.error(`Error adding user ${email}:`, error);
      }
 });

}

async function getSharePointGroupMembers(groupName) {
    const group = await _web.siteGroups.getByName(groupName);
    const members = await group.users.get();
    return members;
}

async function getGroupIdByName(groupName) {
    try {
        const groups = await pnp.sp.web.siteGroups.filter(`Title eq '${groupName}'`).get();
        if (groups.length > 0) {
            return groups[0].Id; // Return the ID of the first matching group
        } else {
            throw new Error("Group not found.");
        }
    } catch (error) {
        console.error("Error getting group ID:", error);
        throw error;
    }
}

async function createPhases(ProjectNo){

    if($('div.customPhases').length > 0)
        return;

    let phases = $("label:contains('Phases')").next();
    let phaseBlankCell = phases.find('.d-sm-block').last(); //.eq(1);
    let serviceUrl = `${restUrl}?command=project_phases&ProjectId=${ProjectNo}`;
    let Array = await fetchRequestUrl('GET', serviceUrl, true);

    let createBtn = false;
    if(Array.length > 0){
      let imgUrlLegend = `${_webUrl}${_layout}/Images/legend.png`

      let radioButton = '';
      Array[0].Phases.forEach((phase, index) => {

        let phaseName = `Phase ${phase}`
        let isPhaseFound = phases.find(`label:contains('${phaseName}')`);

        if(isPhaseFound.length > 0){
            return;
        }
        else createBtn = true;

         radioButton += `
            <div id= 'Phase${phase}' class="customPhases" style="padding-left: 8px">
                <input type="radio" id="radio${index}" name="phases" value="${phaseName}">
                <label for="radio${index}" style="padding: 4px;">${phaseName}</label>
            </div>`;
      });

      if(!createBtn){
        phases.find('div.d-none').first().remove()
        return
      }

      let htmlCtlr = `<fieldset id='legId' style="padding:1px !important; border: 1px solid #666 !important;">
                          <legend style="font-family: Calibri; font-size: 12pt; font-weight: bold;padding: 1px 10px !important;float:none;width:auto;">
                             <img src="${imgUrlLegend}" style="width: 16px; vertical-align: middle;">
                              <b id="phaseHeadId">Select phase below to create</b>
                          </legend>`

      phaseBlankCell.append(htmlCtlr) //("<b id='phaseHeadId'>Select phase below to create</b>");

      debugger
      let fieldset = $('#legId');
      fieldset.append(radioButton); // Append the radio button to the target element

        let button = $('<button>')
            .text('Create Phase') // Set the button text
            .addClass('btn btn-primary') // Add classes (e.g., Bootstrap classes)
            .attr('type', 'button') // Set the type attribute
            .css({
                'margin-left': '12px', // Set margin-left
                'margin-top': '11px'   // Set margin-top
              })
            //.attr('disabled', disableBtn)
            .on('click', async function () { // Add an onClick event listener
                let selectedInput = $('input[name="phases"]:checked');
                if(selectedInput.length === 0){
                alert('select phase to create');
                return;
                }

                showPreloader();

                let selectedOption = selectedInput.val();
                let fldPhases = fd.field("Phases").options
                fldPhases.push(selectedOption)

                fd.field("Phases").options = fldPhases;

                await fetchProjectTeamMethod(ProjectNo, selectedOption)

                let inputId = selectedOption.replace(/\s+/g, '')
                debugger;
                setTimeout(() => {
                    $(`#${inputId}`).remove();

                    if($('div.customPhases').length === 0)
                        fieldset.remove();

                  }, 500);
                //$(this).prop('disabled', true);

            });
            fieldset.append(button);
    }
}

function checkForLabelAndAppend(selectedOption) {
    let newOptionLabel = $("label:contains('Phases')").next().find(`label:contains('${selectedOption}')`);
    if (newOptionLabel.length > 1) {
        newOptionLabel.first().append('<span style="font-size: 12px; font-weight: bold; text-align: center; color: #936106cf;"> In Progress</span>');
    }
}
//#endregion

//#region Soap Call Function

function fetchProjectTeam(method, serviceUrl, isAsync) {
    return new Promise((resolve, reject) => {
        var xhr = new XMLHttpRequest();
        xhr.open(method, serviceUrl, isAsync);

        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        const response = xhr.responseText; // Use xhr.responseText to access the response text
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");
                        resolve(xmlDoc); // Resolve the promise with the parsed XML document
                    } catch (err) {
                        reject(new Error(`Failed to parse XML: ${err.message}`));
                    }
                } else {
                    reject(new Error(`Failed to get valid response: ${xhr.statusText}`));
                }
            }
        };

        xhr.setRequestHeader('Content-Type', 'text/xml');
        xhr.send();
    });
}

function fetchRequestUrl(method, serviceUrl, isAsync){
    return new Promise((resolve, reject) => {
        var xhr = new XMLHttpRequest();
        xhr.open(method, serviceUrl, isAsync);

        xhr.onreadystatechange = function(){
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    const response = JSON.parse(xhr.responseText);
                    resolve(response);
                } else reject(new Error(`Failed to get valid response: ${xhr.statusText}`));
            }
        };

        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.send();
    });
}

//#endregion