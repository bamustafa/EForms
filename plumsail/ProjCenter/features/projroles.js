var onProjRoleRender = async function(){

    debugger;
    let dtName = 'prdt'
    handelTables('div.w-100');
    let result = await getRoles();

    setPageStyle();
    
    _isQM = JSON.parse(localStorage.getItem('_isQM')); 
    _isSus = JSON.parse(localStorage.getItem('_isSus'));
    _isBuilding = JSON.parse(localStorage.getItem('_isBuilding'));
   
    if(!_isQM && !_isSus || _isSus && !_isBuilding) 
      fd.control(dtName).readonly = true;

    if(result.length > 0){
        
        let query= '';
        result.map( (item, index) => {
            let title = item.title
            if(index === 0) 
                query = ` Title eq '${title}' `
            else query += `or Title eq '${title}' `
        });
      
        await filterDataTable_LookupField(dtName, 'ProjRole', query, null, 'EmpName', result);
        fixTableCols(dtName);
    }
}

let getRoles = async function(){
    let ownerFields = 'Owner/Title,Owner/Id'
    let GroupFields = 'Group/Title,Group/Id'
    let cols = `Id,Title,${ownerFields},${GroupFields}`;

    return await _web.lists.getByTitle(Roles).items.select(cols).expand(['Owner','Group']).getAll().then(async (items)=>{
        if(items.length > 0){
            let result = [], members;
            for (const item of items){
                let owner = item['Owner'] !== undefined && item['Owner'] !== null && item['Owner'].length > 0 ? item['Owner']  : null
                if(owner !== null){
                    let User = owner.filter(item=>{ return item.Title === CurrentUser.Title });
                    if(User.length > 0){
                        let group = item['Group'] !== null ? item['Group'].Title  : ''
                   
                       let members = await getSPtGroupMembers(group)
                        
                        result.push({
                            Id: item.Id,
                            title: item.Title,
                            owner: owner,
                            group: group,
                            groupMembers: members
                        });
                    }
                }
            }
            return result
        }
    });
}

