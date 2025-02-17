var onDesignProcessRender = async function(){

    const startTime = performance.now();
    let fields = {
        Process: fd.field('Process'),
        CADJust: fd.field('CADJust'),
        DRDApplicable: fd.field('DRDApplicable'),
        DRDJustification: fd.field('DRDJustification')
    }

    let process = fields.Process.value;
    process = process !== undefined && process !== null && process !== '' ? process : '';

    let cadJust = fields.CADJust.value;
    cadJust = cadJust !== undefined && cadJust !== null && cadJust !== '' ? cadJust : '';

    let operation = process
    if(process === '' || process === 'Other')
        operation = 'process';
       
     await hideTabFields(fields, operation);

     
    let subCategory = fd.field("SubCategory").value;
    await isDRD(subCategory);

    await onProcessChange(fields);
    await onDRDApplicableChange(fields);

    const endTime = performance.now();
    const elapsedTime = endTime - startTime;
    console.log(`onDesignProcessRender: ${elapsedTime} milliseconds`);
}

const hideTabFields = async function(fields, operation){

    if(operation === 'process' || operation === 'BIM'){
      _HideFields([fields.CADJust], true);
      fields.CADJust.required = false;

      if(operation === 'process')
        fd.container('BIMAccordion').hidden = true;
    }
    else{
        _HideFields([fields.CADJust], false);
        fields.CADJust.required = true;
        fd.container('BIMAccordion').hidden = true;
    }
    
    let drdApplicable = fields.DRDApplicable.value;
    drdApplicable = drdApplicable !== undefined && drdApplicable !== null && drdApplicable !== '' ? drdApplicable : '';

    if(drdApplicable === 'No')
        _HideFields([fields.DRDJustification], false);
    else _HideFields([fields.DRDJustification], true);
    
        
}

const onProcessChange = async function(fields){
    fields.Process.$on('change', async function(){
        let value = this.value;
        let category = fd.field("Category").value.LookupValue;
        
        if(value === 'CADD'){
            fd.container('BIMAccordion').hidden = true;
            if(category === 'Buildings' || category === 'Infrastructure Transportation'){
              $(fields.CADJust.$parent.$el).show();
              fields.CADJust.required = true;
            }
        }
        else if(value === 'BIM'){
          fd.container('BIMAccordion').hidden = false;
          $(fields.CADJust.$parent.$el).hide();
          fields.CADJust.required = false;

          handleSingleTable()
        }
        else await hideTabFields(fields, 'process');
    })
}

const onDRDApplicableChange = async function(fields){
    fields.DRDApplicable.$on('change', async function(){
        let value = this.value;
    
        if(value === 'No'){
            $(fields.DRDJustification.$parent.$el).show();
            fields.DRDJustification.required = true;
        }
        else {
            $(fields.DRDJustification.$parent.$el).hide();
            fields.DRDJustification.required = false;
        }
    })
}