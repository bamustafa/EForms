var onRender = async function(layout, moduleName, formType) {

    updateTotal();
}

function updateTotal() { 
   
    let a = parseFloat(fd.field('A').value) || 0;  
    let b = parseFloat(fd.field('B').value) || 0;
    
    
    let total1 = a * b;
    

    fd.field('COUNT').value = total1.toFixed(2);  
}
    