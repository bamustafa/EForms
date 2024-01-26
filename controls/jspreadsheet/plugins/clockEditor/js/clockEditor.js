var clockEditor = function() {
    // Create the editor once to persist the whole lifecycle of the spreadsheet
    var editor = document.createElement('input');
    editor.classList.add('jss_object');
    editor.type = 'text';
    editor.style.width = '100%';
 
    // Close event
    var closeEvent = null;
 
    // Create instance of the clock picker
    $(editor).clockpicker({ afterDone: function() {
        closeEvent();
    }});
 
    // JSS editor
    var methods = {};
 
    methods.createCell = function(cell, value, x, y, instance, options) {
        cell.innerHTML = value;
    }
 
    methods.updateCell = function(cell, value, x, y, instance, options) {
        if (cell) {
            cell.innerHTML = value;
        }
    }
 
    methods.openEditor = function(cell, value, x, y, instance, options) {
        // Append the clock picker to the input container
        instance.parent.input.appendChild(editor);
        // Make sure input container is not editable
        instance.parent.input.setAttribute('contentEditable', false);
        // Set the current value
        editor.value = value;
        // Focus to open the clock picker
        editor.focus();
        // Make sure JSS object class is set to keep the focus on the spreadsheet
        if (! document.querySelector('.clockpicker-popover').classList.contains('jss_object')) {
            document.querySelector('.clockpicker-popover').classList.add('jss_object');
        }
        // Update close event
        closeEvent = function() {
            instance.closeEditor(cell, true);
        }
    }
 
    methods.closeEditor = function(cell, save, x, y, instance, options) {
        if (save) {
            cell.innerHTML = editor.value;
        } else {
            cell.innerHTML  = '';
        }
        // Return value
        return cell.innerHTML;
    }
 
    return methods;
}();