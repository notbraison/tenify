const fs = require('fs');

// Read the JSON file
fs.readFile('tenants.json', 'utf8', (err, data) => {
    if (err) {
        console.error('Error reading file:', err);
        return;
    }

    // Parse JSON data
    let tenants = JSON.parse(data);

    // Update values
    function updatePhone(n,p){
    tenants.forEach(tenant => {
        if (tenant.houseNumber === n) {
            tenant.phoneNumber = p; // Update phone number for tenant with id n
        }
    });    
    }

    function updateTenantName(n,p){
    tenants.forEach(tenant => {
        if (tenant.houseNumber === n) {
            tenant.name = p; // Update name for tenant with given house number
        }
    });    
    }
    

    // Write updated data back to the file
    fs.writeFile('tenants.json', JSON.stringify(tenants, null, 2), 'utf8', (err) => {
        if (err) {
            console.error('Error writing file:', err);
            return;
        }
        console.log('File updated successfully');
    });
});

