document.getElementById('employee_name').addEventListener('input', function() {
    var employeeName = this.value;
    
    // Send an Axios POST request
    axios.post('/search', {
        employee_name: employeeName
    })
    .then(function (response) {
        var employees = response.data.employees;
        console.log("Employees data:", employees);
        var listHtml = '<h3>Select an Employee</h3><ul>';
        
        if (employees.length > 0) {
            employees.forEach(function(employee) {
                console.log("Employee details:", employee);
                var employeeDetails = JSON.stringify(employee)  // Convert the employee object to a string for use in `onclick`
                console.log("Employee details stringified:", employeeDetails);
                listHtml += `<li>
                                <a href="#" data-employee='${encodeURIComponent(employeeDetails)}' onclick="selectEmployee(this)">
                                    ${employee.employee_name} (${employee.today_date})
                                </a>
                             </li>`;
            });
        } else {
            listHtml += '<li>No employees found</li>';
        }

        listHtml += '</ul>';
        document.getElementById('employee-list').innerHTML = listHtml;
        document.getElementById('employee-list').style.display = 'block';
        document.getElementById('employee_name').style.width='100%'
        document.querySelector('.form_div').style.height = 'auto';
        document.getElementById('employee-list').style.width='17rem'
        document.querySelector('.flex_first_div').style.gap = '9rem';
    })
    .catch(function (error) {
        console.error('Error:', error);
    });
});

// Function to handle when an employee is clicked
function selectEmployee(element) {
    var employeeDetails = decodeURIComponent(element.getAttribute('data-employee'));
    
    try {
        var employee = JSON.parse(employeeDetails);
        
        console.log(employeeDetails)
        basic_details=[
            'employee_name',
            'employee_id',
            'employee_email',
            'project',
            'designation'
        ]
        basic_details.forEach(function(detail){
            var inputField=document.getElementById(detail)
            if(inputField){
                inputField.value=employee[detail]
            }
        })
        Object.keys(employee).forEach(function(key){
            if (key.endsWith('_target')){
                var inputField=document.getElementById(key);
                if(inputField){
                    inputField.value=employee[key]
                }
            }
        })
        // Loop through each key-value pair and populate the corresponding input field
        // for (var key in employee) {
            
        //     if (employee.hasOwnProperty(key)) {
        //         var inputField = document.getElementById(key);
        //         console.log(inputField)
        //         if (inputField) {
        //             inputField.value = employee[key];
        //         }
        //     }
        // }
    } catch (e) {
        console.error('Failed to parse employee details:', e);
    }

    document.getElementById('employee-list').style.display = 'none';
}
const currentDate=new Date()
const formattedDate = currentDate.toISOString().split('T')[0];
document.getElementById('today_date').value=formattedDate