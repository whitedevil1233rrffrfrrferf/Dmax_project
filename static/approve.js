document.getElementById('select-all').addEventListener('click', function() {
    let checkboxes = document.querySelectorAll('.user-checkbox');
    checkboxes.forEach(checkbox => checkbox.checked = this.checked);
});
document.getElementById('approve-selected').addEventListener('click', function() {
    alert("hh")
    let selectedUsers = [];
    document.querySelectorAll('.user-checkbox:checked').forEach(checkbox => {
        selectedUsers.push(checkbox.value);
    });

    if (selectedUsers.length === 0) {
        alert("No users selected!");
        return;
    }

    fetch(approveUrl, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify({ emp_ids: selectedUsers })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            alert("Users approved successfully!");
            location.reload(); // Refresh the page to update the table
        } else {
            alert("Error approving users.");
        }
    })
    .catch(error => console.error("Error:", error));
});

