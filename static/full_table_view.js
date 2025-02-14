function handleFilters() {
    const date=document.getElementById("date").value
    
    const month=document.getElementById("month").value
    const year=document.getElementById("year").value
    let urlParams = new URLSearchParams(window.location.search); 
    if (date) {
        urlParams.set('date', date);
    } else {
        // If no date is selected, remove it from the URL parameters
        urlParams.delete('date');
    }
    if (month) {
        urlParams.set('month', month);
    } else {
        // If no date is selected, remove it from the URL parameters
        urlParams.delete('month');
    }
    if (year) {
        urlParams.set('year', year);
    } else {
        // If no date is selected, remove it from the URL parameters
        urlParams.delete('year');
    }
    window.location.href = `${homeUrl}?${urlParams.toString()}`;
    
}
document.getElementById("selectAll").addEventListener("change", function() {
    let checkboxes = document.querySelectorAll(".rowCheckbox");
    checkboxes.forEach(checkbox => {
        checkbox.checked = this.checked;
    });
});
document.getElementById("deleteSelected").addEventListener("click", function() {
    let selectedIds = Array.from(document.querySelectorAll(".rowCheckbox:checked"))
                           .map(checkbox => checkbox.value);

    if (selectedIds.length === 0) {
        alert("No rows selected!");
        return;
    }
    let confirmDelete = confirm("Are you sure you want to delete the selected entries?");
    if (!confirmDelete) {
        return;  // Stop execution if user cancels
    }
    fetch(deleteUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids: selectedIds })
    })
    .then(response => {
        if (response.ok) {
            window.location.reload();  // Refresh the page after successful deletion
        } else {
            alert("Error deleting entries.");
        }
    }).catch(error => console.error("Error:", error));
});