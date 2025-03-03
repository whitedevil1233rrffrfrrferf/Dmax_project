const monthsDict = {
    "1": "January", "2": "February", "3": "March", "4": "April",
    "5": "May", "6": "June", "7": "July", "8": "August",
    "9": "September", "10": "October", "11": "November", "12": "December"
};
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
document.querySelectorAll('.approve-btn').forEach(button => {
    button.addEventListener('click', function() {
        const email = this.dataset.email;
        const month = this.dataset.month;
        const year = this.dataset.year;
        const monthName = monthsDict[month] || month;
        const confirmApproval = confirm(`Are you sure you want to approve the targets for ${monthName} ${year}?`);
        
        if (!confirmApproval) {
            return;  // ⛔ Stop execution if user cancels
        }
        fetch(approveUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, month, year })
        }).then(response => response.json())
        .then(data => {
            if (data.success) {
                location.reload();  // ✅ Reload to fetch updated approval status
            } else {
                alert("Previous month scores has to be approved first.");
            }
        })
        .catch(error => console.error("Fetch error:", error));
    
    });
});