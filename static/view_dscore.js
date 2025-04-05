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

        // Step 1: Initial confirmation
        const confirmSubmit = confirm("Are you sure you want to approve the Dmax score?");
        if (!confirmSubmit) return; // Stop if user cancels

        // Step 2: Check Operational Excellence
        fetch(checkOperationalExcellenceUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, month, year })
        })
        .then(response => response.json())
        .then(data => {
            if (data.confirm_needed) { 
                // If OpEx check fails, ask again
                const confirmProceed = confirm(data.message);
                if (!confirmProceed) return;

                // Step 3: Send final approval request
                fetch(approveUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email, month, year })
                })
                .then(response => response.json())
                .then(finalData => {
                    if (finalData.success) {
                        location.reload();  // ✅ Refresh on success
                    } else {
                        alert(finalData.message);  // Show backend error
                    }
                })
                .catch(error => console.error("Final approval error:", error));
            } else {
                // No OpEx issue, directly approve
                fetch(approveUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email, month, year })
                })
                .then(response => response.json())
                .then(finalData => {
                    if (finalData.success) {
                        location.reload();  // ✅ Refresh on success
                    } else {
                        alert(finalData.message);
                    }
                })
                .catch(error => console.error("Final approval error:", error));
            }
        })
        .catch(error => console.error("Operational Excellence check error:", error));
    });
});