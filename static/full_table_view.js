function handleFilters() {
    const date=document.getElementById("date").value
    alert(date)
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