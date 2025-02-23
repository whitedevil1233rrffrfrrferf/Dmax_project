function handleFilters() {
    const month=document.getElementById("month").value
    const year=document.getElementById("year").value
    const project=document.getElementById("project").value
    let urlParams = new URLSearchParams(window.location.search); 
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
    if (project) {
        urlParams.set('project', project);
    } else {
        // If no date is selected, remove it from the URL parameters
        urlParams.delete('project');
    }
    window.location.href = `${homeUrl}?${urlParams.toString()}`;
    
}