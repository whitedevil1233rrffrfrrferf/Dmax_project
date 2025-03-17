const labels = chartData.map(item => item.name);
const scores = chartData.map(item => item.score);

const ctx = document.getElementById('myChart');
const earning=document.getElementById('earning')
function isValidData(labels, data) {
  return labels.length > 0 && data.some(value => value !== 0);
}
if (isValidData(labels, scores)) {
  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Dmax Score',
        data: scores,
        backgroundColor: ['rgba(212, 253, 217, 0.8)',  // Light Blue
          'rgba(148, 223, 166, 0.8)',],
        borderColor: ['rgba(77, 168, 218, 1)',  
          'rgba(255, 140, 157, 1)', ],
        borderWidth: 1
      }]
    },
    options: {
      scales: {
        y: {
          beginAtZero: true
        }
      }
    }
  });
}
else {
  ctx.parentElement.style.display = "none"; // Hide the div if no valid data
}

if (isValidData(desig_name, desig_Scores)) {
  new Chart(earning, {
    type: 'doughnut',
    data: {
      labels: desig_name,
      datasets: [{
        label: '# of Votes',
        data: desig_Scores,
        backgroundColor: ['rgba(54, 162, 235, 0.6)', 'rgba(255, 99, 132, 0.6)'],
        borderColor: ['rgba(54, 162, 235, 1)', 'rgba(255, 99, 132, 1)'],
        borderWidth: 1
      }]
    },
    options: {
      scales: {
        y: {
          beginAtZero: true
        }
      }
    }
  });
}else{
  earning.parentElement.style.display = "none";
}
