const labels = chartData.map(item => item.name);
const scores = chartData.map(item => item.score);

const ctx = document.getElementById('myChart');
const earning=document.getElementById('earning')
function isValidData(labels, data) {
  return labels.length > 0 && data.some(value => value !== 0);
}

const emptyLabels = ['No Data'];
const emptyScores = [0];

const chart = new Chart(ctx, {
  type: 'bar',
  data: {
    labels: isValidData(labels, scores) ? labels : emptyLabels,
    datasets: [{
      label: 'Dmax Score',
      data: isValidData(labels, scores) ? scores : emptyScores,
      backgroundColor: isValidData(labels, scores) ? 
        ['rgba(212, 253, 217, 0.8)', 'rgba(148, 223, 166, 0.8)'] : 
        ['rgba(200, 200, 200, 0.5)'],  // Light gray for empty graph
      borderColor: isValidData(labels, scores) ? 
        ['rgba(77, 168, 218, 1)', 'rgba(255, 140, 157, 1)'] : 
        ['rgba(150, 150, 150, 1)'],  // Gray border for empty graph
      borderWidth: 1
    }]
  },
  options: {
    scales: {
      y: {
        beginAtZero: true,
        suggestedMax: 10  // Forces a visible Y-axis
      }
    },
    plugins: {
      legend: {
        display: true  // Ensure the legend is visible
      }
    }
  }
});


const desig_emptyLabels = ['No Data'];
const desig_emptyScores = [1];

new Chart(earning, {
  type: 'doughnut',
  data: {
    labels: isValidData(desig_name, desig_Scores) ? desig_name : desig_emptyLabels,
    datasets: [{
      label: 'No',
      data: isValidData(desig_name, desig_Scores) ? desig_Scores : desig_emptyScores,
      backgroundColor: isValidData(desig_name, desig_Scores) ? 
        ['rgba(54, 162, 235, 0.6)', 'rgba(255, 99, 132, 0.6)'] : 
        ['rgba(200, 200, 200, 0.5)'],  // Light gray for empty graph
      borderColor: isValidData(desig_name, desig_Scores) ? 
        ['rgba(54, 162, 235, 1)', 'rgba(255, 99, 132, 1)'] : 
        ['rgba(150, 150, 150, 1)'],  // Gray border for empty graph
      borderWidth: 1
    }]
  },
  options: {
    plugins: {
      legend: {
        display: true  // Ensures legend is visible
      }
    }
  }
});
