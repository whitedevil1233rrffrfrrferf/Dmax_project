function toggleProjectTable() {
  const projectSelectValue = document.getElementById("project").value;

  const projectSections = {
    Akyrian: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    Auxo: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    "Geek Ladder": [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    Avanti: ["test_cases_section", "defects_section", "others_section"],
    Bench: ["training_section"],
    Training: ["training_section"],
    "Fora Travels": [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    Indihood: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
      "indihood_test_scripts_section",
    ],
    IPS: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    IQHive: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    LevelBlue: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
    "Opus Clip": ["test_cases_section", "defects_section", "others_section"],
    "Web Development": ["web_section"],
    Trademo: [
      "test_cases_section",
      "defects_section",
      "others_section",
      "training_section",
    ],
    Heymax: [
      "test_cases_section",
      "defects_section",
      "others_section",
      "training_section",
    ],
    ONECLICKLCA: [
      "test_cases_section",
      "defects_section",
      "test_scripts_section",
      "others_section",
    ],
  };
  const allSections = new Set(Object.values(projectSections).flat());
  allSections.forEach((sectionId) => {
    document.getElementById(sectionId).style.display = "none";
  });
  if (projectSelectValue === "Bench") {
    // For Bench, show all sections
    allSections.forEach((sectionId) => {
      document.getElementById(sectionId).style.display = "block";
    });
  } else if (projectSections[projectSelectValue]) {
    projectSections[projectSelectValue].forEach((sectionId) => {
      document.getElementById(sectionId).style.display = "block";
    });
  }
}

// Call the function when the document is loaded
document.addEventListener("DOMContentLoaded", toggleProjectTable);
document.getElementById("myForm").addEventListener("submit", function (event) {
  const confirmed = confirm("Are you sure you want to submit the form?");
  if (!confirmed) {
    event.preventDefault(); // Stop form submission if not confirmed
  } else {
    alert("Form submitted successfully!");
    // Optional: you may not see this alert if the form redirects immediately
  }
});
Docu;
