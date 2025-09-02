const appState = {
  workbook: null,
  quotationName: "",
  flightTicket: "",
  sites: [],
  accommodations: {},
  meals: { ticket: "", accommodation: "", count: 0 },
  trips: []
};

// Step navigation
function nextStep(n) {
  document.querySelectorAll(".step").forEach(s => s.classList.remove("active"));
  document.getElementById("step" + n).classList.add("active");
}
function prevStep(n) { nextStep(n); }

// Upload Excel
document.getElementById("upload").addEventListener("change", function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    appState.workbook = XLSX.read(data, { type: "array" });
    alert("Excel uploaded successfully!");
  };
  reader.readAsArrayBuffer(file);
});

// Final Excel generation
function generateExcel() {
  if (!appState.workbook) {
    alert("Please upload Excel first.");
    return;
  }

  // Collect all data
  appState.quotationName = document.getElementById("quotationName").value;
  appState.flightTicket = document.getElementById("flightTicket").value;
  appState.sites = Array.from(document.querySelectorAll("#step1 input[type=checkbox]:checked"))
                        .map(cb => cb.value);

  appState.accommodations = {
    Cairo: { nights: +document.getElementById("cairoNights").value, price: +document.getElementById("cairoPrice").value },
    Sharm: { nights: +document.getElementById("sharmNights").value, price: +document.getElementById("sharmPrice").value },
    Luxor: { nights: +document.getElementById("luxorNights").value, price: +document.getElementById("luxorPrice").value }
  };

  appState.meals = {
    ticket: document.getElementById("guideTicket").value,
    accommodation: document.getElementById("guideAccommodation").value,
    count: +document.getElementById("mealsCount").value
  };

  appState.trips = Array.from(document.querySelectorAll("#step4 input[type=checkbox]:checked"))
                        .map(cb => cb.value);

  // Apply to workbook
  const ws = appState.workbook.Sheets[appState.workbook.SheetNames[0]];

  // Example cell mapping – adjust as per your template
  XLSX.utils.sheet_add_aoa(ws, [[appState.flightTicket]], { origin: "F10" });
  XLSX.utils.sheet_add_aoa(ws, [appState.sites], { origin: "G10" });

  let row = 20;
  for (const city in appState.accommodations) {
    const { nights, price } = appState.accommodations[city];
    XLSX.utils.sheet_add_aoa(ws, [[nights, price]], { origin: `B${row}` });
    row++;
  }

  XLSX.utils.sheet_add_aoa(ws, [[appState.meals.count]], { origin: "J30" });
  XLSX.utils.sheet_add_aoa(ws, [[appState.meals.ticket, appState.meals.accommodation]], { origin: "K30" });

  XLSX.utils.sheet_add_aoa(ws, [appState.trips], { origin: "H40" });

  // Export
  const wbout = XLSX.write(appState.workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = (appState.quotationName || "EditedQuotation") + ".xlsx";
  link.click();
}
