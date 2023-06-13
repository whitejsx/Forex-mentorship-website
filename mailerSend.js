document.getElementById("myForm").addEventListener("submit", function(event) {
    event.preventDefault(); // Prevent form submission
  
    // Get form data
    var name = document.getElementById("name").value;
    var email = document.getElementById("email").value;
    var message = document.getElementById("message").value;
  
    // Create a workbook
    var workbook = XLSX.utils.book_new();
    var sheetData = [
      ["Name", "Email", "Message"],
      [name, email, message]
    ];
    var worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Form Data");
  
    // Convert workbook to binary data
    var excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  
    // Save the Excel file
    saveExcelFile(excelBuffer, "form_data.xlsx");
  
    // Clear the form
    document.getElementById("myForm").reset();
  });
  
  function saveExcelFile(buffer, filename) {
    var blob = new Blob([buffer], { type: "application/octet-stream" });
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
  