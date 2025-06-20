var gk_isXlsx = false;
var gk_xlsxFileLookup = {};
var gk_fileData = {};

function filledCell(cell) {
  return cell !== "" && cell != null;
}

function loadFileData(filename) {
  if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
    try {
      var workbook = XLSX.read(gk_fileData[filename], { type: "base64" });
      var firstSheetName = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[firstSheetName];

      // Convert sheet to JSON to filter blank rows
      var jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        blankrows: false,
        defval: "",
      });
      // Filter out blank rows (rows where all cells are empty, null, or undefined)
      var filteredData = jsonData.filter((row) => row.some(filledCell));

      // Heuristic to find the header row by ignoring rows with fewer filled cells than the next row
      var headerRowIndex = filteredData.findIndex(
        (row, index) =>
          row.filter(filledCell).length >=
          filteredData[index + 1]?.filter(filledCell).length
      );
      // Fallback
      if (headerRowIndex === -1 || headerRowIndex > 25) {
        headerRowIndex = 0;
      }

      // Convert filtered JSON back to CSV
      var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex));
      csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
      return csv;
    } catch (e) {
      console.error(e);
      return "";
    }
  }
  return gk_fileData[filename] || "";
}

function getContainerDetails() {
  const containerType = document.getElementById("containerType").value;
  let containerDetails = "<h3><center><u>";
  let containerLength, containerWidth, containerHeight, containerVolume;

  switch (containerType) {
    case "20":
      containerDetails += "20ft Standard Container";
      containerLength = 5898;
      containerWidth = 2352;
      containerHeight = 2393;
      containerVolume = 33.2;
      break;
    case "40":
      containerDetails += "40ft Standard Container";
      containerLength = 12032;
      containerWidth = 2352;
      containerHeight = 2393;
      containerVolume = 67.7;
      break;
    case "20hq":
      containerDetails += "20ft High Cube Container";
      containerLength = 5898;
      containerWidth = 2352;
      containerHeight = 2690;
      containerVolume = 37.0;
      break;
    case "40hq":
      containerDetails += "40ft High Cube Container";
      containerLength = 12032;
      containerWidth = 2352;
      containerHeight = 2698;
      containerVolume = 76.4;
      break;
    default:
      containerDetails += "45ft High Cube Container";
      containerLength = 13544;
      containerWidth = 2352;
      containerHeight = 2698;
      containerVolume = 86.0;
  }
  containerDetails += "</u></center></h3>";

  let totalUtilizedCbm = 0;
  let totalCartonQty = 0;

  function calculateLoadingDetails(
    cartonLength,
    cartonWidth,
    cartonHeight,
    bulgingLength,
    bulgingWidth,
    bulgingHeight,
    orderQty,
    index
  ) {
    if (!cartonLength || !cartonWidth || !cartonHeight || !orderQty) return "";
    let details = "<table border='1' style='width:100%'>";

    // Carton Measurement Dimensions
    details += `<tr><td>Carton Size dimensions are:<br>L = ${cartonLength} mm, W = ${cartonWidth} mm, H = ${cartonHeight} mm</td></tr>`;

    // Additional measurements
    const adjustedLength = cartonLength + (bulgingLength || 0);
    const adjustedWidth = cartonWidth + (bulgingWidth || 0);
    const adjustedHeight = cartonHeight + (bulgingHeight || 0);

    // Calculate loading quantity without flat
    const columns = Math.floor(containerLength / adjustedLength);
    const rowsHorizontally = Math.floor(containerWidth / adjustedWidth);
    const rowsVertically = Math.floor(containerHeight / adjustedHeight);
    const loadingQty1 = columns * rowsHorizontally * rowsVertically;

    details += `<tr><td class="td"><font class="fc">Loading qty without flat: ${loadingQty1} Boxes</font></td></tr>`;
    details += `<tr><td class="td">Number of columns: ${columns}<br/> Number of rows horizontally: ${rowsHorizontally}<br/> Number of rows vertically: ${rowsVertically}</td></tr>`;

    // Calculate flat loading quantity of rest of the length
    const remainingLength = containerLength % adjustedLength;
    const flatColumns = Math.floor(remainingLength / adjustedWidth);
    const flatRowsHorizontally = Math.floor(containerWidth / adjustedLength);
    const flatRowsVertically = Math.floor(containerHeight / adjustedHeight);
    const loadingQty2 = flatColumns * flatRowsHorizontally * flatRowsVertically;

    details += `<tr><td class="td1"><b>Flat loading qty rest of the length: ${loadingQty2} Boxes</b></td></tr>`;
    details += `<tr><td class="td1">Number of flat columns rest length: ${flatColumns}<br/> Number of flat rows horizontally: ${flatRowsHorizontally}<br/> Number of flat rows vertically: ${flatRowsVertically}</td></tr>`;

    // Calculate flat loading quantity rest of the height
    const remainingHeight = containerHeight % adjustedHeight;
    const flatColumnsHeight = Math.floor(remainingHeight / adjustedWidth);
    const flatRowsHeightHorizontally = Math.floor(
      containerWidth / adjustedHeight
    );
    const flatRowsHeightLength = Math.floor(containerLength / adjustedLength);
    const loadingQty3 =
      flatColumnsHeight * flatRowsHeightHorizontally * flatRowsHeightLength;

    details += `<tr><td class="td2"><b>Flat loading qty rest of the height: ${loadingQty3} Boxes</b></td></tr>`;
    details += `<tr><td class="td2">Number of flat columns rest height: ${flatRowsHeightLength}<br/> Number of flat rows horizontally: ${flatRowsHeightHorizontally}<br/> Number of flat rows vertically: ${flatColumnsHeight}</td></tr>`;
    details += `<tr><td class="td"><font class="fc2">Total Loading Qty with flat: ${
      loadingQty1 + loadingQty2 + loadingQty3
    } Boxes</font></td></tr>`;

    // Calculate Utilized CBM and Empty CBM
    const utilizedCbm = (
      (adjustedLength * adjustedWidth * adjustedHeight * orderQty) /
      1000000000
    ).toFixed(2);
    const emptyCbm = (containerVolume - utilizedCbm).toFixed(2);

    details += `<tr><td>Utilized CBM: ${utilizedCbm} CBM</td></tr>`;
    details += `<tr><td>Empty CBM: ${emptyCbm} CBM</td></tr>`;
    details += "</table>";

    // Update table cells
    document.getElementById(`utilizedCbmVal${index}`).innerText = utilizedCbm;
    document.getElementById(`emptyCbmVal${index}`).innerText = emptyCbm;

    // Add to total utilized CBM
    totalUtilizedCbm += parseFloat(utilizedCbm);

    return details;
  }

  // Process all 12 cartons
  const cartons = Array.from({ length: 12 }, (_, i) => i + 1).map((i) => ({
    cartonLength:
      parseFloat(document.getElementById(`cartonLength${i}`).value) || 0,
    cartonWidth:
      parseFloat(document.getElementById(`cartonWidth${i}`).value) || 0,
    cartonHeight:
      parseFloat(document.getElementById(`cartonHeight${i}`).value) || 0,
    orderQty: parseFloat(document.getElementById(`orderQty${i}`).value) || 0,
    bulgingLength:
      parseFloat(document.getElementById(`bulgingLength${i}`).value) || 0,
    bulgingWidth:
      parseFloat(document.getElementById(`bulgingWidth${i}`).value) || 0,
    bulgingHeight:
      parseFloat(document.getElementById(`bulgingHeight${i}`).value) || 0,
    index: i,
  }));

  // Calculate total carton quantity
  totalCartonQty = cartons.reduce((sum, carton) => sum + carton.orderQty, 0);

  // Generate details for each pair of cartons
  for (let i = 0; i < cartons.length; i += 2) {
    containerDetails += `<table border='1' style='width:100%'><tr><th><br/>Carton Size ${String(
      i + 1
    ).padStart(2, "0")} Details</th><th><br/>Carton Size ${String(
      i + 2
    ).padStart(2, "0")} Details</th></tr>`;
    containerDetails += "<tr>";
    containerDetails += `<td>${calculateLoadingDetails(
      cartons[i].cartonLength,
      cartons[i].cartonWidth,
      cartons[i].cartonHeight,
      cartons[i].bulgingLength,
      cartons[i].bulgingWidth,
      cartons[i].bulgingHeight,
      cartons[i].orderQty,
      cartons[i].index
    )}</td>`;
    containerDetails += `<td>${calculateLoadingDetails(
      cartons[i + 1]?.cartonLength || 0,
      cartons[i + 1]?.cartonWidth || 0,
      cartons[i + 1]?.cartonHeight || 0,
      cartons[i + 1]?.bulgingLength || 0,
      cartons[i + 1]?.bulgingWidth || 0,
      cartons[i + 1]?.bulgingHeight || 0,
      cartons[i + 1]?.orderQty || 0,
      cartons[i + 1]?.index || i + 2
    )}</td>`;
    containerDetails += "</tr></table><br>";
  }

  // Update total CBM and carton quantity values
  document.getElementById("totalUtilizedCbm").innerText =
    totalUtilizedCbm.toFixed(2);
  document.getElementById("totalEmptyCbm").innerText = (
    containerVolume - totalUtilizedCbm
  ).toFixed(2);
  document.getElementById("totalCartonQty").innerText = totalCartonQty;

  document.getElementById("containerDetails").innerHTML = containerDetails;
}

function clearData() {
  document.getElementById("containerType").value = "20";
  for (let i = 1; i <= 12; i++) {
    document.getElementById(`cartonLength${i}`).value = "";
    document.getElementById(`cartonWidth${i}`).value = "";
    document.getElementById(`cartonHeight${i}`).value = "";
    document.getElementById(`orderQty${i}`).value = "";
    document.getElementById(`bulgingLength${i}`).value = "0";
    document.getElementById(`bulgingWidth${i}`).value = "12";
    document.getElementById(`bulgingHeight${i}`).value = "3";
    document.getElementById(`utilizedCbmVal${i}`).innerText = "";
    document.getElementById(`emptyCbmVal${i}`).innerText = "";
  }
  document.getElementById("totalUtilizedCbm").innerText = "0.00";
  document.getElementById("totalEmptyCbm").innerText = "0.00";
  document.getElementById("totalCartonQty").innerText = "0";
  document.getElementById("containerDetails").innerHTML = "";
}