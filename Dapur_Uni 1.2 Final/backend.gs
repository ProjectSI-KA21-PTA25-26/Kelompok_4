// =======================================
// ROUTER
// =======================================
function doGet(e) {
  const action = e.parameter.action || "";

  if (action === "getMenu") {
    return getMenu();
  }

  if (action === "getBookedDates") {
    return getBookedDates();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = e.parameter.action || "";

  if (action === "submitOrder") {
    return submitOrder(e);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =======================================
// GET MENU
// =======================================
function getMenu() {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("Menu");

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getValues();
  values.shift();

  const menu = values
    .filter(row => row[11] !== false)
    .map(row => ({
      id: row[0],
      title: row[1],
      img: row[2],
      desc: row[3],
      price: row[4],
      priceUnit: row[5],
      portion: row[6],
      food: row[7] ? row[7].split(";").map(f => f.trim()) : [],
      category: row[8],
      tags: row[9] ? row[9].split(";").map(t => t.trim()) : [],
      minPax: row[10]
    }));

  return ContentService
    .createTextOutput(JSON.stringify(menu))
    .setMimeType(ContentService.MimeType.JSON);
}

// =======================================
// GET BOOKED DATES
// =======================================
function getBookedDates() {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("Orders");

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ fullBookDates: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getValues();
  values.shift(); // header

  const map = {};

  values.forEach(row => {
    const d = row[4];        // tanggalAcara
    const status = row[14]; // statusOrder

    if (!d || !status) return;
    if (status === "Cancelled" || status === "Batal") return;

    let dateObj;

    if (d instanceof Date) {
      dateObj = d;
    } else {
      dateObj = new Date(d);
    }

    if (isNaN(dateObj.getTime())) return;

    const dateStr = Utilities.formatDate(
      dateObj,
      "Asia/Jakarta",
      "yyyy-MM-dd"
    );

    if (!map[dateStr]) {
      map[dateStr] = { pending: 0, confirmed: 0 };
    }

    if (status === "Confirmed") {
      map[dateStr].confirmed++;
    } else if (status === "Pending") {
      map[dateStr].pending++;
    }
  });

  const fullBookDates = Object.keys(map).filter(date => {
    const { pending, confirmed } = map[date];
    return confirmed >= 2 || (pending + confirmed) >= 3;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ fullBookDates }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =======================================
// FORMAT MENU (HUMAN READABLE)
// =======================================
function formatMenuReadable(menuData) {
  if (!Array.isArray(menuData)) return "";

  return menuData
    .map(m => `${m.title} (${m.pax} pax)`)
    .join(", ");
}

// =======================================
// SUBMIT ORDER
// =======================================
function submitOrder(e) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("Orders");

  if (!e.postData || !e.postData.contents) {
  return ContentService
    .createTextOutput(JSON.stringify({ error: "No post data" }))
    .setMimeType(ContentService.MimeType.JSON);
}

const data = JSON.parse(e.postData.contents);

sheet.appendRow([
  data.orderId,
  new Date(),
  data.namaPemesan,
  data.noWa,
  data.tanggalAcara,
  data.jamAcara,
  data.tipePengiriman,
  data.kota,
  data.kecamatan,
  data.kelurahan,
  data.alamat,
  formatMenuReadable(data.menuData),
  data.totalPax,
  data.metodeBayar,
  "Pending",
  data.catatan || ""
]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
