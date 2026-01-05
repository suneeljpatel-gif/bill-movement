
let masterData = [];

console.log("Fetching Excel...");

fetch("masters/master_data.xlsx")
    .then(res => {
        console.log("Fetch response:", res.status);
        return res.arrayBuffer();
    })
    .then(data => {
        console.log("Excel loaded");

        const workbook = XLSX.read(data, { type: "array" });
        console.log("Sheets:", workbook.SheetNames);

        const sheet1 = XLSX.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[0]]
        );

        const sheet2 = XLSX.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[1]]
        );

        console.log("Sheet1 Data:", sheet1);
        console.log("Sheet2 Data:", sheet2);
    })
    .catch(err => console.error("Excel Error:", err));

fetch("masters/master_data.xlsx")
    .then(res => res.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: "array" });

        const sheet1 = XLSX.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[0]]
        );

        const sheet2 = XLSX.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[1]]
        );

        masterData = sheet1;

        loadBillType(sheet1);
        loadAuthorities(sheet2);
    });

function loadBillType(data) {
    const billType = document.getElementById("billType");
    billType.innerHTML = `<option value="">Select Bill Type</option>`;

    [...new Set(data.map(d => d["Bill Type"]))].forEach(bt => {
        billType.innerHTML += `<option value="${bt}">${bt}</option>`;
    });
}

document.getElementById("billType").addEventListener("change", function () {
    const selected = masterData.find(d => d["Bill Type"] === this.value);

    if (selected) {
        document.getElementById("partyName").value = selected["Party Name"];
        document.getElementById("poNumber").value = selected["SAP PO Number"];
        document.getElementById("vendorCode").value = selected["Vendor Code"];
        document.getElementById("AccountDetail").value = selected["Account Detail"];
    }
});

function loadAuthorities(data) {
    const p = document.getElementById("preparedBy");
    const e = document.getElementById("engIncharge");
    const a = document.getElementById("approvedBy");

    data.forEach(row => {
        p.innerHTML += `<option>${row["Prepared by"]}</option>`;
        e.innerHTML += `<option>${row["Engineering In charge"]}</option>`;
        a.innerHTML += `<option>${row["Approved By"]}</option>`;
    });
}

let records = JSON.parse(localStorage.getItem("billRecords")) || [];
let editIndex = -1;

document.getElementById("billForm").addEventListener("submit", function (e) {
    e.preventDefault();

    const record = {
        billType: billType.value,
        partyName: partyName.value,
        poNumber: poNumber.value,
        vendorCode: vendorCode.value,
        AccountDetail: AccountDetail.value,
        billFrom: billFrom.value,
        billAmount: billAmount.value,
        sesNo: sesNo.value,
        parkInvoice: parkInvoice.value,
        preparedBy: preparedBy.value,
        engIncharge: engIncharge.value,
        approvedBy: approvedBy.value
    };

    if (editIndex === -1) {
        records.push(record);
    } else {
        records[editIndex] = record;
        editIndex = -1;
    }

    localStorage.setItem("billRecords", JSON.stringify(records));
    displayTable();
    this.reset();
});

function displayTable() {
    const tbody = document.querySelector("#dataTable tbody");
    tbody.innerHTML = "";

    records.forEach((r, i) => {
        tbody.innerHTML += `
            <tr>
                <td>${r.billType}</td>
                <td>${r.partyName}</td>
                <td>${r.poNumber}</td>
                <td>${r.billAmount}</td>
                <td>
                    <button class="action-btn edit-btn" onclick="editRow(${i})">Edit</button>
                    <button class="action-btn delete-btn" onclick="deleteRow(${i})">Delete</button>
                    <button class="action-btn pdf-btn" onclick="exportPDF(${i})">PDF</button>
                </td>
            </tr>
        `;
    });
}

function editRow(index) {
    const r = records[index];

    billType.value = r.billType;
    partyName.value = r.partyName;
    poNumber.value = r.poNumber;
    vendorCode.value = r.vendorCode;
    AccountDetail.value = r.AccountDetail;
    billFrom.value = r.billFrom;
    billAmount.value = r.billAmount;
    sesNo.value = r.sesNo;
    parkInvoice.value = r.parkInvoice;
    preparedBy.value = r.preparedBy;
    engIncharge.value = r.engIncharge;
    approvedBy.value = r.approvedBy;

    editIndex = index;
}

function deleteRow(index) {
    if (!confirm("Delete this record?")) return;
    records.splice(index, 1);
    localStorage.setItem("billRecords", JSON.stringify(records));
    displayTable();
}

displayTable();

function exportPDF(index) {
    const r = records[index];

    document.getElementById("pBillType").innerText = r.billType;
    document.getElementById("pParty").innerText = r.party;
    document.getElementById("pPO").innerText = r.po;
    document.getElementById("pAmount").innerText = r.amount;

    const element = document.getElementById("pdfTemplate");

    html2pdf()
        .set({
            margin: 10,
            filename: `BillMovement_${index + 1}.pdf`,
            html2canvas: { scale: 2 },
            jsPDF: { format: 'a4', orientation: 'portrait' }
        })
        .from(element)
        .save();
}
