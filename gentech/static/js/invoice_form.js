(function () {
  const customerSelect = document.getElementById("customer_id");
  const customerName = document.getElementById("customer_name");
  const customerAddress = document.getElementById("customer_address");
  const customerGstin = document.getElementById("customer_gstin");
  const referenceInput = document.getElementById("customer_reference");
  const transportInput = document.getElementById("transport");

  const table = document.getElementById("line-items");
  const tbody = table.querySelector("tbody");
  const addRowButton = document.getElementById("add-line-item");
  const maxRows = Number.parseInt(table.dataset.maxRows || "8", 10);

  const items = JSON.parse(document.getElementById("items-master").textContent || "[]");

  function normalize(text) {
    return String(text || "").trim().toLowerCase();
  }

  const itemMap = new Map();
  items.forEach((item) => {
    itemMap.set(normalize(item.description), item);
  });

  function buildDescriptionDataList() {
    const dataList = document.createElement("datalist");
    dataList.id = "item-descriptions";
    items.forEach((item) => {
      const option = document.createElement("option");
      option.value = item.description;
      dataList.appendChild(option);
    });
    document.body.appendChild(dataList);
  }

  function asNumber(value) {
    const parsed = Number.parseFloat(String(value || "").replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function applyDescriptionDefaults(row, force) {
    const descriptionInput = row.querySelector(".desc-input");
    const hsnInput = row.querySelector(".hsn-input");
    const priceInput = row.querySelector(".price-input");
    const item = itemMap.get(normalize(descriptionInput.value));

    if (!item) {
      return;
    }

    if (force || !String(hsnInput.value || "").trim()) {
      hsnInput.value = item.hsn_sac || "";
    }

    if (force || !String(priceInput.value || "").trim()) {
      priceInput.value = String(item.default_unit_price ?? "");
    }
  }

  function renumberRows() {
    [...tbody.querySelectorAll("tr")].forEach((row, index) => {
      row.querySelector(".sl-no").textContent = String(index + 1);
    });
  }

  function recalculate() {
    let subtotal = 0;

    tbody.querySelectorAll("tr").forEach((row) => {
      applyDescriptionDefaults(row, false);
      const qtyInput = row.querySelector(".qty-input");
      const priceInput = row.querySelector(".price-input");
      const amountInput = row.querySelector(".amount-input");

      const qty = asNumber(qtyInput.value);
      const price = asNumber(priceInput.value);
      const amount = qty * price;
      amountInput.value = amount ? amount.toFixed(2) : "";
      subtotal += amount;
    });

    const transport = asNumber(transportInput.value);
    const cgst = subtotal * 0.09;
    const sgst = subtotal * 0.09;
    const grand = subtotal + cgst + sgst + transport;

    document.getElementById("total_amount").textContent = subtotal.toFixed(2);
    document.getElementById("cgst_total").textContent = cgst.toFixed(2);
    document.getElementById("sgst_total").textContent = sgst.toFixed(2);
    document.getElementById("transport_total").textContent = transport.toFixed(2);
    document.getElementById("grand_total").textContent = grand.toFixed(2);
  }

  function bindRowEvents(row) {
    const descriptionInput = row.querySelector(".desc-input");
    const qtyInput = row.querySelector(".qty-input");
    const hsnInput = row.querySelector(".hsn-input");
    const priceInput = row.querySelector(".price-input");
    const removeButton = row.querySelector(".remove-row-btn");

    descriptionInput.addEventListener("change", function () {
      applyDescriptionDefaults(row, false);
      recalculate();
    });
    descriptionInput.addEventListener("blur", function () {
      applyDescriptionDefaults(row, false);
      recalculate();
    });

    [qtyInput, hsnInput, priceInput].forEach((input) => {
      input.addEventListener("input", recalculate);
      input.addEventListener("change", recalculate);
    });

    removeButton.addEventListener("click", function () {
      row.remove();
      if (!tbody.querySelector("tr")) {
        addItemRow();
      }
      renumberRows();
      recalculate();
    });
  }

  function addItemRow() {
    const rowCount = tbody.querySelectorAll("tr").length;
    if (rowCount >= maxRows) {
      window.alert("Maximum line item rows reached for this invoice.");
      return;
    }

    const row = document.createElement("tr");
    row.innerHTML =
      '<td class="sl-no"></td>' +
      '<td><input name="description" class="desc-input" list="item-descriptions" placeholder="Select or type description" /></td>' +
      '<td><input name="qty" class="qty-input" /></td>' +
      '<td><input name="hsn_sac" class="hsn-input" /></td>' +
      '<td><input name="unit_price" class="price-input" /></td>' +
      '<td><input class="amount-input" readonly /></td>' +
      '<td><button type="button" class="remove-row-btn">Remove</button></td>';

    tbody.appendChild(row);
    bindRowEvents(row);
    renumberRows();
    recalculate();
  }

  function fillCustomerFromSelection() {
    const selected = customerSelect.options[customerSelect.selectedIndex];
    if (!selected || !selected.value) {
      return;
    }

    customerName.value = selected.dataset.name || "";
    customerAddress.value = selected.dataset.address || "";
    customerGstin.value = selected.dataset.gstin || "";

    const selectedReference = String(selected.dataset.reference || "").trim();
    referenceInput.value = selectedReference || "";
  }

  buildDescriptionDataList();

  addRowButton.addEventListener("click", addItemRow);
  customerSelect.addEventListener("change", fillCustomerFromSelection);
  transportInput.addEventListener("input", recalculate);
  transportInput.addEventListener("change", recalculate);

  addItemRow();

  if (customerSelect.value) {
    fillCustomerFromSelection();
  }
  recalculate();
})();
