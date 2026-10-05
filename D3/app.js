/* =========================================================
   ONLINE RETAIL ANALYTICS DASHBOARD
   XLSB VERSION - FIXED AXES / ZOOM / DONUT LEGEND
   ========================================================= */

const DATA_URL = new URL(
    "Online_Retail_Cleaned_Final-1.xlsb",
    document.baseURI
).href;

let globalDataset = [];
let filteredData = [];

let selectedCountry = "ALL";
let selectedType = "ALL";
let selectedYear = "ALL";
let selectedMonth = "ALL";

let localProductMetric = "value";
let localProductTop = 10;
let localDonutTop = 5;
let localScatterPrice = "all";
let localScatterQty = 0;

const COUNTRY_COLOR_MAP = {};
const COUNTRY_HUE_STEP = 137.508;

const TYPE_COLORS = {
    "Sale": "#55b99b",
    "Return/Cancelled": "#e88ca7"
};

const BAR_COLORS = [
    "#9c8be2", "#8b7bd8", "#b18bd1", "#c19bd9",
    "#edc17f", "#71c6ac", "#d9a6c5", "#a7a0df",
    "#e5b6c8", "#8fc9b5"
];

const MONTH_NAMES_TH = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน",
    "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม",
    "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
];

document.addEventListener("DOMContentLoaded", () => {
    setupEvents();
    loadData();
});

/* =========================================================
   HELPERS
   ========================================================= */

function getCountryColor(country) {
    if (!COUNTRY_COLOR_MAP[country]) {
        const index = Object.keys(COUNTRY_COLOR_MAP).length;
        const hue = (index * COUNTRY_HUE_STEP) % 360;
        COUNTRY_COLOR_MAP[country] =
            `hsl(${hue.toFixed(1)}, 58%, 58%)`;
    }
    return COUNTRY_COLOR_MAP[country];
}

function getContainerSize(node, fallbackWidth = 600, fallbackHeight = 350) {
    const bounds = node.getBoundingClientRect();
    return {
        width: Math.max(fallbackWidth, bounds.width || fallbackWidth),
        height: Math.max(fallbackHeight, bounds.height || fallbackHeight)
    };
}

function addAxisLabel(g, {
    x = 0,
    y = 0,
    text = "",
    anchor = "middle",
    rotate = null
} = {}) {
    const label = g.append("text")
        .attr("class", "axis-label")
        .attr("x", x)
        .attr("y", y)
        .attr("text-anchor", anchor)
        .style("font-size", "12px")
        .style("font-weight", "600")
        .style("fill", "#806f89")
        .text(text);

    if (rotate !== null) {
        label.attr("transform", `rotate(${rotate},${x},${y})`);
    }

    return label;
}

function safeText(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/* =========================================================
   EVENTS
   ========================================================= */

function setupEvents() {
    d3.select("#countryFilter").on("change", function () {
        selectedCountry = this.value;
        applyFilters();
    });

    d3.select("#typeFilter").on("change", function () {
        selectedType = this.value;
        applyFilters();
    });

    d3.select("#yearFilter").on("change", function () {
        selectedYear = this.value;
        selectedMonth = "ALL";
        populateMonthDropdown();
        d3.select("#monthFilter").property("value", "ALL");
        applyFilters();
    });

    d3.select("#monthFilter").on("change", function () {
        selectedMonth = this.value;
        applyFilters();
    });

    d3.select("#localProductMetric").on("change", function () {
        localProductMetric = this.value;
        renderBarChart();
    });

    d3.select("#localProductTop").on("change", function () {
        localProductTop = Number(this.value);
        renderBarChart();
    });

    d3.select("#localDonutTop").on("change", function () {
        localDonutTop = Number(this.value);
        renderDonutChart();
    });

    d3.select("#localScatterPrice").on("change", function () {
        localScatterPrice = this.value;
        renderScatterChart();
    });

    d3.select("#localScatterQty").on("change", function () {
        localScatterQty = Number(this.value);
        renderScatterChart();
    });

    d3.select("#resetBtn").on("click", resetDashboard);

    window.addEventListener("resize", () => {
        if (globalDataset.length) renderCharts();
    });
}

/* =========================================================
   LOAD XLSB
   ========================================================= */

function loadExternalScript(src, globalName) {
    return new Promise((resolve, reject) => {
        if (globalName && window[globalName]) {
            resolve();
            return;
        }

        const existing = document.querySelector(`script[src="${src}"]`);
        if (existing) {
            existing.addEventListener("load", resolve, { once: true });
            existing.addEventListener("error", () => reject(new Error(`โหลดไลบรารีไม่ได้: ${src}`)), { once: true });
            return;
        }

        const script = document.createElement("script");
        script.src = src;
        script.async = false;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`โหลดไลบรารีไม่ได้: ${src}`));
        document.head.appendChild(script);
    });
}

async function ensureLibraries() {
    if (typeof d3 === "undefined") {
        await loadExternalScript("https://d3js.org/d3.v7.min.js", "d3");
    }

    if (typeof XLSX === "undefined") {
        await loadExternalScript(
            "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
            "XLSX"
        );
    }
}

async function loadData() {
    try {
        console.log("กำลังเตรียม D3 + SheetJS...");
        await ensureLibraries();
        console.log("กำลังโหลด:", DATA_URL);

        if (typeof XLSX === "undefined") {
            throw new Error("ไม่พบ SheetJS (XLSX)");
        }

        const response = await fetch(
            DATA_URL + "?v=" + Date.now(),
            { cache: "no-store" }
        );

        if (!response.ok) {
            throw new Error(
                "โหลดไฟล์ XLSB ไม่ได้: HTTP " + response.status
            );
        }

        const buffer = await response.arrayBuffer();

        if (!buffer || buffer.byteLength === 0) {
            throw new Error("ไฟล์ XLSB ว่าง");
        }

        console.log("XLSB size:", buffer.byteLength, "bytes");

        const workbook = XLSX.read(buffer, {
            type: "array",
            cellDates: true,
            dense: true
        });

        if (!workbook.SheetNames || !workbook.SheetNames.length) {
            throw new Error("ไม่พบ Sheet ในไฟล์ XLSB");
        }

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
            throw new Error("ไม่พบข้อมูลใน Sheet");
        }

        console.log("Sheet:", sheetName);

        const rows = XLSX.utils.sheet_to_json(worksheet, {
            defval: "",
            raw: true
        });

        if (!rows.length) {
            throw new Error("Sheet ไม่มีข้อมูล");
        }

        console.log("Rows:", rows.length);
        console.log("Columns:", Object.keys(rows[0]));

        convertData(rows);

    } catch (error) {
        console.error("XLSB ERROR:", error);
        showLoadError(error);
    }
}

/* =========================================================
   CONVERT DATA
   ========================================================= */

function convertData(rows) {
    globalDataset = rows.map(d => {
        const quantity = Number(d.Quantity) || 0;
        const unitPrice = Number(d.UnitPrice) || 0;

        let lineAmount = Number(d.LineAmount);

        if (!Number.isFinite(lineAmount)) {
            lineAmount = quantity * unitPrice;
        }

        const date = parseInvoiceDate(d.InvoiceDate);

        let transactionType =
            String(d.TransactionType || "").trim();

        if (!transactionType) {
            transactionType =
                quantity < 0
                    ? "Return/Cancelled"
                    : "Sale";
        }

        const typeLower = transactionType.toLowerCase();

        transactionType =
            typeLower.includes("return") ||
            typeLower.includes("cancel")
                ? "Return/Cancelled"
                : "Sale";

        return {
            Description: d.Description || "Uncategorized",
            Quantity: quantity,
            UnitPrice: unitPrice,
            CustomerID: d.CustomerID,
            Country: d.Country || "Unknown",
            TransactionType: transactionType,
            LineAmount: lineAmount,
            InvoiceDate: d.InvoiceDate,
            Date: date,
            Year: date ? date.getFullYear() : null,
            Month: date ? date.getMonth() + 1 : null
        };
    }).filter(d =>
        d.Date instanceof Date &&
        !Number.isNaN(d.Date.getTime())
    );

    if (!globalDataset.length) {
        throw new Error(
            "อ่าน XLSB ได้ แต่ไม่พบข้อมูลวันที่ที่ถูกต้อง"
        );
    }

    filteredData = [...globalDataset];

    populateCountryDropdown();
    populateYearDropdown();
    populateMonthDropdown();
    updateDashboard();

    console.log("=================================");
    console.log("XLSB LOADED SUCCESSFULLY");
    console.log("จำนวนข้อมูล:", globalDataset.length);
    console.log(
        "ประเทศ:",
        new Set(globalDataset.map(d => d.Country)).size
    );
    console.log("=================================");
}

/* =========================================================
   DATE PARSER
   ========================================================= */

function parseInvoiceDate(value) {
    if (
        value instanceof Date &&
        !Number.isNaN(value.getTime())
    ) {
        return value;
    }

    if (value === null || value === undefined || value === "") {
        return null;
    }

    if (typeof value === "number" && Number.isFinite(value)) {
        const excelEpoch = new Date(
            Date.UTC(1899, 11, 30)
        );

        const date = new Date(
            excelEpoch.getTime() +
            value * 86400000
        );

        return Number.isNaN(date.getTime())
            ? null
            : date;
    }

    const raw = String(value).trim();

    const nativeDate = new Date(raw);

    if (!Number.isNaN(nativeDate.getTime())) {
        return nativeDate;
    }

    const formats = [
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M"
    ];

    for (const format of formats) {
        const date = d3.timeParse(format)(raw);
        if (date) return date;
    }

    return null;
}

/* =========================================================
   ERROR
   ========================================================= */

function showLoadError(error) {
    const message =
        error && error.message
            ? error.message
            : String(error);

    [
        "#barChart",
        "#donutChart",
        "#columnChart",
        "#scatterChart"
    ].forEach(selector => {
        const box = d3.select(selector);
        if (box.empty()) return;

        box.html("");

        box.append("div")
            .style("padding", "30px")
            .style("text-align", "center")
            .style("color", "#dc2626")
            .style("font-weight", "600")
            .text("โหลดข้อมูลไม่สำเร็จ: " + message);
    });
}

/* =========================================================
   DROPDOWNS
   ========================================================= */

function populateCountryDropdown() {
    const select = d3.select("#countryFilter");

    select.selectAll("option:not(:first-child)").remove();

    Array.from(
        new Set(globalDataset.map(d => d.Country))
    )
        .filter(Boolean)
        .sort()
        .forEach(country => {
            select.append("option")
                .attr("value", country)
                .text(country);
        });
}

function populateYearDropdown() {
    const select = d3.select("#yearFilter");

    select.selectAll("option:not(:first-child)").remove();

    Array.from(
        new Set(
            globalDataset
                .map(d => d.Year)
                .filter(Boolean)
        )
    )
        .sort((a, b) => a - b)
        .forEach(year => {
            select.append("option")
                .attr("value", year)
                .text(year);
        });
}

function populateMonthDropdown() {
    const select = d3.select("#monthFilter");

    select.selectAll("option:not(:first-child)").remove();

    const months = Array.from(
        new Set(
            globalDataset
                .filter(d =>
                    selectedYear === "ALL" ||
                    String(d.Year) === String(selectedYear)
                )
                .map(d => d.Month)
                .filter(Boolean)
        )
    ).sort((a, b) => a - b);

    months.forEach(month => {
        select.append("option")
            .attr("value", month)
            .text(
                String(month).padStart(2, "0") +
                " - " +
                MONTH_NAMES_TH[month - 1]
            );
    });

    if (
        selectedMonth !== "ALL" &&
        !months.includes(Number(selectedMonth))
    ) {
        selectedMonth = "ALL";
    }

    select.property("value", selectedMonth);
}

/* =========================================================
   FILTERS
   ========================================================= */

function applyFilters() {
    filteredData = globalDataset.filter(d => {
        const matchCountry =
            selectedCountry === "ALL" ||
            d.Country === selectedCountry;

        const matchType =
            selectedType === "ALL" ||
            d.TransactionType === selectedType;

        const matchYear =
            selectedYear === "ALL" ||
            String(d.Year) === String(selectedYear);

        const matchMonth =
            selectedMonth === "ALL" ||
            String(d.Month) === String(selectedMonth);

        return (
            matchCountry &&
            matchType &&
            matchYear &&
            matchMonth
        );
    });

    updateDashboard();
}

function resetDashboard() {
    selectedCountry = "ALL";
    selectedType = "ALL";
    selectedYear = "ALL";
    selectedMonth = "ALL";

    localProductMetric = "value";
    localProductTop = 10;
    localDonutTop = 5;
    localScatterPrice = "all";
    localScatterQty = 0;

    d3.select("#countryFilter").property("value", "ALL");
    d3.select("#typeFilter").property("value", "ALL");
    d3.select("#yearFilter").property("value", "ALL");
    d3.select("#localProductMetric").property("value", "value");
    d3.select("#localProductTop").property("value", "10");
    d3.select("#localDonutTop").property("value", "5");
    d3.select("#localScatterPrice").property("value", "all");
    d3.select("#localScatterQty").property("value", "0");

    populateMonthDropdown();
    d3.select("#monthFilter").property("value", "ALL");

    filteredData = [...globalDataset];
    updateDashboard();
}

/* =========================================================
   DASHBOARD
   ========================================================= */

function updateDashboard() {
    renderKPIs();
    renderCharts();
}

function renderKPIs() {
    const totalSales = d3.sum(
        filteredData,
        d => Number(d.LineAmount) || 0
    );

    const totalQty = d3.sum(
        filteredData,
        d => Number(d.Quantity) || 0
    );

    const totalOrders = filteredData.length;

    const totalCountries = new Set(
        filteredData.map(d => d.Country)
    ).size;

    d3.select("#kpiTotalSales")
        .text(`£${d3.format(",.2f")(totalSales)}`);

    d3.select("#kpiTotalQty")
        .text(`${d3.format(",")(totalQty)} ชิ้น`);

    d3.select("#kpiTotalOrders")
        .text(`${d3.format(",")(totalOrders)} รายการ`);

    d3.select("#kpiTotalCountries")
        .text(`${totalCountries} ประเทศ`);
}

function renderCharts() {
    renderBarChart();
    renderDonutChart();
    renderColumnChart();
    renderScatterChart();
}

/* =========================================================
   TOOLTIP
   ========================================================= */

const tooltip = d3.select("#tooltip");

function showTooltip(event, content) {
    if (tooltip.empty()) return;

    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;

    let left = event.pageX + 15;
    let top = event.pageY - 28;

    if (left > window.scrollX + viewportWidth - 300) {
        left = event.pageX - 295;
    }

    if (top > window.scrollY + viewportHeight - 150) {
        top = event.pageY - 150;
    }

    tooltip
        .html(content)
        .style("opacity", 1)
        .style("left", left + "px")
        .style("top", top + "px");
}

function hideTooltip() {
    if (!tooltip.empty()) {
        tooltip.style("opacity", 0);
    }
}

/* =========================================================
   1. BAR CHART
   ========================================================= */

function renderBarChart() {
    const container = d3.select("#barChart");
    if (container.empty()) return;

    container.html("");

    const node = container.node();
    const bounds = node.getBoundingClientRect();

    const margin = {
        top: 20,
        right: 30,
        bottom: 58,
        left: 205
    };

    const width = Math.max(
        200,
        bounds.width - margin.left - margin.right
    );

    const height = Math.max(
        180,
        bounds.height - margin.top - margin.bottom
    );

    const productData = Array.from(
        d3.rollup(
            filteredData,
            values => ({
                Sales: d3.sum(
                    values,
                    d => Number(d.LineAmount) || 0
                ),
                Qty: d3.sum(
                    values,
                    d => Number(d.Quantity) || 0
                )
            }),
            d => d.Description
        ),
        ([description, stats]) => ({
            FullDesc: description,
            ShortDesc:
                description.length > 25
                    ? description.substring(0, 22) + "..."
                    : description,
            Sales: stats.Sales,
            Qty: stats.Qty
        })
    )
    .sort((a, b) =>
        localProductMetric === "qty"
            ? b.Qty - a.Qty
            : b.Sales - a.Sales
    )
    .slice(0, localProductTop);

    if (!productData.length) {
        container.append("div")
            .style("padding", "80px")
            .style("text-align", "center")
            .text("ไม่มีข้อมูล");
        return;
    }

    const svg = container.append("svg")
        .attr("width", "100%")
        .attr("height", "100%")
        .attr("viewBox", `0 0 ${bounds.width} ${bounds.height}`)
        .style("overflow", "visible");

    const g = svg.append("g")
        .attr(
            "transform",
            `translate(${margin.left},${margin.top})`
        );

    const y = d3.scaleBand()
        .domain(productData.map(d => d.ShortDesc))
        .range([0, height])
        .padding(0.22);

    const metricMax = d3.max(
        productData,
        d =>
            localProductMetric === "qty"
                ? d.Qty
                : d.Sales
    ) || 1;

    const x = d3.scaleLinear()
        .domain([0, metricMax * 1.1])
        .nice()
        .range([0, width]);

    g.append("g")
        .call(d3.axisLeft(y));

    g.append("g")
        .attr("transform", `translate(0,${height})`)
        .call(
            d3.axisBottom(x)
                .ticks(5)
                .tickFormat(d =>
                    localProductMetric === "qty"
                        ? d3.format(",")(d)
                        : "£" +
                          d3.format(".0f")(d / 1000) +
                          "k"
                )
        );

    addAxisLabel(g, {
        x: width / 2,
        y: height + 48,
        text:
            localProductMetric === "qty"
                ? "จำนวนชิ้น"
                : "มูลค่ารายการ (£)"
    });

    addAxisLabel(g, {
        x: -height / 2,
        y: -margin.left + 18,
        text: "รายชื่อสินค้า",
        rotate: -90
    });

    g.selectAll(".bar-rect")
        .data(productData)
        .enter()
        .append("rect")
        .attr("class", "bar-rect")
        .attr("x", 0)
        .attr("y", d => y(d.ShortDesc))
        .attr("height", y.bandwidth())
        .attr("rx", 5)
        .attr(
            "fill",
            (d, i) => BAR_COLORS[i % BAR_COLORS.length]
        )
        .attr("width", 0)
        .transition().duration(850).delay((d,i)=>i*55).ease(d3.easeCubicOut)
        .attr(
            "width",
            d => x(localProductMetric === "qty" ? d.Qty : d.Sales)
        )
        .on("mouseover", (event, d) => {
            showTooltip(
                event,
                `<b>${safeText(d.FullDesc)}</b>
                 <br>ยอดขาย: £${d3.format(",.2f")(d.Sales)}
                 <br>จำนวน: ${d3.format(",")(d.Qty)} ชิ้น`
            );
        })
        .on("mouseout", hideTooltip);
}

/* =========================================================
   2. DONUT CHART
   ========================================================= */

function renderDonutChart() {
    const container = d3.select("#donutChart");
    if (container.empty()) return;

    container.html("");

    const node = container.node();
    const bounds = node.getBoundingClientRect();

    const fullWidth = Math.max(360, bounds.width || 600);
    const height = Math.max(350, bounds.height || 350);

    const donutWidth =
        fullWidth <= 600
            ? Math.floor(fullWidth * 0.56)
            : Math.floor(fullWidth * 0.58);

    const legendWidth = fullWidth - donutWidth - 10;

    const countryRollup = Array.from(
        d3.rollup(
            filteredData,
            values => ({
                Value: d3.sum(
                    values,
                    d => Number(d.LineAmount) || 0
                ),
                Count: values.length
            }),
            d => d.Country
        ),
        ([Country, stats]) => ({
            Country,
            Value: stats.Value,
            Count: stats.Count
        })
    )
    .filter(d => d.Value > 0)
    .sort((a, b) => b.Value - a.Value);

    if (!countryRollup.length) {
        container.append("div")
            .style("padding", "80px")
            .style("text-align", "center")
            .text("ไม่มีข้อมูล");
        return;
    }

    let countryData;

    if (localDonutTop === 999) {
        countryData = countryRollup;
    } else {
        const top = countryRollup.slice(0, localDonutTop);
        const others = countryRollup.slice(localDonutTop);

        countryData = [...top];

        if (others.length) {
            countryData.push({
                Country: "Others",
                Value: d3.sum(others, d => d.Value),
                Count: d3.sum(others, d => d.Count)
            });
        }
    }

    const totalValue = d3.sum(
        countryData,
        d => Math.max(0, d.Value)
    );

    const svg = container.append("svg")
        .attr("width", donutWidth)
        .attr("height", height)
        .attr("viewBox", `0 0 ${donutWidth} ${height}`)
        .style("flex", `0 0 ${donutWidth}px`)
        .style("width", `${donutWidth}px`)
        .style("height", `${height}px`)
        .style("overflow", "visible");

    const radius = Math.max(
        70,
        Math.min(
            donutWidth * 0.43,
            height * 0.40
        )
    );

    const centerX = donutWidth / 2;
    const centerY = height / 2;

    const g = svg.append("g")
        .attr(
            "transform",
            `translate(${centerX},${centerY})`
        );

    const pie = d3.pie()
        .sort(null)
        .value(d => Math.max(0, d.Value));

    const arc = d3.arc()
        .innerRadius(radius * 0.55)
        .outerRadius(radius);

    const hoverArc = d3.arc()
        .innerRadius(radius * 0.52)
        .outerRadius(radius * 1.05);

    const startArc = d3.arc()
        .innerRadius(radius * 0.55)
        .outerRadius(radius * 0.55);

    const pieData = pie(countryData);

    const paths = g.selectAll(".country-slice")
        .data(pieData)
        .enter()
        .append("path")
        .attr("class", "country-slice")
        .attr("fill", d =>
            d.data.Country === "Others"
                ? "#CBD5E1"
                : getCountryColor(d.data.Country)
        )
        .attr("stroke", "#ffffff")
        .attr("stroke-width", 2)
        .attr("d", startArc)
        .style("opacity", 0.15)
        .style("cursor", d =>
            d.data.Country === "Others"
                ? "default"
                : "pointer"
        );

    paths.transition()
        .duration(900)
        .delay((d,i)=>i*70)
        .ease(d3.easeCubicOut)
        .style("opacity",1)
        .attrTween("d", function(d) {
            const interpolate = d3.interpolate(startArc(d), arc(d));
            return t => interpolate(t);
        });

    paths.on("mouseover", function (event, d) {
            d3.select(this)
                .transition()
                .duration(120)
                .attr("d", hoverArc);

            const percent =
                totalValue > 0
                    ? (d.data.Value / totalValue) * 100
                    : 0;

            showTooltip(
                event,
                `<b>${safeText(d.data.Country)}</b>
                 <br>จำนวนรายการ: ${d3.format(",")(d.data.Count)}
                 <br>มูลค่า: £${d3.format(",.2f")(d.data.Value)}
                 <br>สัดส่วน: ${percent.toFixed(1)}%`
            );
        })
        .on("mouseout", function (event, d) {
            d3.select(this)
                .transition()
                .duration(120)
                .attr("d", arc);

            hideTooltip();
        })
        .on("click", function (event, d) {
            if (d.data.Country !== "Others") {
                selectedCountry = d.data.Country;

                d3.select("#countryFilter")
                    .property("value", selectedCountry);

                applyFilters();
            }
        });

    g.append("text")
        .attr("text-anchor", "middle")
        .attr("dy", "-0.15em")
        .style("font-size", "15px")
        .style("font-weight", "500")
        .style("fill", "#64748b")
        .text("ยอดขาย");

    g.append("text")
        .attr("text-anchor", "middle")
        .attr("dy", "1.15em")
        .style("font-size", "17px")
        .style("font-weight", "700")
        .style("fill", "#40364b")
        .text(`£${d3.format(",.0f")(totalValue)}`);

    /* =====================================================
       LEGEND - FULL NAME + SCROLL
       ===================================================== */

    const legend = container.append("div")
        .attr("class", "country-legend-panel")
        .style("width", `${Math.max(legendWidth, 120)}px`)
        .style("flex", "1 1 auto");

    countryData.forEach(d => {
        const item = legend.append("div")
            .attr("class", "country-legend-item")
            .attr("title", d.Country)
            .style(
                "cursor",
                d.Country === "Others"
                    ? "default"
                    : "pointer"
            );

        item.append("span")
            .attr("class", "country-legend-dot")
            .style(
                "background",
                d.Country === "Others"
                    ? "#CBD5E1"
                    : getCountryColor(d.Country)
            );

        item.append("span")
            .attr("class", "country-legend-name")
            .text(d.Country);

        item.on("mouseenter", function (event) {
            paths
                .filter(p =>
                    p.data.Country === d.Country
                )
                .transition()
                .duration(120)
                .attr("d", hoverArc);

            const percent =
                totalValue > 0
                    ? (d.Value / totalValue) * 100
                    : 0;

            showTooltip(
                event,
                `<b>${safeText(d.Country)}</b>
                 <br>จำนวนรายการ: ${d3.format(",")(d.Count)}
                 <br>มูลค่า: £${d3.format(",.2f")(d.Value)}
                 <br>สัดส่วน: ${percent.toFixed(1)}%`
            );
        });

        item.on("mouseleave", function () {
            paths
                .filter(p =>
                    p.data.Country === d.Country
                )
                .transition()
                .duration(120)
                .attr("d", arc);

            hideTooltip();
        });

        item.on("click", function () {
            if (d.Country !== "Others") {
                selectedCountry = d.Country;

                d3.select("#countryFilter")
                    .property("value", selectedCountry);

                applyFilters();
            }
        });
    });
}

/* =========================================================
   3. COLUMN CHART
   ========================================================= */

function renderColumnChart() {
    const container = d3.select("#columnChart");
    if (container.empty()) return;

    container.html("");

    const node = container.node();
    const bounds = node.getBoundingClientRect();

    const margin = {
        top: 25,
        right: 25,
        bottom: 70,
        left: 80
    };

    const width = Math.max(
        200,
        bounds.width - margin.left - margin.right
    );

    const height = Math.max(
        180,
        bounds.height - margin.top - margin.bottom
    );

    const dataset = globalDataset.filter(d => {
        const countryOK =
            selectedCountry === "ALL" ||
            d.Country === selectedCountry;

        const yearOK =
            selectedYear === "ALL" ||
            String(d.Year) === String(selectedYear);

        const monthOK =
            selectedMonth === "ALL" ||
            String(d.Month) === String(selectedMonth);

        return countryOK && yearOK && monthOK;
    });

    const typeData = [
        {
            Type: "Sale",
            Sales: d3.sum(
                dataset.filter(
                    d => d.TransactionType === "Sale"
                ),
                d => Number(d.LineAmount) || 0
            ),
            Count: dataset.filter(
                d => d.TransactionType === "Sale"
            ).length
        },
        {
            Type: "Return/Cancelled",
            Sales: d3.sum(
                dataset.filter(
                    d =>
                        d.TransactionType ===
                        "Return/Cancelled"
                ),
                d => Number(d.LineAmount) || 0
            ),
            Count: dataset.filter(
                d =>
                    d.TransactionType ===
                    "Return/Cancelled"
            ).length
        }
    ];

    const svg = container.append("svg")
        .attr("width", "100%")
        .attr("height", "100%")
        .attr("viewBox", `0 0 ${bounds.width} ${bounds.height}`)
        .style("overflow", "visible");

    const g = svg.append("g")
        .attr(
            "transform",
            `translate(${margin.left},${margin.top})`
        );

    const x = d3.scaleBand()
        .domain(typeData.map(d => d.Type))
        .range([0, width])
        .padding(0.45);

    const maxValue =
        d3.max(typeData, d => d.Sales) || 1;

    const y = d3.scaleLinear()
        .domain([0, maxValue * 1.15])
        .nice()
        .range([height, 0]);

    g.append("g")
        .attr("transform", `translate(0,${height})`)
        .call(d3.axisBottom(x));

    g.append("g")
        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickFormat(
                    d =>
                        "£" +
                        d3.format(".0f")(d / 1000) +
                        "k"
                )
        );

    addAxisLabel(g, {
        x: width / 2,
        y: height + 52,
        text: "ประเภทรายการ"
    });

    addAxisLabel(g, {
        x: -height / 2,
        y: -margin.left + 22,
        text: "มูลค่ารวม (£)",
        rotate: -90
    });

    g.selectAll(".column")
        .data(typeData)
        .enter()
        .append("rect")
        .attr("class", "column")
        .attr("x", d => x(d.Type))
        .attr("width", x.bandwidth())
        .attr("y", height)
        .attr("height", 0)
        .attr("rx", 6)
        .attr(
            "fill",
            d => TYPE_COLORS[d.Type]
        )
        .transition().duration(850).delay((d,i)=>i*180).ease(d3.easeCubicOut)
        .attr("y", d => y(d.Sales))
        .attr("height", d => height - y(d.Sales))
        .on("mouseover", (event, d) => {
            showTooltip(
                event,
                `<b>${safeText(d.Type)}</b>
                 <br>มูลค่ารายการ: £${d3.format(",.2f")(d.Sales)}
                 <br>จำนวนรายการ: ${d3.format(",")(d.Count)}`
            );
        })
        .on("mouseout", hideTooltip);
}

/* =========================================================
   4. SCATTER CHART
   FIXED ZOOM + CLIPPING + AXIS LABELS
   ========================================================= */

function renderScatterChart() {
    const container = d3.select("#scatterChart");
    if (container.empty()) return;

    container.html("");

    const node = container.node();
    const bounds = node.getBoundingClientRect();

    const margin = {
        top: 35,
        right: 30,
        bottom: 72,
        left: 75
    };

    const outerWidth = Math.max(
        300,
        bounds.width || 700
    );

    const outerHeight = Math.max(
        280,
        bounds.height || 350
    );

    const width = Math.max(
        200,
        outerWidth - margin.left - margin.right
    );

    const height = Math.max(
        180,
        outerHeight - margin.top - margin.bottom
    );

    let dataset = filteredData.filter(d =>
        Number.isFinite(d.UnitPrice) &&
        Number.isFinite(d.Quantity) &&
        d.UnitPrice >= 0 &&
        d.Quantity >= 0
    );

    if (localScatterPrice !== "all") {
        dataset = dataset.filter(
            d =>
                d.UnitPrice <=
                Number(localScatterPrice)
        );
    }

    if (localScatterQty > 0) {
        dataset = dataset.filter(
            d =>
                d.Quantity >=
                localScatterQty
        );
    }

    let sampleData = dataset;

    if (dataset.length > 500) {
        sampleData = d3.range(500).map(i => {
            const index = Math.floor(
                i *
                (dataset.length - 1) /
                499
            );
            return dataset[index];
        });
    }

    const svg = container.append("svg")
        .attr("width", "100%")
        .attr("height", "100%")
        .attr("viewBox", `0 0 ${outerWidth} ${outerHeight}`)
        .style("overflow", "hidden");

    const defs = svg.append("defs");

    defs.append("clipPath")
        .attr("id", "scatterPlotClip")
        .append("rect")
        .attr("x", 0)
        .attr("y", 0)
        .attr("width", width)
        .attr("height", height);

    const g = svg.append("g")
        .attr(
            "transform",
            `translate(${margin.left},${margin.top})`
        );

    const plot = g.append("g")
        .attr("clip-path", "url(#scatterPlotClip)");

    const xMax =
        d3.max(sampleData, d => d.UnitPrice) || 10;

    const yMax =
        d3.max(sampleData, d => d.Quantity) || 10;

    const x = d3.scaleLinear()
        .domain([0, Math.max(1, xMax * 1.05)])
        .nice()
        .range([0, width]);

    const y = d3.scaleLinear()
        .domain([0, Math.max(1, yMax * 1.05)])
        .nice()
        .range([height, 0]);

    const xAxis = g.append("g")
        .attr(
            "transform",
            `translate(0,${height})`
        )
        .call(
            d3.axisBottom(x)
                .ticks(6)
        );

    const yAxis = g.append("g")
        .call(
            d3.axisLeft(y)
                .ticks(6)
        );

    addAxisLabel(g, {
        x: width / 2,
        y: height + 52,
        text: "ราคาต่อหน่วย (£)"
    });

    addAxisLabel(g, {
        x: -height / 2,
        y: -margin.left + 20,
        text: "ปริมาณสั่งซื้อ (ชิ้น)",
        rotate: -90
    });

    const dots = plot.selectAll(".scatter-point")
        .data(sampleData)
        .enter()
        .append("circle")
        .attr("class", "scatter-point")
        .attr("cx", d => x(d.UnitPrice))
        .attr("cy", d => y(d.Quantity))
        .attr("r", 0)
        .attr(
            "fill",
            d =>
                TYPE_COLORS[d.TransactionType] ||
                TYPE_COLORS.Sale
        )
        .attr("opacity", 0)
        .transition().duration(650).delay((d,i)=>Math.min(i*2,900)).ease(d3.easeCubicOut)
        .attr("r",4)
        .attr("opacity",0.68)
        .on("mouseover", (event, d) => {
            showTooltip(
                event,
                `<b>${safeText(d.Description)}</b>
                 <br>ประเภท: ${safeText(d.TransactionType)}
                 <br>ราคา: £${d3.format(",.2f")(d.UnitPrice)}
                 <br>จำนวน: ${d3.format(",")(d.Quantity)} ชิ้น`
            );
        })
        .on("mouseout", hideTooltip);

    /*
       IMPORTANT:
       Zoom is applied to the plot area only.
       The axes stay in their own position.
       The clipPath prevents points from escaping
       outside the graph.
    */

    const zoom = d3.zoom()
        .scaleExtent([1, 5])
        .extent([
            [0, 0],
            [width, height]
        ])
        .translateExtent([
            [0, 0],
            [width, height]
        ])
        .on("zoom", event => {
            const transform = event.transform;

            const newX = transform.rescaleX(x);
            const newY = transform.rescaleY(y);

            xAxis.call(
                d3.axisBottom(newX)
                    .ticks(6)
            );

            yAxis.call(
                d3.axisLeft(newY)
                    .ticks(6)
            );

            dots
                .attr(
                    "cx",
                    d => newX(d.UnitPrice)
                )
                .attr(
                    "cy",
                    d => newY(d.Quantity)
                );
        });

    /*
       Invisible zoom layer only inside plot area.
       This prevents the whole SVG / labels from moving.
    */

    const zoomLayer = plot.append("rect")
        .attr("class", "zoom-layer")
        .attr("width", width)
        .attr("height", height)
        .style("fill", "none")
        .style("pointer-events", "all");

    zoomLayer.call(zoom);

    /*
       Keep dots above zoom layer visually.
       Pointer events still work on dots because the
       zoom layer is inserted before the dots.
    */

    dots.raise();

    d3.select("#resetZoomBtn")
        .on("click", () => {
            zoomLayer
                .transition()
                .duration(400)
                .call(
                    zoom.transform,
                    d3.zoomIdentity
                );
        });
}
