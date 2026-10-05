/* =========================================================
   ONLINE RETAIL ANALYTICS DASHBOARD
   XLSB VERSION
   ========================================================= */

const DATA_URL = "Online_Retail_Cleaned_Final-1.xlsb";

let globalDataset = [];
let filteredData = [];


/* =========================================================
   GLOBAL FILTER STATE
   ========================================================= */

let selectedCountry = "ALL";
let selectedType = "ALL";
let selectedYear = "ALL";
let selectedMonth = "ALL";

let localProductMetric = "value";
let localProductTop = 10;

let localDonutTop = 5;

let localScatterPrice = "all";
let localScatterQty = 0;


/* =========================================================
   COUNTRY COLORS
   ========================================================= */

const COUNTRY_COLOR_MAP = {};

const COUNTRY_HUE_STEP = 137.508;

function getCountryColor(country) {

    if (!COUNTRY_COLOR_MAP[country]) {

        const index =
            Object.keys(COUNTRY_COLOR_MAP).length;

        const hue =
            (index * COUNTRY_HUE_STEP) % 360;

        COUNTRY_COLOR_MAP[country] =
            `hsl(${hue.toFixed(1)}, 68%, 48%)`;
    }

    return COUNTRY_COLOR_MAP[country];
}


/* =========================================================
   OTHER COLORS
   ========================================================= */

const TYPE_COLORS = {

    "Sale": "#10b981",

    "Return/Cancelled": "#ef4444"

};


const BAR_COLORS = [

    "#3b82f6",
    "#6366f1",
    "#8b5cf6",
    "#a855f7",
    "#d946ef",
    "#ec4899",
    "#f43f5e",
    "#f97316",
    "#eab308",
    "#10b981"

];


/* =========================================================
   MONTH NAMES
   ========================================================= */

const MONTH_NAMES_TH = [

    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม"

];


/* =========================================================
   START
   ========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function () {

        setupEvents();

        loadData();

    }
);


/* =========================================================
   EVENTS
   ========================================================= */

function setupEvents() {


    d3.select("#countryFilter")
        .on("change", function () {

            selectedCountry =
                this.value;

            applyFilters();

        });


    d3.select("#typeFilter")
        .on("change", function () {

            selectedType =
                this.value;

            applyFilters();

        });


    d3.select("#yearFilter")
        .on("change", function () {

            selectedYear =
                this.value;

            selectedMonth =
                "ALL";

            populateMonthDropdown();

            d3.select("#monthFilter")
                .property(
                    "value",
                    "ALL"
                );

            applyFilters();

        });


    d3.select("#monthFilter")
        .on("change", function () {

            selectedMonth =
                this.value;

            applyFilters();

        });


    d3.select("#localProductMetric")
        .on("change", function () {

            localProductMetric =
                this.value;

            renderBarChart();

        });


    d3.select("#localProductTop")
        .on("change", function () {

            localProductTop =
                Number(this.value);

            renderBarChart();

        });


    d3.select("#localDonutTop")
        .on("change", function () {

            localDonutTop =
                Number(this.value);

            renderDonutChart();

        });


    d3.select("#localScatterPrice")
        .on("change", function () {

            localScatterPrice =
                this.value;

            renderScatterChart();

        });


    d3.select("#localScatterQty")
        .on("change", function () {

            localScatterQty =
                Number(this.value);

            renderScatterChart();

        });


    d3.select("#resetBtn")
        .on("click", resetDashboard);


    window.addEventListener(
        "resize",
        function () {

            if (globalDataset.length) {

                renderCharts();

            }

        }
    );

}


/* =========================================================
   LOAD XLSB
   ========================================================= */

async function loadData() {

    try {

        console.log(
            "กำลังโหลด:",
            DATA_URL
        );


        if (
            typeof XLSX ===
            "undefined"
        ) {

            throw new Error(
                "ไม่พบ SheetJS (XLSX)"
            );

        }


        const response =
            await fetch(
                DATA_URL +
                "?v=" +
                Date.now(),
                {
                    cache: "no-store"
                }
            );


        if (!response.ok) {

            throw new Error(
                "โหลดไฟล์ XLSB ไม่ได้: HTTP " +
                response.status
            );

        }


        const buffer =
            await response.arrayBuffer();


        if (
            !buffer ||
            buffer.byteLength === 0
        ) {

            throw new Error(
                "ไฟล์ XLSB ว่าง"
            );

        }


        console.log(
            "XLSB size:",
            buffer.byteLength,
            "bytes"
        );


        const workbook =
            XLSX.read(
                buffer,
                {
                    type: "array",
                    cellDates: true,
                    dense: true
                }
            );


        if (
            !workbook.SheetNames ||
            !workbook.SheetNames.length
        ) {

            throw new Error(
                "ไม่พบ Sheet ในไฟล์ XLSB"
            );

        }


        const sheetName =
            workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[
                sheetName
            ];


        if (!worksheet) {

            throw new Error(
                "ไม่พบข้อมูลใน Sheet"
            );

        }


        console.log(
            "Sheet:",
            sheetName
        );


        const rows =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: "",
                    raw: true
                }
            );


        if (!rows.length) {

            throw new Error(
                "Sheet ไม่มีข้อมูล"
            );

        }


        console.log(
            "Rows:",
            rows.length
        );


        console.log(
            "Columns:",
            Object.keys(rows[0])
        );


        convertData(rows);


    }

    catch (error) {

        console.error(
            "XLSB ERROR:",
            error
        );


        showLoadError(
            error
        );

    }

}


/* =========================================================
   CONVERT DATA
   ========================================================= */

function convertData(rows) {


    globalDataset =
        rows.map(
            function (d) {


                const quantity =
                    Number(
                        d.Quantity
                    ) || 0;


                const unitPrice =
                    Number(
                        d.UnitPrice
                    ) || 0;


                let lineAmount =
                    Number(
                        d.LineAmount
                    );


                if (
                    !Number.isFinite(
                        lineAmount
                    )
                ) {

                    lineAmount =
                        quantity *
                        unitPrice;

                }


                const date =
                    parseInvoiceDate(
                        d.InvoiceDate
                    );


                let transactionType =
                    String(
                        d.TransactionType ||
                        ""
                    ).trim();


                if (!transactionType) {

                    transactionType =
                        quantity < 0
                            ? "Return/Cancelled"
                            : "Sale";

                }


                const typeLower =
                    transactionType
                        .toLowerCase();


                if (
                    typeLower.includes(
                        "return"
                    ) ||
                    typeLower.includes(
                        "cancel"
                    )
                ) {

                    transactionType =
                        "Return/Cancelled";

                }

                else {

                    transactionType =
                        "Sale";

                }


                return {

                    Description:
                        d.Description ||
                        "Uncategorized",

                    Quantity:
                        quantity,

                    UnitPrice:
                        unitPrice,

                    CustomerID:
                        d.CustomerID,

                    Country:
                        d.Country ||
                        "Unknown",

                    TransactionType:
                        transactionType,

                    LineAmount:
                        lineAmount,

                    InvoiceDate:
                        d.InvoiceDate,

                    Date:
                        date,

                    Year:
                        date
                            ? date.getFullYear()
                            : null,

                    Month:
                        date
                            ? date.getMonth() + 1
                            : null

                };

            }
        )
        .filter(
            function (d) {

                return (
                    d.Date instanceof Date &&
                    !Number.isNaN(
                        d.Date.getTime()
                    )
                );

            }
        );


    if (!globalDataset.length) {

        throw new Error(
            "อ่าน XLSB ได้ แต่ไม่พบข้อมูลวันที่ที่ถูกต้อง"
        );

    }


    filteredData =
        [
            ...globalDataset
        ];


    populateCountryDropdown();

    populateYearDropdown();

    populateMonthDropdown();

    updateDashboard();


    console.log(
        "================================="
    );

    console.log(
        "XLSB LOADED SUCCESSFULLY"
    );

    console.log(
        "จำนวนข้อมูล:",
        globalDataset.length
    );

    console.log(
        "ประเทศ:",
        new Set(
            globalDataset.map(
                d => d.Country
            )
        ).size
    );

    console.log(
        "================================="
    );

}


/* =========================================================
   PARSE DATE
   ========================================================= */

function parseInvoiceDate(value) {


    if (
        value instanceof Date &&
        !Number.isNaN(
            value.getTime()
        )
    ) {

        return value;

    }


    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    /*
       Excel serial date
       เช่น 40513.35138888889
    */

    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        const excelEpoch =
            new Date(
                Date.UTC(
                    1899,
                    11,
                    30
                )
            );


        const date =
            new Date(
                excelEpoch.getTime() +
                value *
                86400000
            );


        return (
            Number.isNaN(
                date.getTime()
            )
                ? null
                : date
        );

    }


    const raw =
        String(value).trim();


    /*
       Try native date
    */

    const nativeDate =
        new Date(raw);


    if (
        !Number.isNaN(
            nativeDate.getTime()
        )
    ) {

        return nativeDate;

    }


    /*
       MM/DD/YYYY AM PM
    */

    let date =
        d3.timeParse(
            "%m/%d/%Y %I:%M:%S %p"
        )(raw);


    if (date) return date;


    date =
        d3.timeParse(
            "%m/%d/%Y %I:%M %p"
        )(raw);


    if (date) return date;


    /*
       24-hour
    */

    date =
        d3.timeParse(
            "%m/%d/%Y %H:%M:%S"
        )(raw);


    if (date) return date;


    return d3.timeParse(
        "%m/%d/%Y %H:%M"
    )(raw);

}


/* =========================================================
   ERROR MESSAGE
   ========================================================= */

function showLoadError(error) {


    const message =
        error &&
        error.message
            ? error.message
            : String(error);


    const targets = [

        "#barChart",
        "#donutChart",
        "#columnChart",
        "#scatterChart"

    ];


    targets.forEach(
        function (selector) {

            const box =
                d3.select(
                    selector
                );


            box.html("");


            box.append("div")
                .style(
                    "padding",
                    "30px"
                )
                .style(
                    "text-align",
                    "center"
                )
                .style(
                    "color",
                    "#dc2626"
                )
                .style(
                    "font-weight",
                    "600"
                )
                .text(
                    "โหลดข้อมูลไม่สำเร็จ: " +
                    message
                );

        }
    );

}


/* =========================================================
   COUNTRY DROPDOWN
   ========================================================= */

function populateCountryDropdown() {


    const select =
        d3.select(
            "#countryFilter"
        );


    select
        .selectAll(
            "option:not(:first-child)"
        )
        .remove();


    const countries =
        Array.from(
            new Set(
                globalDataset.map(
                    d => d.Country
                )
            )
        )
        .filter(Boolean)
        .sort();


    countries.forEach(
        function (country) {

            select
                .append("option")
                .attr(
                    "value",
                    country
                )
                .text(
                    country
                );

        }
    );

}


/* =========================================================
   YEAR DROPDOWN
   ========================================================= */

function populateYearDropdown() {


    const select =
        d3.select(
            "#yearFilter"
        );


    select
        .selectAll(
            "option:not(:first-child)"
        )
        .remove();


    const years =
        Array.from(
            new Set(
                globalDataset
                    .map(
                        d => d.Year
                    )
                    .filter(Boolean)
            )
        )
        .sort(
            (a, b) =>
                a - b
        );


    years.forEach(
        function (year) {

            select
                .append("option")
                .attr(
                    "value",
                    year
                )
                .text(
                    year
                );

        }
    );

}


/* =========================================================
   MONTH DROPDOWN
   ========================================================= */

function populateMonthDropdown() {


    const select =
        d3.select(
            "#monthFilter"
        );


    select
        .selectAll(
            "option:not(:first-child)"
        )
        .remove();


    const months =
        Array.from(
            new Set(

                globalDataset

                    .filter(
                        function (d) {

                            return (
                                selectedYear ===
                                "ALL" ||

                                String(
                                    d.Year
                                ) ===
                                String(
                                    selectedYear
                                )
                            );

                        }
                    )

                    .map(
                        d => d.Month
                    )

                    .filter(Boolean)

            )
        )
        .sort(
            (a, b) =>
                a - b
        );


    months.forEach(
        function (month) {

            select
                .append("option")
                .attr(
                    "value",
                    month
                )
                .text(

                    String(
                        month
                    ).padStart(
                        2,
                        "0"
                    ) +

                    " - " +

                    MONTH_NAMES_TH[
                        month - 1
                    ]

                );

        }
    );


    if (
        selectedMonth !== "ALL" &&
        !months.includes(
            Number(
                selectedMonth
            )
        )
    ) {

        selectedMonth =
            "ALL";

    }


    select.property(
        "value",
        selectedMonth
    );

}


/* =========================================================
   APPLY FILTERS
   ========================================================= */

function applyFilters() {


    filteredData =
        globalDataset.filter(
            function (d) {


                const matchCountry =
                    selectedCountry ===
                    "ALL" ||
                    d.Country ===
                    selectedCountry;


                const matchType =
                    selectedType ===
                    "ALL" ||
                    d.TransactionType ===
                    selectedType;


                const matchYear =
                    selectedYear ===
                    "ALL" ||
                    String(d.Year) ===
                    String(selectedYear);


                const matchMonth =
                    selectedMonth ===
                    "ALL" ||
                    String(d.Month) ===
                    String(selectedMonth);


                return (
                    matchCountry &&
                    matchType &&
                    matchYear &&
                    matchMonth
                );

            }
        );


    updateDashboard();

}


/* =========================================================
   RESET
   ========================================================= */

function resetDashboard() {


    selectedCountry =
        "ALL";

    selectedType =
        "ALL";

    selectedYear =
        "ALL";

    selectedMonth =
        "ALL";


    localProductMetric =
        "value";

    localProductTop =
        10;

    localDonutTop =
        5;

    localScatterPrice =
        "all";

    localScatterQty =
        0;


    d3.select(
        "#countryFilter"
    )
    .property(
        "value",
        "ALL"
    );


    d3.select(
        "#typeFilter"
    )
    .property(
        "value",
        "ALL"
    );


    d3.select(
        "#yearFilter"
    )
    .property(
        "value",
        "ALL"
    );


    d3.select(
        "#localProductMetric"
    )
    .property(
        "value",
        "value"
    );


    d3.select(
        "#localProductTop"
    )
    .property(
        "value",
        "10"
    );


    d3.select(
        "#localDonutTop"
    )
    .property(
        "value",
        "5"
    );


    d3.select(
        "#localScatterPrice"
    )
    .property(
        "value",
        "all"
    );


    d3.select(
        "#localScatterQty"
    )
    .property(
        "value",
        "0"
    );


    populateMonthDropdown();


    d3.select(
        "#monthFilter"
    )
    .property(
        "value",
        "ALL"
    );


    filteredData =
        [
            ...globalDataset
        ];


    updateDashboard();

}


/* =========================================================
   UPDATE DASHBOARD
   ========================================================= */

function updateDashboard() {

    renderKPIs();

    renderCharts();

}


/* =========================================================
   KPI
   ========================================================= */

function renderKPIs() {


    const totalSales =
        d3.sum(
            filteredData,
            d =>
                Number(
                    d.LineAmount
                ) || 0
        );


    const totalQty =
        d3.sum(
            filteredData,
            d =>
                Number(
                    d.Quantity
                ) || 0
        );


    const totalOrders =
        filteredData.length;


    const totalCountries =
        new Set(
            filteredData.map(
                d => d.Country
            )
        ).size;


    d3.select(
        "#kpiTotalSales"
    )
    .text(
        `£${d3.format(
            ",.2f"
        )(totalSales)}`
    );


    d3.select(
        "#kpiTotalQty"
    )
    .text(
        `${d3.format(
            ","
        )(totalQty)} ชิ้น`
    );


    d3.select(
        "#kpiTotalOrders"
    )
    .text(
        `${d3.format(
            ","
        )(totalOrders)} รายการ`
    );


    d3.select(
        "#kpiTotalCountries"
    )
    .text(
        `${totalCountries} ประเทศ`
    );

}


/* =========================================================
   CHARTS
   ========================================================= */

function renderCharts() {

    renderBarChart();

    renderDonutChart();

    renderColumnChart();

    renderScatterChart();

}


/* =========================================================
   TOOLTIP
   ========================================================= */

const tooltip =
    d3.select(
        "#tooltip"
    );


function showTooltip(
    event,
    content
) {

    tooltip
        .html(content)
        .style(
            "opacity",
            1
        )
        .style(
            "left",
            event.pageX +
            15 +
            "px"
        )
        .style(
            "top",
            event.pageY -
            28 +
            "px"
        );

}


function hideTooltip() {

    tooltip.style(
        "opacity",
        0
    );

}


/* =========================================================
   1. BAR CHART
   ========================================================= */

function renderBarChart() {


    const container =
        d3.select(
            "#barChart"
        );


    container.html("");


    const node =
        container.node();


    if (!node) return;


    const bounds =
        node.getBoundingClientRect();


    const margin = {

        top: 20,
        right: 30,
        bottom: 55,
        left: 190

    };


    const width =
        Math.max(
            200,
            bounds.width -
            margin.left -
            margin.right
        );


    const height =
        Math.max(
            180,
            bounds.height -
            margin.top -
            margin.bottom
        );


    const productData =
        Array.from(

            d3.rollup(

                filteredData,

                function (values) {

                    return {

                        Sales:
                            d3.sum(
                                values,
                                d =>
                                    Number(
                                        d.LineAmount
                                    ) || 0
                            ),

                        Qty:
                            d3.sum(
                                values,
                                d =>
                                    Number(
                                        d.Quantity
                                    ) || 0
                            )

                    };

                },

                d =>
                    d.Description

            ),

            function (entry) {

                const Description =
                    entry[0];

                const stats =
                    entry[1];

                return {

                    FullDesc:
                        Description,

                    ShortDesc:
                        Description.length >
                        22
                            ? Description.substring(
                                0,
                                20
                            ) +
                            "..."
                            : Description,

                    Sales:
                        stats.Sales,

                    Qty:
                        stats.Qty

                };

            }

        )
        .sort(

            function (a, b) {

                return (
                    localProductMetric ===
                    "qty"

                        ? b.Qty -
                          a.Qty

                        : b.Sales -
                          a.Sales
                );

            }

        )
        .slice(
            0,
            localProductTop
        );


    if (!productData.length) {

        container
            .append("div")
            .style(
                "padding",
                "80px"
            )
            .style(
                "text-align",
                "center"
            )
            .text(
                "ไม่มีข้อมูล"
            );

        return;

    }


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                "100%"
            )
            .attr(
                "height",
                "100%"
            )
            .attr(
                "viewBox",
                `0 0 ${bounds.width} ${bounds.height}`
            );


    const g =
        svg
            .append("g")
            .attr(
                "transform",
                `translate(${margin.left},${margin.top})`
            );


    const y =
        d3.scaleBand()
            .domain(
                productData.map(
                    d => d.ShortDesc
                )
            )
            .range(
                [0, height]
            )
            .padding(
                0.25
            );


    const metricMax =
        d3.max(
            productData,
            d =>
                localProductMetric ===
                "qty"
                    ? d.Qty
                    : d.Sales
        ) || 1;


    const x =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    metricMax * 1.1
                ]
            )
            .range(
                [0, width]
            );


    g.append("g")
        .call(
            d3.axisLeft(y)
        );


    g.append("g")
        .attr(
            "transform",
            `translate(0,${height})`
        )
        .call(
            d3.axisBottom(x)
                .ticks(5)
                .tickFormat(
                    function (d) {

                        return (
                            localProductMetric ===
                            "qty"

                                ? d3.format(",")(d)

                                : "£" +
                                  d3.format(
                                      ".0f"
                                  )(d / 1000) +
                                  "k"
                        );

                    }
                )
        );


    g.selectAll(
        ".bar"
    )
    .data(
        productData
    )
    .enter()
    .append("rect")
    .attr(
        "class",
        "bar"
    )
    .attr(
        "x",
        0
    )
    .attr(
        "y",
        d =>
            y(
                d.ShortDesc
            )
    )
    .attr(
        "height",
        y.bandwidth()
    )
    .attr(
        "rx",
        4
    )
    .attr(
        "fill",
        (d, i) =>
            BAR_COLORS[
                i %
                BAR_COLORS.length
            ]
    )
    .attr(
        "width",
        d =>
            x(
                localProductMetric ===
                "qty"
                    ? d.Qty
                    : d.Sales
            )
    )
    .on(
        "mouseover",
        function (
            event,
            d
        ) {

            showTooltip(

                event,

                `<b>${d.FullDesc}</b>
                 <br>ยอดขาย: £${d3.format(",.2f")(d.Sales)}
                 <br>จำนวน: ${d3.format(",")(d.Qty)} ชิ้น`

            );

        }
    )
    .on(
        "mouseout",
        hideTooltip
    );


}


/* =========================================================
   2. DONUT CHART
   FIXED FULL CIRCLE
   ========================================================= */

function renderDonutChart() {


    const container =
        d3.select(
            "#donutChart"
        );


    if (container.empty()) {
        return;
    }


    container.html("");


    const node =
        container.node();


    const bounds =
        node.getBoundingClientRect();


    /*
       ใช้พื้นที่ทั้งหมดของ chart-wrapper
       ไม่แบ่ง width เป็น 62%
       จึงไม่ตัดวงกลม
    */

    const width =
        Math.max(
            300,
            bounds.width ||
            600
        );


    /*
       กำหนดความสูงขั้นต่ำ
       เพื่อให้วงกลมเต็ม
    */

    const height =
        Math.max(
            360,
            bounds.height ||
            420
        );


    /* =====================================================
       GROUP COUNTRY DATA
       ===================================================== */

    const countryRollup =
        Array.from(

            d3.rollup(

                filteredData,

                function (values) {

                    return {

                        Value:
                            d3.sum(
                                values,
                                d =>
                                    Number(
                                        d.LineAmount
                                    ) || 0
                            ),

                        Count:
                            values.length

                    };

                },

                d =>
                    d.Country ||
                    "Unknown"

            ),

            function (entry) {

                return {

                    Country:
                        entry[0],

                    Value:
                        entry[1].Value,

                    Count:
                        entry[1].Count

                };

            }

        )
        .sort(
            (a, b) =>
                b.Value -
                a.Value
        );


    if (!countryRollup.length) {

        container
            .append("div")
            .style(
                "padding",
                "100px 20px"
            )
            .style(
                "text-align",
                "center"
            )
            .style(
                "color",
                "#64748b"
            )
            .text(
                "ไม่มีข้อมูลสำหรับกราฟ"
            );

        return;

    }


    const totalValue =
        d3.sum(
            countryRollup,
            d =>
                d.Value
        );


    /* =====================================================
       TOP + OTHERS
       ===================================================== */

    let countryData = [];


    if (
        localDonutTop ===
        999
    ) {

        countryData =
            [
                ...countryRollup
            ];

    }

    else {

        const topCount =
            Math.max(
                1,
                localDonutTop
            );


        countryData =
            countryRollup.slice(
                0,
                topCount
            );


        const others =
            countryRollup.slice(
                topCount
            );


        if (
            others.length
        ) {

            const othersValue =
                d3.sum(
                    others,
                    d =>
                        d.Value
                );


            const othersCount =
                d3.sum(
                    others,
                    d =>
                        d.Count
                );


            if (
                othersValue >
                0
            ) {

                countryData.push({

                    Country:
                        "Others",

                    Value:
                        othersValue,

                    Count:
                        othersCount

                });

            }

        }

    }


    countryData.forEach(
        function (d) {

            if (
                d.Country !==
                "Others"
            ) {

                getCountryColor(
                    d.Country
                );

            }

        }
    );


    /* =====================================================
       SVG FULL WIDTH
       ===================================================== */

    const svg =
        container
            .append("svg")
            .attr(
                "width",
                "100%"
            )
            .attr(
                "height",
                height
            )
            .attr(
                "viewBox",
                `0 0 ${width} ${height}`
            )
            .attr(
                "preserveAspectRatio",
                "xMidYMid meet"
            )
            .style(
                "display",
                "block"
            )
            .style(
                "overflow",
                "visible"
            );


    /*
       IMPORTANT

       radius ต้องคำนวณจากทั้ง width และ height
       และเว้นขอบทั้ง 4 ด้าน
    */

    const radius =
        Math.max(

            70,

            Math.min(

                width * 0.32,

                height * 0.40

            )

        );


    /* =====================================================
       CENTER OF CIRCLE
       ===================================================== */

    const centerX =
        width / 2;


    const centerY =
        height / 2;


    const g =
        svg
            .append("g")
            .attr(
                "transform",
                `translate(${centerX},${centerY})`
            );


    /* =====================================================
       PIE
       ===================================================== */

    const pie =
        d3.pie()
            .sort(null)
            .value(
                d =>
                    Math.max(
                        0,
                        d.Value
                    )
            );


    /* =====================================================
       DONUT ARC
       ===================================================== */

    const arc =
        d3.arc()
            .innerRadius(
                radius * 0.55
            )
            .outerRadius(
                radius
            );


    const hoverArc =
        d3.arc()
            .innerRadius(
                radius * 0.52
            )
            .outerRadius(
                radius * 1.04
            );


    /* =====================================================
       DRAW SLICES
       ===================================================== */

    const paths =
        g
            .selectAll(
                ".country-slice"
            )
            .data(
                pie(countryData)
            )
            .enter()
            .append("path")
            .attr(
                "class",
                "country-slice"
            )
            .attr(
                "fill",
                function (d) {

                    if (
                        d.data.Country ===
                        "Others"
                    ) {

                        return "#CBD5E1";

                    }

                    return getCountryColor(
                        d.data.Country
                    );

                }
            )
            .attr(
                "stroke",
                "#ffffff"
            )
            .attr(
                "stroke-width",
                2
            )
            .style(
                "cursor",
                function (d) {

                    return d.data.Country ===
                        "Others"

                        ? "default"

                        : "pointer";

                }
            )
            .each(
                function (d) {

                    this._current = {

                        startAngle:
                            d.startAngle,

                        endAngle:
                            d.startAngle

                    };

                }
            );


    /* =====================================================
       ANIMATION
       ===================================================== */

    paths
        .transition()
        .duration(700)
        .delay(
            (d, i) =>
                i * 45
        )
        .ease(
            d3.easeCubicOut
        )
        .attrTween(
            "d",
            function (d) {

                const interpolate =
                    d3.interpolate(
                        this._current,
                        d
                    );


                this._current =
                    interpolate(1);


                return function (t) {

                    return arc(
                        interpolate(t)
                    );

                };

            }
        );


    /* =====================================================
       TOOLTIP
       ===================================================== */

    paths

        .on(
            "mouseover",
            function (
                event,
                d
            ) {

                const percent =
                    totalValue > 0

                        ? (
                            d.data.Value /
                            totalValue *
                            100
                        )

                        : 0;


                d3.select(
                    this
                )
                .transition()
                .duration(150)
                .attr(
                    "d",
                    hoverArc(d)
                );


                showTooltip(

                    event,

                    `<b>${d.data.Country}</b>
                     <br>จำนวนรายการ: ${d3.format(",")(d.data.Count)}
                     <br>มูลค่า: £${d3.format(",.2f")(d.data.Value)}
                     <br>สัดส่วน: ${percent.toFixed(1)}%`

                );

            }
        )


        .on(
            "mouseout",
            function (
                event,
                d
            ) {

                d3.select(
                    this
                )
                .transition()
                .duration(150)
                .attr(
                    "d",
                    arc(d)
                );


                hideTooltip();

            }
        )


        .on(
            "click",
            function (
                event,
                d
            ) {

                if (
                    d.data.Country !==
                    "Others"
                ) {

                    selectedCountry =
                        d.data.Country;


                    d3.select(
                        "#countryFilter"
                    )
                    .property(
                        "value",
                        selectedCountry
                    );


                    applyFilters();

                }

            }
        );


    /* =====================================================
       CENTER TEXT
       ===================================================== */

    g.append("text")
        .attr(
            "text-anchor",
            "middle"
        )
        .attr(
            "dy",
            "-0.15em"
        )
        .style(
            "font-size",
            "15px"
        )
        .style(
            "font-weight",
            "500"
        )
        .style(
            "fill",
            "#64748b"
        )
        .style(
            "pointer-events",
            "none"
        )
        .text(
            "ยอดขาย"
        );


    g.append("text")
        .attr(
            "text-anchor",
            "middle"
        )
        .attr(
            "dy",
            "1.15em"
        )
        .style(
            "font-size",
            "18px"
        )
        .style(
            "font-weight",
            "700"
        )
        .style(
            "fill",
            "#0f172a"
        )
        .style(
            "pointer-events",
            "none"
        )
        .text(
            `£${d3.format(
                ",.0f"
            )(totalValue)}`
        );


    /* =====================================================
       LEGEND
       ===================================================== */

    const legend =
        container
            .append("div")
            .style(
                "display",
                "flex"
            )
            .style(
                "flex-wrap",
                "wrap"
            )
            .style(
                "justify-content",
                "center"
            )
            .style(
                "gap",
                "8px 16px"
            )
            .style(
                "padding",
                "5px 20px 15px"
            );


    countryData.forEach(
        function (d) {


            const item =
                legend
                    .append("div")
                    .style(
                        "display",
                        "flex"
                    )
                    .style(
                        "align-items",
                        "center"
                    )
                    .style(
                        "gap",
                        "6px"
                    )
                    .style(
                        "font-size",
                        "12px"
                    )
                    .style(
                        "cursor",
                        d.Country ===
                            "Others"

                            ? "default"

                            : "pointer"
                    );


            item.append("span")
                .style(
                    "width",
                    "11px"
                )
                .style(
                    "height",
                    "11px"
                )
                .style(
                    "border-radius",
                    "50%"
                )
                .style(
                    "display",
                    "inline-block"
                )
                .style(
                    "background",
                    d.Country ===
                        "Others"

                        ? "#CBD5E1"

                        : getCountryColor(
                            d.Country
                        )
                );


            item.append("span")
                .text(
                    d.Country
                );


            item.on(
                "mouseenter",
                function (
                    event
                ) {


                    paths
                        .filter(
                            function (p) {

                                return (
                                    p.data.Country ===
                                    d.Country
                                );

                            }
                        )
                        .transition()
                        .duration(150)
                        .attr(
                            "d",
                            hoverArc
                        );


                    const percent =
                        totalValue > 0

                            ? (
                                d.Value /
                                totalValue *
                                100
                            )

                            : 0;


                    showTooltip(

                        event,

                        `<b>${d.Country}</b>
                         <br>จำนวนรายการ: ${d3.format(",")(d.Count)}
                         <br>มูลค่า: £${d3.format(",.2f")(d.Value)}
                         <br>สัดส่วน: ${percent.toFixed(1)}%`

                    );

                }
            );


            item.on(
                "mouseleave",
                function () {


                    paths
                        .filter(
                            function (p) {

                                return (
                                    p.data.Country ===
                                    d.Country
                                );

                            }
                        )
                        .transition()
                        .duration(150)
                        .attr(
                            "d",
                            arc
                        );


                    hideTooltip();

                }
            );


            item.on(
                "click",
                function () {


                    if (
                        d.Country !==
                        "Others"
                    ) {

                        selectedCountry =
                            d.Country;


                        d3.select(
                            "#countryFilter"
                        )
                        .property(
                            "value",
                            selectedCountry
                        );


                        applyFilters();

                    }

                }
            );

        }
    );

}


/* =========================================================
   3. COLUMN CHART
   ========================================================= */

function renderColumnChart() {


    const container =
        d3.select(
            "#columnChart"
        );


    container.html("");


    const node =
        container.node();


    if (!node) return;


    const bounds =
        node.getBoundingClientRect();


    const margin = {

        top: 25,
        right: 25,
        bottom: 55,
        left: 75

    };


    const width =
        Math.max(
            200,
            bounds.width -
            margin.left -
            margin.right
        );


    const height =
        Math.max(
            180,
            bounds.height -
            margin.top -
            margin.bottom
        );


    /*
       ไม่ใช้ selectedType
       เพราะกราฟนี้ต้องเปรียบเทียบ
       Sale กับ Return
    */

    const dataset =
        globalDataset.filter(
            function (d) {

                const countryOK =
                    selectedCountry ===
                    "ALL" ||
                    d.Country ===
                    selectedCountry;


                const yearOK =
                    selectedYear ===
                    "ALL" ||
                    String(
                        d.Year
                    ) ===
                    String(
                        selectedYear
                    );


                const monthOK =
                    selectedMonth ===
                    "ALL" ||
                    String(
                        d.Month
                    ) ===
                    String(
                        selectedMonth
                    );


                return (
                    countryOK &&
                    yearOK &&
                    monthOK
                );

            }
        );


    const typeData = [

        {
            Type: "Sale",
            Sales: d3.sum(
                dataset.filter(
                    d =>
                        d.TransactionType ===
                        "Sale"
                ),
                d =>
                    Number(
                        d.LineAmount
                    ) || 0
            ),

            Count:
                dataset.filter(
                    d =>
                        d.TransactionType ===
                        "Sale"
                ).length
        },

        {
            Type:
                "Return/Cancelled",

            Sales: d3.sum(
                dataset.filter(
                    d =>
                        d.TransactionType ===
                        "Return/Cancelled"
                ),
                d =>
                    Number(
                        d.LineAmount
                    ) || 0
            ),

            Count:
                dataset.filter(
                    d =>
                        d.TransactionType ===
                        "Return/Cancelled"
                ).length
        }

    ];


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                "100%"
            )
            .attr(
                "height",
                "100%"
            )
            .attr(
                "viewBox",
                `0 0 ${bounds.width} ${bounds.height}`
            );


    const g =
        svg.append("g")
            .attr(
                "transform",
                `translate(${margin.left},${margin.top})`
            );


    const x =
        d3.scaleBand()
            .domain(
                [
                    "Sale",
                    "Return/Cancelled"
                ]
            )
            .range(
                [0, width]
            )
            .padding(
                0.45
            );


    const maxValue =
        d3.max(
            typeData,
            d => d.Sales
        ) || 1;


    const y =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    maxValue * 1.15
                ]
            )
            .range(
                [height, 0]
            );


    g.append("g")
        .attr(
            "transform",
            `translate(0,${height})`
        )
        .call(
            d3.axisBottom(x)
        );


    g.append("g")
        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickFormat(
                    d =>
                        "£" +
                        d3.format(
                            ",.0f"
                        )(
                            d / 1000
                        ) +
                        "k"
                )
        );


    g.selectAll(
        ".column"
    )
    .data(
        typeData
    )
    .enter()
    .append("rect")
    .attr(
        "class",
        "column"
    )
    .attr(
        "x",
        d =>
            x(d.Type)
    )
    .attr(
        "width",
        x.bandwidth()
    )
    .attr(
        "y",
        d =>
            y(d.Sales)
    )
    .attr(
        "height",
        d =>
            height -
            y(d.Sales)
    )
    .attr(
        "rx",
        6
    )
    .attr(
        "fill",
        d =>
            TYPE_COLORS[
                d.Type
            ]
    )
    .on(
        "mouseover",
        function (
            event,
            d
        ) {

            showTooltip(

                event,

                `<b>${d.Type}</b>
                 <br>มูลค่ารายการ: £${d3.format(",.2f")(d.Sales)}
                 <br>จำนวนรายการ: ${d3.format(",")(d.Count)}`

            );

        }
    )
    .on(
        "mouseout",
        hideTooltip
    );

}


/* =========================================================
   4. SCATTER CHART
   ========================================================= */

function renderScatterChart() {


    const container =
        d3.select(
            "#scatterChart"
        );


    container.html("");


    const node =
        container.node();


    if (!node) return;


    const bounds =
        node.getBoundingClientRect();


    const margin = {

        top: 45,
        right: 25,
        bottom: 55,
        left: 60

    };


    const width =
        Math.max(
            200,
            bounds.width -
            margin.left -
            margin.right
        );


    const height =
        Math.max(
            180,
            bounds.height -
            margin.top -
            margin.bottom
        );


    let dataset =
        filteredData.filter(
            function (d) {

                return (
                    d.UnitPrice >= 0 &&
                    d.Quantity >= 0
                );

            }
        );


    if (
        localScatterPrice !==
        "all"
    ) {

        dataset =
            dataset.filter(
                d =>
                    d.UnitPrice <=
                    Number(
                        localScatterPrice
                    )
            );

    }


    if (
        localScatterQty > 0
    ) {

        dataset =
            dataset.filter(
                d =>
                    d.Quantity >=
                    localScatterQty
            );

    }


    /*
       จำกัดจุดไว้ 400 จุด
       เพื่อให้ Browser ลื่น
       แต่ยังคงกระจายจากต้นจนจบข้อมูล
    */

    let sampleData =
        dataset;


    if (
        dataset.length >
        400
    ) {

        sampleData =
            d3.range(
                400
            ).map(
                function (i) {

                    const index =
                        Math.floor(
                            i *
                            (
                                dataset.length -
                                1
                            ) /
                            399
                        );

                    return dataset[
                        index
                    ];

                }
            );

    }


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                "100%"
            )
            .attr(
                "height",
                "100%"
            )
            .attr(
                "viewBox",
                `0 0 ${bounds.width} ${bounds.height}`
            );


    const g =
        svg.append("g")
            .attr(
                "transform",
                `translate(${margin.left},${margin.top})`
            );


    const xMax =
        d3.max(
            sampleData,
            d =>
                d.UnitPrice
        ) || 10;


    const yMax =
        d3.max(
            sampleData,
            d =>
                d.Quantity
        ) || 10;


    const x =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    xMax * 1.05
                ]
            )
            .range(
                [0, width]
            );


    const y =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    yMax * 1.05
                ]
            )
            .range(
                [height, 0]
            );


    const xAxis =
        g.append("g")
            .attr(
                "transform",
                `translate(0,${height})`
            )
            .call(
                d3.axisBottom(x)
                    .ticks(6)
            );


    const yAxis =
        g.append("g")
            .call(
                d3.axisLeft(y)
                    .ticks(6)
            );


    /*
       จุดข้อมูล
    */

    const dots =
        g.selectAll(
            ".scatter-point"
        )
        .data(
            sampleData
        )
        .enter()
        .append("circle")
        .attr(
            "class",
            "scatter-point"
        )
        .attr(
            "cx",
            d =>
                x(d.UnitPrice)
        )
        .attr(
            "cy",
            d =>
                y(d.Quantity)
        )
        .attr(
            "r",
            4
        )
        .attr(
            "fill",
            d =>
                TYPE_COLORS[
                    d.TransactionType
                ] ||
                TYPE_COLORS.Sale
        )
        .attr(
            "opacity",
            0.7
        )
        .on(
            "mouseover",
            function (
                event,
                d
            ) {

                showTooltip(

                    event,

                    `<b>${d.Description}</b>
                     <br>ประเภท: ${d.TransactionType}
                     <br>ราคา: £${d3.format(",.2f")(d.UnitPrice)}
                     <br>จำนวน: ${d3.format(",")(d.Quantity)} ชิ้น`

                );

            }
        )
        .on(
            "mouseout",
            hideTooltip
        );


    /*
       Zoom
    */

    const zoom =
        d3.zoom()
            .scaleExtent(
                [1, 10]
            )
            .translateExtent(
                [
                    [0, 0],
                    [width, height]
                ]
            )
            .extent(
                [
                    [0, 0],
                    [width, height]
                ]
            )
            .on(
                "zoom",
                function (event) {

                    const newX =
                        event.transform
                            .rescaleX(x);


                    const newY =
                        event.transform
                            .rescaleY(y);


                    xAxis.call(
                        d3.axisBottom(
                            newX
                        )
                        .ticks(6)
                    );


                    yAxis.call(
                        d3.axisLeft(
                            newY
                        )
                        .ticks(6)
                    );


                    dots
                        .attr(
                            "cx",
                            d =>
                                newX(
                                    d.UnitPrice
                                )
                        )
                        .attr(
                            "cy",
                            d =>
                                newY(
                                    d.Quantity
                                )
                        );

                }
            );


    svg.call(
        zoom
    );


    d3.select(
        "#resetZoomBtn"
    )
    .on(
        "click",
        function () {

            svg
                .transition()
                .duration(600)
                .call(
                    zoom.transform,
                    d3.zoomIdentity
                );

        }
    );

}
