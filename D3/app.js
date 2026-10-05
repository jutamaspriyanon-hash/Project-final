/* =========================================================
   ONLINE RETAIL ANALYTICS DASHBOARD
   อ่านข้อมูลจาก Online_Retail_Cleaned_Final-1.xlsb โดยตรง
   ========================================================= */

const DATA_URL = "Online_Retail_Cleaned_Final-1.xlsb";

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


/* =========================================================
   START
   ========================================================= */

document.addEventListener("DOMContentLoaded", function () {

    console.log("Dashboard started");

    setupEvents();

    loadXLSB();

});


/* =========================================================
   EVENTS
   ========================================================= */

function setupEvents() {

    d3.select("#countryFilter")
        .on("change", function () {

            selectedCountry = this.value;

            applyFilters();

        });


    d3.select("#typeFilter")
        .on("change", function () {

            selectedType = this.value;

            applyFilters();

        });


    d3.select("#yearFilter")
        .on("change", function () {

            selectedYear = this.value;

            selectedMonth = "ALL";

            populateMonthDropdown();

            applyFilters();

        });


    d3.select("#monthFilter")
        .on("change", function () {

            selectedMonth = this.value;

            applyFilters();

        });


    d3.select("#localProductMetric")
        .on("change", function () {

            localProductMetric = this.value;

            renderBarChart();

        });


    d3.select("#localProductTop")
        .on("change", function () {

            localProductTop = Number(this.value);

            renderBarChart();

        });


    d3.select("#localDonutTop")
        .on("change", function () {

            localDonutTop = Number(this.value);

            renderDonutChart();

        });


    d3.select("#localScatterPrice")
        .on("change", function () {

            localScatterPrice = this.value;

            renderScatterChart();

        });


    d3.select("#localScatterQty")
        .on("change", function () {

            localScatterQty = Number(this.value);

            renderScatterChart();

        });


    d3.select("#resetBtn")
        .on("click", resetDashboard);


    d3.select("#resetZoomBtn")
        .on("click", function () {

            renderScatterChart();

        });

}


/* =========================================================
   LOAD XLSB
   ========================================================= */

async function loadXLSB() {

    try {

        showLoadingMessage();

        console.log(
            "กำลังโหลด:",
            DATA_URL
        );


        /* ตรวจสอบว่า SheetJS โหลดแล้ว */

        if (typeof XLSX === "undefined") {

            throw new Error(
                "ไม่พบ SheetJS (XLSX)"
            );

        }


        /* โหลดไฟล์ XLSB */

        const response =
            await fetch(
                DATA_URL + "?v=" + Date.now()
            );


        if (!response.ok) {

            throw new Error(
                "โหลดไฟล์ไม่ได้: HTTP " +
                response.status
            );

        }


        const buffer =
            await response.arrayBuffer();


        console.log(
            "ขนาดไฟล์:",
            buffer.byteLength,
            "bytes"
        );


        if (buffer.byteLength === 0) {

            throw new Error(
                "ไฟล์ XLSB ว่าง"
            );

        }


        /* อ่าน XLSB */

        const workbook =
            XLSX.read(
                buffer,
                {
                    type: "array",
                    cellDates: true,
                    dense: false
                }
            );


        console.log(
            "Sheets:",
            workbook.SheetNames
        );


        if (
            !workbook.SheetNames ||
            workbook.SheetNames.length === 0
        ) {

            throw new Error(
                "ไม่พบ Sheet ในไฟล์ XLSB"
            );

        }


        /* ใช้ Sheet แรก */

        const sheetName =
            workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[sheetName];


        console.log(
            "กำลังอ่าน Sheet:",
            sheetName
        );


        /* แปลง Sheet เป็น Object */

        const rows =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: "",
                    raw: true
                }
            );


        console.log(
            "จำนวนแถว:",
            rows.length
        );


        if (!rows.length) {

            throw new Error(
                "Sheet ไม่มีข้อมูล"
            );

        }


        console.log(
            "คอลัมน์:",
            Object.keys(rows[0])
        );


        /*
         * แปลงข้อมูล
         */

        convertData(rows);


    } catch (error) {

        console.error(
            "XLSB ERROR:",
            error
        );

        showError(error);

    }

}


/* =========================================================
   CONVERT DATA
   ========================================================= */

function convertData(rows) {

    globalDataset = [];


    rows.forEach(function (row, index) {

        try {

            const description =
                getValue(
                    row,
                    [
                        "Description",
                        "description",
                        "DESCRIPTION"
                    ]
                );


            const quantity =
                toNumber(
                    getValue(
                        row,
                        [
                            "Quantity",
                            "quantity",
                            "QUANTITY"
                        ]
                    )
                );


            const unitPrice =
                toNumber(
                    getValue(
                        row,
                        [
                            "UnitPrice",
                            "Unit Price",
                            "unitprice",
                            "UNITPRICE"
                        ]
                    )
                );


            const customerID =
                getValue(
                    row,
                    [
                        "CustomerID",
                        "Customer ID",
                        "customerid",
                        "CUSTOMERID"
                    ]
                );


            const country =
                getValue(
                    row,
                    [
                        "Country",
                        "country",
                        "COUNTRY"
                    ]
                ) ||
                "Unknown";


            let transactionType =
                getValue(
                    row,
                    [
                        "TransactionType",
                        "Transaction Type",
                        "transactiontype",
                        "TRANSACTIONTYPE"
                    ]
                );


            let lineAmount =
                toNumber(
                    getValue(
                        row,
                        [
                            "LineAmount",
                            "Line Amount",
                            "lineamount",
                            "LINEAMOUNT"
                        ]
                    )
                );


            /*
             * ถ้าไม่มี LineAmount
             * คำนวณ Quantity × UnitPrice
             */

            if (
                !Number.isFinite(lineAmount)
            ) {

                lineAmount =
                    quantity * unitPrice;

            }


            /*
             * ถ้าไม่มี TransactionType
             * ใช้ Quantity เป็นตัวตัดสิน
             */

            if (!transactionType) {

                if (quantity < 0) {

                    transactionType =
                        "Return/Cancelled";

                } else {

                    transactionType =
                        "Sale";

                }

            }


            /*
             * ทำให้ประเภทเป็นมาตรฐาน
             */

            transactionType =
                normalizeTransactionType(
                    transactionType,
                    quantity
                );


            /*
             * วันที่
             */

            const rawDate =
                getValue(
                    row,
                    [
                        "InvoiceDate",
                        "Invoice Date",
                        "invoicedate",
                        "INVOICEDATE"
                    ]
                );


            const date =
                parseExcelDate(
                    rawDate
                );


            /*
             * ปี / เดือน
             */

            let year = null;
            let month = null;


            if (date) {

                year =
                    date.getFullYear();

                month =
                    date.getMonth() + 1;

            }


            /*
             * เก็บข้อมูล
             */

            globalDataset.push({

                Description:
                    String(
                        description ||
                        "Unknown Product"
                    ),

                Quantity:
                    quantity,

                UnitPrice:
                    unitPrice,

                CustomerID:
                    customerID,

                Country:
                    String(country),

                TransactionType:
                    transactionType,

                LineAmount:
                    lineAmount,

                InvoiceDate:
                    rawDate,

                Date:
                    date,

                Year:
                    year,

                Month:
                    month

            });

        } catch (error) {

            console.warn(
                "ข้ามแถว:",
                index,
                error
            );

        }

    });


    console.log(
        "ข้อมูลหลังแปลง:",
        globalDataset.length
    );


    if (!globalDataset.length) {

        throw new Error(
            "ไม่สามารถแปลงข้อมูลจาก XLSB ได้"
        );

    }


    /*
     * เริ่มต้นด้วยข้อมูลทั้งหมด
     */

    filteredData =
        [...globalDataset];


    /*
     * สร้าง Filter
     */

    populateCountryDropdown();

    populateYearDropdown();

    populateMonthDropdown();


    /*
     * แสดง Dashboard
     */

    updateDashboard();


    hideLoadingMessage();


    console.log(
        "โหลดข้อมูลสำเร็จ:",
        globalDataset.length
    );

}


/* =========================================================
   GET VALUE
   ========================================================= */

function getValue(row, names) {

    for (const name of names) {

        if (
            Object.prototype.hasOwnProperty.call(
                row,
                name
            )
        ) {

            return row[name];

        }

    }


    /*
     * ตรวจสอบแบบไม่สนใจตัวพิมพ์
     */

    const keys =
        Object.keys(row);


    for (const name of names) {

        const found =
            keys.find(
                key =>
                    key.toLowerCase() ===
                    name.toLowerCase()
            );


        if (found) {

            return row[found];

        }

    }


    return "";

}


/* =========================================================
   NUMBER
   ========================================================= */

function toNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return NaN;

    }


    if (
        typeof value === "number"
    ) {

        return value;

    }


    const cleaned =
        String(value)
            .replace(/,/g, "")
            .replace(/£/g, "")
            .trim();


    const number =
        Number(cleaned);


    return number;

}


/* =========================================================
   TRANSACTION TYPE
   ========================================================= */

function normalizeTransactionType(
    value,
    quantity
) {

    const text =
        String(value)
            .toLowerCase()
            .trim();


    if (
        text.includes("return") ||
        text.includes("cancel") ||
        text.includes("refund")
    ) {

        return "Return/Cancelled";

    }


    if (
        text === "sale" ||
        text === "ขาย"
    ) {

        return "Sale";

    }


    /*
     * ถ้าไม่รู้ประเภท
     * ใช้ Quantity
     */

    if (
        Number(quantity) < 0
    ) {

        return "Return/Cancelled";

    }


    return "Sale";

}


/* =========================================================
   EXCEL DATE
   ========================================================= */

function parseExcelDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    /*
     * Date object
     */

    if (
        value instanceof Date
    ) {

        if (
            !Number.isNaN(
                value.getTime()
            )
        ) {

            return value;

        }

        return null;

    }


    /*
     * Excel Serial Date
     *
     * Excel ใช้วันที่เริ่มจาก
     * 1899-12-30
     */

    if (
        typeof value === "number"
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


        if (
            !Number.isNaN(
                date.getTime()
            )
        ) {

            return date;

        }

    }


    const text =
        String(value)
            .trim();


    /*
     * ลอง Date.parse
     */

    const nativeDate =
        new Date(text);


    if (
        !Number.isNaN(
            nativeDate.getTime()
        )
    ) {

        return nativeDate;

    }


    /*
     * รูปแบบ dd/mm/yyyy
     */

    let match =
        text.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
        );


    if (match) {

        const day =
            Number(match[1]);

        const month =
            Number(match[2]) - 1;

        const year =
            Number(match[3]);


        const date =
            new Date(
                year,
                month,
                day
            );


        if (
            !Number.isNaN(
                date.getTime()
            )
        ) {

            return date;

        }

    }


    return null;

}


/* =========================================================
   COUNTRY FILTER
   ========================================================= */

function populateCountryDropdown() {

    const countries =
        Array.from(
            new Set(
                globalDataset
                    .map(
                        d => d.Country
                    )
                    .filter(Boolean)
            )
        )
        .sort();


    const select =
        d3.select(
            "#countryFilter"
        );


    select
        .selectAll(
            "option:not(:first-child)"
        )
        .remove();


    countries.forEach(
        country => {

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


    /*
     * อัปเดตจำนวนประเทศ
     */

    const label =
        select
            .node()
            ?.parentElement
            ?.querySelector("label");


    if (label) {

        label.innerHTML =
            `<i class="fa-solid fa-globe"></i>
             ประเทศทั้งหมด (${countries.length} ประเทศ):`;

    }

}


/* =========================================================
   YEAR FILTER
   ========================================================= */

function populateYearDropdown() {

    const years =
        Array.from(
            new Set(
                globalDataset
                    .map(
                        d => d.Year
                    )
                    .filter(
                        d =>
                            d !== null &&
                            d !== undefined
                    )
            )
        )
        .sort(
            (a, b) =>
                a - b
        );


    const select =
        d3.select(
            "#yearFilter"
        );


    select
        .selectAll(
            "option:not(:first-child)"
        )
        .remove();


    years.forEach(
        year => {

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
   MONTH FILTER
   ========================================================= */

function populateMonthDropdown() {

    const months = [

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


    let data =
        globalDataset;


    if (
        selectedYear !== "ALL"
    ) {

        data =
            data.filter(
                d =>
                    String(d.Year) ===
                    String(selectedYear)
            );

    }


    const monthNumbers =
        Array.from(
            new Set(
                data
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


    const select =
        d3.select(
            "#monthFilter"
        );


    select
        .selectAll(
            "option:not(:first-child)"
        )
        .remove();


    monthNumbers.forEach(
        month => {

            select
                .append("option")
                .attr(
                    "value",
                    month
                )
                .text(
                    String(month)
                    .padStart(2, "0") +
                    " - " +
                    months[month - 1]
                );

        }
    );

}


/* =========================================================
   FILTER DATA
   ========================================================= */

function applyFilters() {

    filteredData =
        globalDataset.filter(
            d => {

                const countryOK =
                    selectedCountry === "ALL" ||
                    d.Country ===
                    selectedCountry;


                const typeOK =
                    selectedType === "ALL" ||
                    d.TransactionType ===
                    selectedType;


                const yearOK =
                    selectedYear === "ALL" ||
                    String(d.Year) ===
                    String(selectedYear);


                const monthOK =
                    selectedMonth === "ALL" ||
                    String(d.Month) ===
                    String(selectedMonth);


                return (
                    countryOK &&
                    typeOK &&
                    yearOK &&
                    monthOK
                );

            }
        );


    updateDashboard();

}


/* =========================================================
   RESET
   ========================================================= */

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


    d3.select(
        "#countryFilter"
    ).property(
        "value",
        "ALL"
    );


    d3.select(
        "#typeFilter"
    ).property(
        "value",
        "ALL"
    );


    d3.select(
        "#yearFilter"
    ).property(
        "value",
        "ALL"
    );


    d3.select(
        "#localProductMetric"
    ).property(
        "value",
        "value"
    );


    d3.select(
        "#localProductTop"
    ).property(
        "value",
        "10"
    );


    d3.select(
        "#localDonutTop"
    ).property(
        "value",
        "5"
    );


    d3.select(
        "#localScatterPrice"
    ).property(
        "value",
        "all"
    );


    d3.select(
        "#localScatterQty"
    ).property(
        "value",
        "0"
    );


    populateMonthDropdown();


    d3.select(
        "#monthFilter"
    ).property(
        "value",
        "ALL"
    );


    filteredData =
        [...globalDataset];


    updateDashboard();

}


/* =========================================================
   UPDATE DASHBOARD
   ========================================================= */

function updateDashboard() {

    renderKPIs();

    renderBarChart();

    renderDonutChart();

    renderColumnChart();

    renderScatterChart();

}


/* =========================================================
   KPI
   ========================================================= */

function renderKPIs() {

    const totalSales =
        d3.sum(
            filteredData,
            d =>
                Number(d.LineAmount) ||
                0
        );


    const totalQty =
        d3.sum(
            filteredData,
            d =>
                Number(d.Quantity) ||
                0
        );


    const totalOrders =
        filteredData.length;


    const countries =
        new Set(
            filteredData.map(
                d => d.Country
            )
        );


    d3.select(
        "#kpiTotalSales"
    ).text(
        "£" +
        d3.format(",.2f")(
            totalSales
        )
    );


    d3.select(
        "#kpiTotalQty"
    ).text(
        d3.format(",")(
            totalQty
        ) +
        " ชิ้น"
    );


    d3.select(
        "#kpiTotalOrders"
    ).text(
        d3.format(",")(
            totalOrders
        ) +
        " รายการ"
    );


    d3.select(
        "#kpiTotalCountries"
    ).text(
        countries.size +
        " ประเทศ"
    );

}


/* =========================================================
   TOOLTIP
   ========================================================= */

function showTooltip(
    event,
    html
) {

    const tooltip =
        d3.select(
            "#tooltip"
        );


    tooltip
        .html(html)
        .style(
            "opacity",
            1
        )
        .style(
            "left",
            (event.pageX + 15) +
            "px"
        )
        .style(
            "top",
            (event.pageY - 30) +
            "px"
        );

}


function hideTooltip() {

    d3.select(
        "#tooltip"
    )
    .style(
        "opacity",
        0
    );

}


/* =========================================================
   BAR CHART
   ========================================================= */

function renderBarChart() {

    const container =
        d3.select(
            "#barChart"
        );


    if (container.empty()) {
        return;
    }


    container.html("");


    const element =
        container.node();


    const width =
        element.clientWidth ||
        600;


    const height =
        element.clientHeight ||
        350;


    const margin = {

        top: 20,
        right: 30,
        bottom: 50,
        left: 190

    };


    const innerWidth =
        Math.max(
            100,
            width -
            margin.left -
            margin.right
        );


    const innerHeight =
        Math.max(
            100,
            height -
            margin.top -
            margin.bottom
        );


    const grouped =
        d3.rollup(

            filteredData,

            values => ({

                value:
                    d3.sum(
                        values,
                        d =>
                            Number(
                                d.LineAmount
                            ) || 0
                    ),

                qty:
                    d3.sum(
                        values,
                        d =>
                            Number(
                                d.Quantity
                            ) || 0
                    )

            }),

            d => d.Description

        );


    const data =
        Array.from(
            grouped,
            ([Description, value]) => ({

                Description,
                ...value

            })
        )
        .sort(
            (a, b) => {

                const av =
                    localProductMetric === "qty"
                        ? a.qty
                        : a.value;

                const bv =
                    localProductMetric === "qty"
                        ? b.qty
                        : b.value;

                return bv - av;

            }
        )
        .slice(
            0,
            localProductTop
        );


    if (!data.length) {

        showChartEmpty(
            container
        );

        return;

    }


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                width
            )
            .attr(
                "height",
                height
            )
            .append("g")
            .attr(
                "transform",
                `translate(${margin.left},${margin.top})`
            );


    const y =
        d3.scaleBand()
            .domain(
                data.map(
                    d => d.Description
                )
            )
            .range(
                [
                    0,
                    innerHeight
                ]
            )
            .padding(0.25);


    const maxValue =
        d3.max(
            data,
            d =>
                localProductMetric === "qty"
                    ? d.qty
                    : d.value
        ) || 1;


    const x =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    maxValue * 1.1
                ]
            )
            .range(
                [
                    0,
                    innerWidth
                ]
            );


    svg
        .append("g")
        .call(
            d3.axisLeft(y)
        );


    svg
        .append("g")
        .attr(
            "transform",
            `translate(0,${innerHeight})`
        )
        .call(
            d3.axisBottom(x)
        );


    svg
        .selectAll(".bar")
        .data(data)
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
                y(d.Description)
        )
        .attr(
            "height",
            y.bandwidth()
        )
        .attr(
            "width",
            d =>
                x(
                    localProductMetric === "qty"
                        ? d.qty
                        : d.value
                )
        )
        .attr(
            "fill",
            "#8b5cf6"
        )
        .attr(
            "rx",
            5
        )
        .on(
            "mousemove",
            function (event, d) {

                showTooltip(

                    event,

                    `<b>${d.Description}</b>
                    <br>มูลค่า: £${d3.format(",.2f")(d.value)}
                    <br>จำนวน: ${d3.format(",")(d.qty)} ชิ้น>`

                );

            }
        )
        .on(
            "mouseout",
            hideTooltip
        );

}


/* =========================================================
   DONUT CHART
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


    const width =
        container.node().clientWidth ||
        600;


    const height =
        container.node().clientHeight ||
        350;


    const grouped =
        d3.rollup(

            filteredData,

            values =>
                d3.sum(
                    values,
                    d =>
                        Number(
                            d.LineAmount
                        ) || 0
                ),

            d => d.Country

        );


    let data =
        Array.from(
            grouped,
            ([Country, value]) => ({
                Country,
                value
            })
        )
        .sort(
            (a, b) =>
                b.value - a.value
        );


    if (!data.length) {

        showChartEmpty(
            container
        );

        return;

    }


    if (
        localDonutTop !== 999
    ) {

        const top =
            data.slice(
                0,
                localDonutTop
            );


        const otherValue =
            d3.sum(
                data.slice(
                    localDonutTop
                ),
                d => d.value
            );


        if (otherValue > 0) {

            top.push({

                Country: "อื่น ๆ",

                value: otherValue

            });

        }


        data = top;

    }


    const radius =
        Math.min(
            width,
            height
        ) / 3;


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                width
            )
            .attr(
                "height",
                height
            );


    const g =
        svg
            .append("g")
            .attr(
                "transform",
                `translate(${width / 2},${height / 2})`
            );


    const pie =
        d3.pie()
            .value(
                d => d.value
            );


    const arc =
        d3.arc()
            .innerRadius(
                radius * 0.55
            )
            .outerRadius(
                radius
            );


    const arcs =
        g
            .selectAll("path")
            .data(
                pie(data)
            )
            .enter()
            .append("path")
            .attr(
                "d",
                arc
            )
            .attr(
                "fill",
                function (d, i) {

                    return d.data.Country === "อื่น ๆ"
                        ? "#cbd5e1"
                        : d3.schemeTableau10[
                            i %
                            d3.schemeTableau10.length
                        ];

                }
            )
            .attr(
                "stroke",
                "#ffffff"
            )
            .attr(
                "stroke-width",
                2
            );


    arcs
        .on(
            "mousemove",
            function (event, d) {

                const total =
                    d3.sum(
                        data,
                        x => x.value
                    );


                const percent =
                    total
                        ? (
                            d.data.value /
                            total *
                            100
                        ).toFixed(1)
                        : 0;


                showTooltip(

                    event,

                    `<b>${d.data.Country}</b>
                    <br>มูลค่า: £${d3.format(",.2f")(d.data.value)}
                    <br>สัดส่วน: ${percent}%`

                );

            }
        )
        .on(
            "mouseout",
            hideTooltip
        );


    g
        .append("text")
        .attr(
            "text-anchor",
            "middle"
        )
        .attr(
            "dy",
            "0.3em"
        )
        .style(
            "font-size",
            "16px"
        )
        .style(
            "font-weight",
            "bold"
        )
        .text(
            "ยอดขาย"
        );

}


/* =========================================================
   COLUMN CHART
   ========================================================= */

function renderColumnChart() {

    const container =
        d3.select(
            "#columnChart"
        );


    if (container.empty()) {
        return;
    }


    container.html("");


    const width =
        container.node().clientWidth ||
        600;


    const height =
        container.node().clientHeight ||
        350;


    const data =
        Array.from(

            d3.rollup(

                filteredData,

                values =>
                    d3.sum(
                        values,
                        d =>
                            Number(
                                d.LineAmount
                            ) || 0
                    ),

                d =>
                    d.TransactionType

            ),

            ([type, value]) => ({
                type,
                value
            })

        );


    if (!data.length) {

        showChartEmpty(
            container
        );

        return;

    }


    const margin = {

        top: 20,
        right: 20,
        bottom: 50,
        left: 70

    };


    const innerWidth =
        width -
        margin.left -
        margin.right;


    const innerHeight =
        height -
        margin.top -
        margin.bottom;


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                width
            )
            .attr(
                "height",
                height
            )
            .append("g")
            .attr(
                "transform",
                `translate(${margin.left},${margin.top})`
            );


    const x =
        d3.scaleBand()
            .domain(
                data.map(
                    d => d.type
                )
            )
            .range(
                [
                    0,
                    innerWidth
                ]
            )
            .padding(0.35);


    const max =
        d3.max(
            data,
            d => d.value
        ) || 1;


    const y =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    max * 1.1
                ]
            )
            .range(
                [
                    innerHeight,
                    0
                ]
            );


    svg
        .append("g")
        .attr(
            "transform",
            `translate(0,${innerHeight})`
        )
        .call(
            d3.axisBottom(x)
        );


    svg
        .append("g")
        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickFormat(
                    d =>
                        "£" +
                        d3.format(",.0f")(d)
                )
        );


    svg
        .selectAll(".column")
        .data(data)
        .enter()
        .append("rect")
        .attr(
            "class",
            "column"
        )
        .attr(
            "x",
            d =>
                x(d.type)
        )
        .attr(
            "width",
            x.bandwidth()
        )
        .attr(
            "y",
            d =>
                y(d.value)
        )
        .attr(
            "height",
            d =>
                innerHeight -
                y(d.value)
        )
        .attr(
            "rx",
            6
        )
        .attr(
            "fill",
            d =>
                d.type === "Sale"
                    ? "#10b981"
                    : "#ef4444"
        )
        .on(
            "mousemove",
            function (event, d) {

                showTooltip(

                    event,

                    `<b>${d.type}</b>
                    <br>มูลค่า: £${d3.format(",.2f")(d.value)}`

                );

            }
        )
        .on(
            "mouseout",
            hideTooltip
        );

}


/* =========================================================
   SCATTER
   ========================================================= */

function renderScatterChart() {

    const container =
        d3.select(
            "#scatterChart"
        );


    if (container.empty()) {
        return;
    }


    container.html("");


    let data =
        filteredData.filter(
            d =>
                Number.isFinite(
                    Number(d.UnitPrice)
                ) &&
                Number.isFinite(
                    Number(d.Quantity)
                ) &&
                Number(d.UnitPrice) >= 0
        );


    if (
        localScatterPrice !== "all"
    ) {

        data =
            data.filter(
                d =>
                    Number(
                        d.UnitPrice
                    ) <=
                    Number(
                        localScatterPrice
                    )
            );

    }


    if (
        localScatterQty > 0
    ) {

        data =
            data.filter(
                d =>
                    Number(
                        d.Quantity
                    ) >=
                    localScatterQty
            );

    }


    /*
     * จำกัดจุดเพื่อไม่ให้ Browser หนัก
     */

    if (data.length > 1000) {

        data =
            data.slice(
                0,
                1000
            );

    }


    if (!data.length) {

        showChartEmpty(
            container
        );

        return;

    }


    const width =
        container.node().clientWidth ||
        600;


    const height =
        container.node().clientHeight ||
        350;


    const margin = {

        top: 20,
        right: 30,
        bottom: 50,
        left: 65

    };


    const innerWidth =
        width -
        margin.left -
        margin.right;


    const innerHeight =
        height -
        margin.top -
        margin.bottom;


    const maxPrice =
        d3.max(
            data,
            d =>
                Number(
                    d.UnitPrice
                )
        ) || 1;


    const maxQty =
        d3.max(
            data,
            d =>
                Number(
                    d.Quantity
                )
        ) || 1;


    const x =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    maxPrice * 1.05
                ]
            )
            .range(
                [
                    0,
                    innerWidth
                ]
            );


    const y =
        d3.scaleLinear()
            .domain(
                [
                    0,
                    maxQty * 1.05
                ]
            )
            .range(
                [
                    innerHeight,
                    0
                ]
            );


    const svg =
        container
            .append("svg")
            .attr(
                "width",
                width
            )
            .attr(
                "height",
                height
            )
            .append("g")
            .attr(
                "transform",
                `translate(${margin.left},${margin.top})`
            );


    svg
        .append("g")
        .attr(
            "transform",
            `translate(0,${innerHeight})`
        )
        .call(
            d3.axisBottom(x)
        );


    svg
        .append("g")
        .call(
            d3.axisLeft(y)
        );


    svg
        .selectAll(".dot")
        .data(data)
        .enter()
        .append("circle")
        .attr(
            "class",
            "dot"
        )
        .attr(
            "cx",
            d =>
                x(
                    Number(
                        d.UnitPrice
                    )
                )
        )
        .attr(
            "cy",
            d =>
                y(
                    Number(
                        d.Quantity
                    )
                )
        )
        .attr(
            "r",
            4
        )
        .attr(
            "fill",
            d =>
                d.TransactionType ===
                "Return/Cancelled"
                    ? "#ef4444"
                    : "#10b981"
        )
        .attr(
            "opacity",
            0.65
        )
        .on(
            "mousemove",
            function (event, d) {

                showTooltip(

                    event,

                    `<b>${d.Description}</b>
                    <br>ราคา: £${d3.format(",.2f")(d.UnitPrice)}
                    <br>จำนวน: ${d3.format(",")(d.Quantity)}
                    <br>ประเภท: ${d.TransactionType}`

                );

            }
        )
        .on(
            "mouseout",
            hideTooltip
        );

}


/* =========================================================
   EMPTY CHART
   ========================================================= */

function showChartEmpty(
    container
) {

    container
        .append("div")
        .style(
            "height",
            "100%"
        )
        .style(
            "display",
            "flex"
        )
        .style(
            "align-items",
            "center"
        )
        .style(
            "justify-content",
            "center"
        )
        .style(
            "color",
            "#94a3b8"
        )
        .style(
            "font-size",
            "18px"
        )
        .text(
            "ไม่พบข้อมูล"
        );

}


/* =========================================================
   LOADING
   ========================================================= */

function showLoadingMessage() {

    let box =
        document.getElementById(
            "dataLoading"
        );


    if (!box) {

        box =
            document.createElement(
                "div"
            );


        box.id =
            "dataLoading";


        box.style.cssText = `
            position:fixed;
            top:20px;
            right:20px;
            z-index:99999;
            padding:15px 20px;
            background:#ffffff;
            border-radius:12px;
            box-shadow:0 5px 20px rgba(0,0,0,.15);
            font-family:Sarabun,sans-serif;
        `;


        document.body.appendChild(
            box
        );

    }


    box.innerHTML =
        "กำลังโหลดข้อมูล XLSB...";

}


/* =========================================================
   HIDE LOADING
   ========================================================= */

function hideLoadingMessage() {

    const box =
        document.getElementById(
            "dataLoading"
        );


    if (box) {

        box.remove();

    }

}


/* =========================================================
   ERROR
   ========================================================= */

function showError(error) {

    hideLoadingMessage();


    let box =
        document.getElementById(
            "dataLoadError"
        );


    if (box) {

        box.remove();

    }


    box =
        document.createElement(
            "div"
        );


    box.id =
        "dataLoadError";


    box.style.cssText = `
        position:fixed;
        left:20px;
        right:20px;
        top:20px;
        z-index:99999;
        padding:25px;
        background:#fff1f2;
        color:#881337;
        border:2px solid #fb7185;
        border-radius:15px;
        font-family:Sarabun,sans-serif;
        box-shadow:0 10px 30px rgba(0,0,0,.2);
    `;


    box.innerHTML = `

        <h2>
            ไม่สามารถโหลดข้อมูล XLSB
        </h2>

        <p>
            ไฟล์ที่ต้องการ:
            <b>${DATA_URL}</b>
        </p>

        <p>
            กรุณาตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์เดียวกับ
            <b>index.html</b>
        </p>

        <p>
            รายละเอียด:
            <b>${error.message || error}</b>
        </p>

    `;


    document.body.appendChild(
        box
    );

}
