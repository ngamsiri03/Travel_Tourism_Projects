/* =========================================================
   TRAVEL & TOURISM INTERACTIVE DASHBOARD
   ========================================================= */

const DATA_FILE = "data/Travel And Tourism.xlsx";

let allData = [];
let selectedCountry = "ทั้งหมด";
let selectedTransport = null;

let countryChart = null;
let monthChart = null;
let transportChart = null;
let scatterChart = null;


/* =========================================================
   1. โหลดข้อมูล Excel
========================================================= */

async function loadData() {
    try {
        const response = await fetch(DATA_FILE);

        if (!response.ok) {
            throw new Error("ไม่สามารถโหลดไฟล์ Excel ได้");
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        allData = XLSX.utils.sheet_to_json(worksheet, {
            defval: ""
        });

        prepareData();
        createCountryFilter();
        createCharts();
        updateDashboard();

    } catch (error) {
        console.error(error);

        const loading = document.querySelector(".loading");

        if (loading) {
            loading.innerHTML = `
                <p style="color:red;">
                    ไม่สามารถโหลดข้อมูลได้
                </p>
            `;
        }
    }
}


/* =========================================================
   2. เตรียมข้อมูล
========================================================= */

function prepareData() {

    allData = allData.map(row => {

        let bookingDate = null;

        if (typeof row.Booking_Date === "number") {

            const parsed =
                XLSX.SSF.parse_date_code(
                    row.Booking_Date
                );

            if (parsed) {

                bookingDate =
                    new Date(
                        parsed.y,
                        parsed.m - 1,
                        parsed.d
                    );
            }

        } else {

            const value =
                String(
                    row.Booking_Date || ""
                ).trim();

            if (value) {

                const match =
                    value.match(
                        /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
                    );

                if (match) {

                    const day =
                        parseInt(
                            match[1],
                            10
                        );

                    const month =
                        parseInt(
                            match[2],
                            10
                        );

                    const year =
                        parseInt(
                            match[3],
                            10
                        );

                    bookingDate =
                        new Date(
                            year,
                            month - 1,
                            day
                        );

                } else {

                    const parsedDate =
                        new Date(value);

                    if (
                        !isNaN(
                            parsedDate.getTime()
                        )
                    ) {

                        bookingDate =
                            parsedDate;
                    }
                }
            }
        }

        return {
            ...row,

            Booking_Date_Obj:
                bookingDate,

            Customer_Age:
                Number(row.Customer_Age) || 0,

            Number_of_Travellers:
                Number(row.Number_of_Travellers) || 0,

            Number_of_Adults:
                Number(row.Number_of_Adults) || 0,

            Number_of_Children:
                Number(row.Number_of_Children) || 0,

            Number_of_Rooms:
                Number(row.Number_of_Rooms) || 0,

            Number_of_Nights:
                Number(row.Number_of_Nights) || 0,

            Discount_Amount:
                Number(row.Discount_Amount) || 0,

            Customer_Rating:
                Number(row.Customer_Rating) || 0,

            Total_Trip_Cost:
                Number(row.Total_Trip_Cost) || 0
        };
    });
}


/* =========================================================
   3. รายชื่อประเทศ
========================================================= */

function getCountries() {

    return [
        ...new Set(
            allData
                .map(row => row.Destination_Country)
                .filter(Boolean)
        )
    ].sort();
}


/* =========================================================
   4. สร้าง Dropdown ประเทศ
========================================================= */

function createCountryFilter() {

    const select =
        document.getElementById("countryFilter");

    if (!select) return;

    select.innerHTML = "";

    const allOption =
        document.createElement("option");

    allOption.value = "ทั้งหมด";
    allOption.textContent = "ทุกประเทศ";

    select.appendChild(allOption);

    getCountries().forEach(country => {

        const option =
            document.createElement("option");

        option.value = country;
        option.textContent = country;

        select.appendChild(option);
    });

    select.value = selectedCountry;

    select.addEventListener("change", function () {

        selectedCountry = this.value;

        animateDashboardUpdate();
    });
}


/* =========================================================
   5. กรองข้อมูล
========================================================= */

function getFilteredData() {

    return allData.filter(row => {

        const countryOK =
            selectedCountry === "ทั้งหมด" ||
            row.Destination_Country === selectedCountry;

        const transportOK =
            !selectedTransport ||
            row.Transportation_Type === selectedTransport;

        return countryOK && transportOK;
    });
}


/* =========================================================
   6. ข้อมูลสำหรับ Doughnut
========================================================= */

function getTransportData() {

    return allData.filter(row => {

        return selectedCountry === "ทั้งหมด" ||
            row.Destination_Country === selectedCountry;

    });
}


/* =========================================================
   7. สี Transport แบบคงที่
========================================================= */

const transportColors = {

    "Flight": "#38BDF8",

    "Train": "#8B5CF6",

    "Car": "#F59E0B"

};


/* =========================================================
   8. สีสำรอง
========================================================= */

function getTransportColor(type) {

    if (transportColors[type]) {
        return transportColors[type];
    }

    return "#94A3B8";
}


/* =========================================================
   9. Animation ตัวเลข
========================================================= */

function animateNumber(element, newValue, duration = 700) {

    if (!element) return;

    const startValue =
        Number(
            element.dataset.value || 0
        );

    const endValue =
        Number(newValue) || 0;

    const startTime = performance.now();

    function updateNumber(currentTime) {

        const progress =
            Math.min(
                (currentTime - startTime) / duration,
                1
            );

        const eased =
            1 - Math.pow(1 - progress, 3);

        const current =
            startValue +
            (endValue - startValue) * eased;

        element.textContent =
            Math.round(current).toLocaleString();

        if (progress < 1) {

            requestAnimationFrame(updateNumber);

        } else {

            element.dataset.value =
                endValue;
        }
    }

    requestAnimationFrame(updateNumber);
}


/* =========================================================
   10. Summary Cards
========================================================= */

function updateSummary(data) {

    const totalBookings =
        data.length;

    const averageCost =
        data.length
            ? data.reduce(
                (sum, row) =>
                    sum + row.Total_Trip_Cost,
                0
            ) / data.length
            : 0;

    const averageNights =
        data.length
            ? data.reduce(
                (sum, row) =>
                    sum + row.Number_of_Nights,
                0
            ) / data.length
            : 0;


    const elements = {

        bookings:
            document.getElementById("totalBookings"),

        topCountry:
            document.getElementById("topCountry"),

        avgCost:
            document.getElementById("avgCost"),

        avgNights:
            document.getElementById("avgNights")
    };


    animateNumber(
        elements.bookings,
        totalBookings
    );


    animateDecimal(
        elements.avgCost,
        Number(
            elements.avgCost.dataset.value || 0
        ),
        averageCost
    );


    animateDecimal(
        elements.avgNights,
        Number(
            elements.avgNights.dataset.value || 0
        ),
        averageNights
    );
}


/* =========================================================
   11. Animation ตัวเลขทศนิยม
========================================================= */

function animateDecimal(
    element,
    startValue,
    endValue,
    duration = 700
) {

    if (!element) return;

    const startTime =
        performance.now();

    function update(currentTime) {

        const progress =
            Math.min(
                (currentTime - startTime) / duration,
                1
            );

        const eased =
            1 - Math.pow(1 - progress, 3);

        const value =
            startValue +
            (endValue - startValue) * eased;

        element.textContent =
            value.toLocaleString(
                "en-US",
                {
                    minimumFractionDigits: 1,
                    maximumFractionDigits: 1
                }
            );

        if (progress < 1) {

            requestAnimationFrame(update);

        } else {

            element.dataset.value =
                endValue;
        }
    }

    requestAnimationFrame(update);
}


/* =========================================================
   12. Country Chart
========================================================= */

function getCountryChartData(data) {

    const countryCounts = {};

    data.forEach(row => {

        const country =
            row.Destination_Country;

        if (!country) return;

        countryCounts[country] =
            (countryCounts[country] || 0) + 1;
    });

    let countries =
        Object.keys(countryCounts)
            .sort(
                (a, b) =>
                    countryCounts[b] -
                    countryCounts[a]
            );

    let values =
        countries.map(
            country =>
                countryCounts[country]
        );

    if (countries.length === 0) {

        if (selectedCountry !== "ทั้งหมด") {

            countries = [
                selectedCountry
            ];

            values = [0];

        } else {

            countries = [
                "ไม่มีข้อมูล"
            ];

            values = [0];
        }
    }

    return {
        countries,
        values
    };
}


/* =========================================================
   13. Month Chart
========================================================= */

function getMonthChartData(data) {

    const monthNames = [
        "ม.ค.",
        "ก.พ.",
        "มี.ค.",
        "เม.ย.",
        "พ.ค.",
        "มิ.ย.",
        "ก.ค.",
        "ส.ค.",
        "ก.ย.",
        "ต.ค.",
        "พ.ย.",
        "ธ.ค."
    ];

    const counts =
        new Array(12).fill(0);

    data.forEach(row => {

        if (!row.Booking_Date_Obj) return;

        const month =
            row.Booking_Date_Obj.getMonth();

        if (
            month >= 0 &&
            month <= 11
        ) {
            counts[month]++;
        }
    });

    return {
        labels: monthNames,
        values: counts
    };
}


/* =========================================================
   14. Transport Chart
========================================================= */

function getTransportChartData(data) {

    const transportTypes = [
        "Flight",
        "Train",
        "Car"
    ];

    const counts =
        transportTypes.map(type => {

            return data.filter(
                row =>
                    row.Transportation_Type === type
            ).length;

        });

    return {
        labels: transportTypes,
        values: counts
    };
}


/* =========================================================
   15. Scatter Data
========================================================= */

function getScatterData(data) {

    return data
        .filter(row =>
            row.Number_of_Nights >= 0 &&
            row.Total_Trip_Cost >= 0
        )
        .map(row => ({

            x: row.Number_of_Nights,

            y: row.Total_Trip_Cost

        }));
}


/* =========================================================
   16. สร้าง Chart ครั้งแรก
========================================================= */

function createCharts() {

    createCountryChart();

    createMonthChart();

    createTransportChart();

    createScatterChart();
}


/* =========================================================
   17. Country Chart
========================================================= */

function createCountryChart() {

    const canvas =
        document.getElementById("countryChart");

    if (!canvas) return;

    const ctx =
        canvas.getContext("2d");

    countryChart =
        new Chart(ctx, {

            type: "bar",

            data: {

                labels: [],

                datasets: [{

                    label:
                        "จำนวนการจอง",

                    data: [],

                    backgroundColor:
                        "rgba(59, 130, 246, 0.75)",

                    borderColor:
                        "#3B82F6",

                    borderWidth: 1,

                    borderRadius: 5,

                    barThickness: 18,

                    maxBarThickness: 18,

                    categoryPercentage: 0.65,

                    barPercentage: 0.80
                }]
            },

            options: {

                indexAxis: "y",

                responsive: true,

                maintainAspectRatio: false,

                animation: {

                    duration: 900,

                    easing: "easeOutQuart"
                },

                plugins: {

                    legend: {
                        display: false
                    },

                    tooltip: {

                        callbacks: {

                            label: function(context) {

                                return (
                                    " จำนวนการจอง: " +
                                    context.parsed.x.toLocaleString()
                                );
                            }
                        }
                    },

                    datalabels: {

                        anchor: "end",

                        align: "right",

                        color: "#334155",

                        font: {

                            weight: "600",

                            size: 11
                        },

                        formatter: value =>
                            value.toLocaleString()
                    }
                },

                scales: {

                    x: {

                        beginAtZero: true,

                        ticks: {

                            precision: 0
                        }
                    },

                    y: {

                        ticks: {

                            font: {

                                size: 11
                            }
                        }
                    }
                }
            },

            plugins: [
                ChartDataLabels
            ]
        });
}


/* =========================================================
   18. Month Chart
========================================================= */

function createMonthChart() {

    const canvas =
        document.getElementById("monthChart");

    if (!canvas) return;

    const ctx =
        canvas.getContext("2d");

    monthChart =
        new Chart(ctx, {

            type: "line",

            data: {

                labels: [],

                datasets: [{

                    label:
                        "จำนวนการจอง",

                    data: [],

                    borderColor:
                        "#6366F1",

                    backgroundColor:
                        "rgba(99, 102, 241, 0.12)",

                    borderWidth: 3,

                    tension: 0.25,

                    fill: true,

                    pointRadius: 4,

                    pointHoverRadius: 7,

                    pointBackgroundColor:
                        "#6366F1",

                    pointBorderColor:
                        "#FFFFFF",

                    pointBorderWidth: 2
                }]
            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: {

                    duration: 900,

                    easing: "easeOutQuart"
                },

                interaction: {

                    intersect: false,

                    mode: "index"
                },

                plugins: {

                    tooltip: {

                        callbacks: {

                            label: function(context) {

                                return (
                                    " จำนวนการจอง: " +
                                    context.parsed.y.toLocaleString()
                                );
                            }
                        }
                    },

                    legend: {

                        display: false
                    }
                },

                scales: {

                    y: {

                        beginAtZero: true,

                        ticks: {

                            precision: 0
                        }
                    }
                }
            }
        });
}


/* =========================================================
   19. Doughnut Chart
========================================================= */

function createTransportChart() {

    const canvas =
        document.getElementById("transportChart");

    if (!canvas) return;

    const ctx =
        canvas.getContext("2d");

    transportChart =
        new Chart(ctx, {

            type: "doughnut",

            data: {

                labels: [
                    "Flight",
                    "Train",
                    "Car"
                ],

                datasets: [{

                    data: [0, 0, 0],

                    backgroundColor: [
                        transportColors.Flight,
                        transportColors.Train,
                        transportColors.Car
                    ],

                    borderWidth: 3,

                    borderColor: "#FFFFFF",

                    offset: [0, 0, 0]
                }]
            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                cutout: "62%",

                animation: {

                    duration: 900,

                    easing: "easeOutQuart"
                },

                plugins: {

                    legend: {

                        position: "bottom"
                    },

                    tooltip: {

                        callbacks: {

                            label: function(context) {

                                const values =
                                    context.dataset.data;

                                const total =
                                    values.reduce(
                                        (a, b) => a + b,
                                        0
                                    );

                                const value =
                                    context.raw;

                                const percent =
                                    total
                                        ? (
                                            value /
                                            total *
                                            100
                                        ).toFixed(1)
                                        : 0;

                                return (
                                    ` ${context.label}: ${value.toLocaleString()} (${percent}%)`
                                );
                            }
                        }
                    },

                    datalabels: {

                        color: "#FFFFFF",

                        font: {

                            weight: "bold",

                            size: 12
                        },

                        formatter: function(value, context) {

                            const values =
                                context.chart.data.datasets[0].data;

                            const total =
                                values.reduce(
                                    (a, b) => a + b,
                                    0
                                );

                            if (!total) return "";

                            return (
                                (
                                    value /
                                    total *
                                    100
                                ).toFixed(1) + "%"
                            );
                        }
                    }
                },

                onClick: function(event, elements) {

                    if (!elements.length) return;

                    const index =
                        elements[0].index;

                    const type =
                        this.data.labels[index];

                    if (selectedTransport === type) {

                        selectedTransport = null;

                    } else {

                        selectedTransport = type;
                    }

                    animateDashboardUpdate();
                }
            },

            plugins: [
                ChartDataLabels
            ]
        });
}


/* =========================================================
   20. Scatter Chart
========================================================= */

function createScatterChart() {

    const canvas =
        document.getElementById("scatterChart");

    if (!canvas) return;

    const ctx =
        canvas.getContext("2d");

    scatterChart =
        new Chart(ctx, {

            type: "scatter",

            data: {

                datasets: [{

                    label:
                        "จำนวนคืนพัก vs ค่าใช้จ่าย",

                    data: [],

                    backgroundColor:
                        "rgba(236, 72, 153, 0.65)",

                    borderColor:
                        "#EC4899",

                    pointRadius: 5,

                    pointHoverRadius: 8
                }]
            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                animation: {

                    duration: 900,

                    easing: "easeOutQuart"
                },

                plugins: {

                    legend: {

                        display: false
                    },

                    tooltip: {

                        callbacks: {

                            label: function(context) {

                                return [
                                    ` จำนวนคืน: ${context.parsed.x}`,
                                    ` ค่าใช้จ่าย: ${context.parsed.y.toLocaleString()}`
                                ];
                            }
                        }
                    }
                },

                scales: {

                    x: {

                        title: {

                            display: true,

                            text:
                                "จำนวนคืนพัก"
                        },

                        beginAtZero: true
                    },

                    y: {

                        title: {

                            display: true,

                            text:
                                "ค่าใช้จ่ายรวม"
                        },

                        beginAtZero: true
                    }
                }
            }
        });
}


/* =========================================================
   21. Update Country Chart
========================================================= */

function updateCountryChart(data) {

    if (!countryChart) return;

    const chartData =
        getCountryChartData(data);

    countryChart.data.labels =
        chartData.countries;

    countryChart.data.datasets[0].data =
        chartData.values;

    countryChart.update();
}


/* =========================================================
   22. Update Month Chart
========================================================= */

function updateMonthChart(data) {

    if (!monthChart) return;

    const chartData =
        getMonthChartData(data);

    monthChart.data.labels =
        chartData.labels;

    monthChart.data.datasets[0].data =
        chartData.values;

    monthChart.update();
}


/* =========================================================
   23. Update Doughnut
========================================================= */

function updateTransportChart() {

    if (!transportChart) return;

    const data =
        getTransportData();

    const chartData =
        getTransportChartData(data);

    transportChart.data.labels =
        chartData.labels;

    transportChart.data.datasets[0].data =
        chartData.values;

    transportChart.data.datasets[0].backgroundColor =
        chartData.labels.map(
            type =>
                getTransportColor(type)
        );

    transportChart.data.datasets[0].offset =
        chartData.labels.map(
            type =>
                selectedTransport === type
                    ? 15
                    : 0
        );

    transportChart.data.datasets[0].backgroundColor =
        chartData.labels.map(type => {

            const color =
                getTransportColor(type);

            if (
                selectedTransport &&
                selectedTransport !== type
            ) {

                return hexToRgba(
                    color,
                    0.25
                );
            }

            return color;
        });

    transportChart.update();
}


/* =========================================================
   24. Update Scatter
========================================================= */

function updateScatterChart(data) {

    if (!scatterChart) return;

    scatterChart.data.datasets[0].data =
        getScatterData(data);

    scatterChart.update();
}


/* =========================================================
   25. HEX → RGBA
========================================================= */

function hexToRgba(hex, alpha) {

    hex =
        hex.replace("#", "");

    const r =
        parseInt(
            hex.substring(0, 2),
            16
        );

    const g =
        parseInt(
            hex.substring(2, 4),
            16
        );

    const b =
        parseInt(
            hex.substring(4, 6),
            16
        );

    return `
        rgba(
            ${r},
            ${g},
            ${b},
            ${alpha}
        )
    `;
}


/* =========================================================
   26. Update Filter Text
========================================================= */

function updateFilterText() {

    const element =
        document.getElementById("activeFilter");

    if (!element) return;

    let text = "";

    if (
        selectedCountry !== "ทั้งหมด" &&
        selectedTransport
    ) {

        text =
            `ประเทศ: ${selectedCountry} | ประเภท: ${selectedTransport}`;

    } else if (
        selectedCountry !== "ทั้งหมด"
    ) {

        text =
            `ประเทศ: ${selectedCountry}`;

    } else if (
        selectedTransport
    ) {

        text =
            `ประเภท: ${selectedTransport}`;

    } else {

        text =
            "แสดงข้อมูลทั้งหมด";
    }

    element.textContent =
        text;
}


/* =========================================================
   27. Update Dashboard
========================================================= */

function updateDashboard() {

    const filteredData =
        getFilteredData();

    updateSummary(
        filteredData
    );

    updateCountryChart(
        filteredData
    );

    updateMonthChart(
        filteredData
    );

    updateTransportChart();

    updateScatterChart(
        filteredData
    );

    updateFilterText();

    updateTopCountry(
        filteredData
    );
}


/* =========================================================
   28. ประเทศที่มีการจองสูงสุด
========================================================= */

function updateTopCountry(data) {

    const element =
        document.getElementById("topCountry");

    if (!element) return;

    const counts = {};

    data.forEach(row => {

        const country =
            row.Destination_Country;

        if (!country) return;

        counts[country] =
            (counts[country] || 0) + 1;
    });

    const countries =
        Object.keys(counts);

    if (!countries.length) {

        element.textContent =
            "-";

        return;
    }

    const top =
        countries.sort(
            (a, b) =>
                counts[b] -
                counts[a]
        )[0];

    element.textContent =
        top;
}


/* =========================================================
   29. Animation ตอน Filter
========================================================= */

function animateDashboardUpdate() {

    const dashboard =
        document.querySelector(
            ".dashboard"
        );

    if (dashboard) {

        dashboard.classList.remove(
            "filter-refresh"
        );

        void dashboard.offsetWidth;

        dashboard.classList.add(
            "filter-refresh"
        );
    }

    updateDashboard();
}


/* =========================================================
   30. ปุ่ม Clear Filter
========================================================= */

function clearFilters() {

    selectedCountry =
        "ทั้งหมด";

    selectedTransport =
        null;

    const select =
        document.getElementById(
            "countryFilter"
        );

    if (select) {

        select.value =
            "ทั้งหมด";
    }

    animateDashboardUpdate();
}


/* =========================================================
   31. รองรับปุ่ม Clear Filter
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function() {

        const clearButton =
            document.getElementById(
                "clearFilter"
            );

        if (clearButton) {

            clearButton.addEventListener(
                "click",
                clearFilters
            );
        }

    }
);


/* =========================================================
   32. เริ่มระบบ
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function() {

        loadData();

    }
);
