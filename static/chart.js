document.addEventListener("DOMContentLoaded", () => {
    const container = document.getElementById("chart-data");
    
    // Check if chart container exists (only if there are expenses)
    if (!container) {
        return;
    }

    try {
        const dates = JSON.parse(container.dataset.dates);
        const amounts = JSON.parse(container.dataset.amounts);

        // Check if we have data
        if (!dates || !amounts || dates.length === 0) {
            return;
        }

        const ctx = document.getElementById("expenseChart");
        
        if (!ctx) {
            console.error("Canvas element not found");
            return;
        }

        new Chart(ctx.getContext("2d"), {
            type: "line",
            data: {
                labels: dates,
                datasets: [{
                    label: "Daily Expenses (₵)",
                    data: amounts,
                    borderColor: "#38bdf8",
                    backgroundColor: "rgba(56, 189, 248, 0.25)",
                    tension: 0.4,
                    fill: true,
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    pointBackgroundColor: "#38bdf8",
                    pointBorderColor: "#020617",
                    pointBorderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: true,
                        labels: {
                            color: "#e5e7eb",
                            font: {
                                size: 14
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: "rgba(2, 6, 23, 0.9)",
                        titleColor: "#38bdf8",
                        bodyColor: "#e5e7eb",
                        borderColor: "#38bdf8",
                        borderWidth: 1,
                        padding: 12,
                        displayColors: false,
                        callbacks: {
                            label: function(context) {
                                return "₵" + context.parsed.y.toFixed(2);
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: {
                            color: "#1e293b",
                            drawBorder: false
                        },
                        ticks: {
                            color: "#94a3b8"
                        }
                    },
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: "#1e293b",
                            drawBorder: false
                        },
                        ticks: {
                            color: "#94a3b8",
                            callback: function(value) {
                                return "₵" + value;
                            }
                        }
                    }
                }
            }
        });
    } catch (error) {
        console.error("Error rendering chart:", error);
    }
});