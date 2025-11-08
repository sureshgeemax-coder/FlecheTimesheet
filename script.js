// Populate year dropdown
const yearSelect = document.getElementById('year');
const currentYear = new Date().getFullYear();
for (let i = currentYear - 5; i <= currentYear + 5; i++) {
    const option = document.createElement('option');
    option.value = i;
    option.textContent = i;
    yearSelect.appendChild(option);
}

// Get month names
const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];

const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Time Charged options
const timeChargedOptions = ['1', 'Public Holiday', 'Sick Leave', 'Annual Leave'];

// Engagement options
const engagementOptions = [
    'Onsite Changi Business Park',
    'Asia Square TownSquare',
    'Work from Home'
];

// Form submission handler
document.getElementById('timesheetForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const staffName = document.getElementById('staffName').value;
    const companyName = document.getElementById('companyName').value;
    const clientName = document.getElementById('clientName').value;
    const month = parseInt(document.getElementById('month').value);
    const year = parseInt(document.getElementById('year').value);
    
    if (!staffName || month === '' || year === '') {
        alert('Please fill in all required fields');
        return;
    }
    
    await generateExcel(staffName, companyName, clientName, month, year);
});

async function generateExcel(staffName, companyName, clientName, month, year) {
    // Check if ExcelJS is loaded
    if (typeof ExcelJS === 'undefined') {
        alert('ExcelJS library failed to load. Please check your internet connection and try again.');
        return;
    }
    
    // Create a new workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('TimeSheet');
    
    // Generate days for the month
    const days = getDaysInMonth(year, month);
    
    let currentRow = 1;
    
    // Add 1 line gap before title
    currentRow += 1;
    
    // Title row - centered
    worksheet.mergeCells(currentRow, 1, currentRow, 4);
    const titleCell = worksheet.getCell(currentRow, 1);
    titleCell.value = 'Staff TimeSheet';
    titleCell.font = { bold: true, size: 16 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    currentRow += 1; // 1 line gap after title
    
    // Input information section
    const inputStartRow = currentRow;
    
    // Staff Name row - merge B and C
    const staffNameLabelCell = worksheet.getCell(currentRow, 1);
    staffNameLabelCell.value = 'Staff Name:';
    staffNameLabelCell.font = { bold: true };
    worksheet.mergeCells(currentRow, 2, currentRow, 3);
    const staffNameValueCell = worksheet.getCell(currentRow, 2);
    staffNameValueCell.value = staffName;
    staffNameValueCell.alignment = { horizontal: 'left', vertical: 'middle' };
    currentRow++;
    
    // Company Name row - merge B and C
    const companyNameLabelCell = worksheet.getCell(currentRow, 1);
    companyNameLabelCell.value = 'Company Name:';
    companyNameLabelCell.font = { bold: true };
    worksheet.mergeCells(currentRow, 2, currentRow, 3);
    const companyNameValueCell = worksheet.getCell(currentRow, 2);
    companyNameValueCell.value = companyName;
    companyNameValueCell.alignment = { horizontal: 'left', vertical: 'middle' };
    currentRow++;
    
    // Client Name row - merge B and C
    const clientNameLabelCell = worksheet.getCell(currentRow, 1);
    clientNameLabelCell.value = 'Client Name:';
    clientNameLabelCell.font = { bold: true };
    worksheet.mergeCells(currentRow, 2, currentRow, 3);
    const clientNameValueCell = worksheet.getCell(currentRow, 2);
    clientNameValueCell.value = clientName;
    clientNameValueCell.alignment = { horizontal: 'left', vertical: 'middle' };
    currentRow++;
    
    // Month row - merge B and C
    const monthLabelCell = worksheet.getCell(currentRow, 1);
    monthLabelCell.value = 'Month:';
    monthLabelCell.font = { bold: true };
    worksheet.mergeCells(currentRow, 2, currentRow, 3);
    const monthValueCell = worksheet.getCell(currentRow, 2);
    monthValueCell.value = monthNames[month] + ' ' + year;
    monthValueCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    const inputEndRow = currentRow;
    
    // Add borders to input section (excluding column D) - only to input rows, not the gap
    for (let row = inputStartRow; row <= inputEndRow; row++) {
        ['A', 'B', 'C'].forEach(col => {
            const cell = worksheet.getCell(col + row);
            if (!cell.border) cell.border = {};
            cell.border.top = { style: 'thin' };
            cell.border.bottom = { style: 'thin' };
            cell.border.left = { style: 'thin' };
            cell.border.right = { style: 'thin' };
        });
    }
    
    // Add 1 line gap after Month row (without borders)
    currentRow += 1;
    
    // Get Singapore public holidays for the year
    const publicHolidays = getSingaporePublicHolidays(year);
    
    // Table header
    const headerRow = currentRow;
    worksheet.getCell(currentRow, 1).value = 'Day of Month';
    worksheet.getCell(currentRow, 2).value = 'Day of Week';
    worksheet.getCell(currentRow, 3).value = 'Time Charged (Days)';
    worksheet.getCell(currentRow, 4).value = 'Engagement';
    
    // Style header row with center alignment
    ['A', 'B', 'C', 'D'].forEach(col => {
        const cell = worksheet.getCell(col + currentRow);
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    
    currentRow++;
    const dataStartRow = currentRow;
    
    // Generate rows for each day
    days.forEach(day => {
        const date = new Date(year, month, day);
        const dayOfWeek = dayNames[date.getDay()];
        const isWeekend = date.getDay() === 0 || date.getDay() === 6;
        const isPublicHolidayDay = isPublicHoliday(date, publicHolidays);
        
        // Day of Month
        const dayCell = worksheet.getCell(currentRow, 1);
        dayCell.value = day;
        dayCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Day of Week
        const dayOfWeekCell = worksheet.getCell(currentRow, 2);
        dayOfWeekCell.value = dayOfWeek;
        dayOfWeekCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Time Charged (Days)
        const timeChargedCell = worksheet.getCell(currentRow, 3);
        timeChargedCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Engagement
        const engagementCell = worksheet.getCell(currentRow, 4);
        engagementCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Set default value "1" for weekdays (not weekends or public holidays)
        if (!isWeekend && !isPublicHolidayDay) {
            timeChargedCell.value = '1';
            // Set default Engagement value for weekdays
            engagementCell.value = 'Onsite Changi Business Park';
        }
        
        // Add borders and styling to all cells
        ['A', 'B', 'C', 'D'].forEach(col => {
            const cell = worksheet.getCell(col + currentRow);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
        
        // Highlight public holidays in yellow (all public holidays, regardless of weekday/weekend)
        if (isPublicHolidayDay) {
            // All public holidays highlighted in yellow
            ['A', 'B', 'C', 'D'].forEach(col => {
                const cell = worksheet.getCell(col + currentRow);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFF00' } // Yellow
                };
            });
            // Time Charged should be blank for public holidays
            timeChargedCell.value = '';
        } else if (isWeekend) {
            // Regular weekend (not a public holiday) - highlight in yellow
            ['A', 'B', 'C', 'D'].forEach(col => {
                const cell = worksheet.getCell(col + currentRow);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFF00' } // Yellow
                };
            });
        }
        
        // Add data validation for Time Charged (for all days except weekends and public holidays)
        if (!isWeekend && !isPublicHolidayDay) {
            timeChargedCell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`"${timeChargedOptions.join('","')}"`]
            };
        }
        
        // Add data validation for Engagement (for all days except weekends and public holidays)
        if (!isWeekend && !isPublicHolidayDay) {
            engagementCell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`"${engagementOptions.join('","')}"`]
            };
        }
        
        currentRow++;
    });
    
    const dataEndRow = currentRow - 1;
    currentRow += 1; // Reduced spacing
    
    // Total row
    const totalRow = currentRow;
    const totalLabelCell = worksheet.getCell(currentRow, 1);
    totalLabelCell.value = 'Total Time Charged (Days):';
    totalLabelCell.font = { bold: true };
    totalLabelCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Formula to count cells that contain "1" in Time Charged column
    // This counts only working days (where Time Charged = "1")
    const totalFormula = `=COUNTIF(C${dataStartRow}:C${dataEndRow},"1")`;
    const totalValueCell = worksheet.getCell(currentRow, 2);
    totalValueCell.value = { formula: totalFormula };
    totalValueCell.font = { bold: true };
    totalValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
    
    currentRow += 2; // Reduced spacing
    
    // Signature section
    // Format current date as DD-MMM-YYYY (e.g., 08-Nov-2025)
    const currentDate = new Date();
    const day = String(currentDate.getDate()).padStart(2, '0');
    const monthNamesShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthShort = monthNamesShort[currentDate.getMonth()];
    const yearFull = currentDate.getFullYear();
    const dateValue = `${day}-${monthShort}-${yearFull}`;
    const formattedDate = `Date : ${dateValue}`;
    
    // Staff Name and Signature row
    const staffRow = currentRow;
    const staffSignatureLabelCell = worksheet.getCell(currentRow, 1);
    staffSignatureLabelCell.value = 'Staff Name and Signature:';
    staffSignatureLabelCell.font = { bold: true };
    
    // Merge columns B and C for staff name
    worksheet.mergeCells(staffRow, 2, staffRow, 3);
    const staffNameCell = worksheet.getCell(currentRow, 2);
    staffNameCell.value = staffName;
    staffNameCell.alignment = { horizontal: 'left', vertical: 'middle' };
    staffNameCell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
    };
    
    // Date cell with bold "Date :" part
    const staffDateCell = worksheet.getCell(currentRow, 4);
    staffDateCell.value = { richText: [
        { text: 'Date : ', font: { bold: true } },
        { text: dateValue }
    ]};
    
    // Add borders to individual cells in Staff row (columns A, D)
    ['A', 'D'].forEach(col => {
        const cell = worksheet.getCell(col + staffRow);
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    
    currentRow++;
    
    // Add 3 blank lines
    currentRow += 3;
    
    // Client Name and Signature row
    const clientRow = currentRow;
    const clientSignatureLabelCell = worksheet.getCell(currentRow, 1);
    clientSignatureLabelCell.value = 'Client Name and Signature:';
    clientSignatureLabelCell.font = { bold: true };
    worksheet.getCell(currentRow, 2).value = clientName;
    
    // Date cell with bold "Date :" part
    const clientDateCell = worksheet.getCell(currentRow, 4);
    clientDateCell.value = { richText: [
        { text: 'Date : ', font: { bold: true } },
        { text: dateValue }
    ]};
    
    // Add borders to individual cells in Client row (columns A, B, D)
    ['A', 'B', 'D'].forEach(col => {
        const cell = worksheet.getCell(col + clientRow);
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    
    // Set column widths (optimized for one page)
    worksheet.getColumn(1).width = 25; // Day of Month
    worksheet.getColumn(2).width = 15; // Day of Week
    worksheet.getColumn(3).width = 20; // Time Charged
    worksheet.getColumn(4).width = 28; // Engagement
    worksheet.getColumn(5).width = 18; // Date column
    
    // Set page setup to fit on one page
    worksheet.pageSetup = {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1, // Fit to 1 page height
        orientation: 'landscape',
        paperSize: 9, // A4
        margins: {
            left: 0.25,
            right: 0.25,
            top: 0.25,
            bottom: 0.25,
            header: 0.1,
            footer: 0.1
        },
        printTitlesRow: '1:1' // Repeat header row
    };
    
    // Set print area to ensure everything is included
    const lastRow = currentRow;
    worksheet.pageSetup.printArea = `A1:E${lastRow}`;
    
    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `TimeSheet_${staffName.replace(/\s+/g, '_')}_${monthNames[month]}_${year}.xlsx`;
    link.click();
    window.URL.revokeObjectURL(url);
}

function getDaysInMonth(year, month) {
    const days = [];
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    for (let i = 1; i <= daysInMonth; i++) {
        days.push(i);
    }
    
    return days;
}

// Get Singapore public holidays for a given year
function getSingaporePublicHolidays(year) {
    const holidays = [];
    
    // Fixed holidays
    holidays.push(new Date(year, 0, 1)); // New Year's Day
    holidays.push(new Date(year, 4, 1)); // Labour Day
    holidays.push(new Date(year, 7, 9)); // National Day
    holidays.push(new Date(year, 11, 25)); // Christmas
    
    // Variable holidays - approximate dates (these vary by year)
    // Chinese New Year (2 days) - usually late January or early February
    // Good Friday - varies
    // Vesak Day - varies
    // Hari Raya Puasa - varies
    // Hari Raya Haji - varies
    // Deepavali - varies
    
    // Singapore public holidays lookup table
    // Format: [month, day] where month is 0-indexed (0=Jan, 11=Dec)
    const holidayLookup = {
        2024: [
            [0, 1],      // New Year's Day
            [1, 10], [1, 11], [1, 12], // Chinese New Year (3 days including substitute)
            [2, 29],    // Good Friday
            [4, 1],     // Labour Day
            [4, 22],    // Vesak Day
            [4, 10],    // Hari Raya Puasa
            [5, 17],    // Hari Raya Haji
            [7, 9],     // National Day
            [10, 31],   // Deepavali
            [11, 25]    // Christmas
        ],
        2025: [
            [0, 1],     // New Year's Day
            [0, 29], [0, 30], [0, 31], // Chinese New Year (3 days)
            [3, 18],    // Good Friday
            [4, 1],     // Labour Day
            [4, 12],    // Vesak Day
            [4, 31],    // Hari Raya Puasa
            [5, 7],     // Hari Raya Haji
            [7, 9],     // National Day
            [10, 20],   // Deepavali
            [11, 25]    // Christmas
        ],
        2026: [
            [0, 1],     // New Year's Day
            [1, 16], [1, 17], [1, 18], // Chinese New Year (3 days)
            [3, 3],     // Good Friday
            [4, 1],     // Labour Day
            [5, 1],     // Vesak Day
            [5, 20],    // Hari Raya Puasa
            [5, 27],    // Hari Raya Haji
            [7, 9],     // National Day
            [11, 8],    // Deepavali
            [11, 25]    // Christmas
        ],
        2027: [
            [0, 1],     // New Year's Day
            [1, 6], [1, 7], [1, 8], // Chinese New Year (3 days)
            [3, 26],    // Good Friday
            [4, 1],     // Labour Day
            [4, 21],    // Vesak Day
            [5, 10],    // Hari Raya Puasa
            [5, 17],    // Hari Raya Haji
            [7, 9],     // National Day
            [10, 28],   // Deepavali
            [11, 25]    // Christmas
        ],
        2028: [
            [0, 1],     // New Year's Day
            [1, 26], [1, 27], [1, 28], // Chinese New Year (3 days)
            [4, 14],    // Good Friday
            [4, 1],     // Labour Day
            [5, 9],     // Vesak Day
            [4, 29],    // Hari Raya Puasa
            [5, 5],     // Hari Raya Haji
            [7, 9],     // National Day
            [11, 16],   // Deepavali
            [11, 25]    // Christmas
        ]
    };
    
    if (holidayLookup[year]) {
        return holidayLookup[year].map(([month, day]) => new Date(year, month, day));
    }
    
    // Fallback: return fixed holidays if year not in lookup
    return holidays;
}

// Check if a date is a Singapore public holiday
function isPublicHoliday(date, publicHolidays) {
    const dateStr = date.toDateString();
    return publicHolidays.some(holiday => holiday.toDateString() === dateStr);
}

