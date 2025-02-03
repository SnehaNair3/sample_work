const ExcelJS = require('exceljs');
const fs = require('fs');
const csv = require('csv-parser');

// Utility function to load Excel file
async function loadExcel(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1); // Assuming data is in the first worksheet

    const rows = [];
    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
        const rowValues = row.values.slice(1); // Remove the first empty cell
        rows.push(rowValues);
    });

    return rows;
}

// Process the data for reporting
function processData(data) {
    const byGender = data.reduce((acc, row) => {
        const gender = row.genderid === 'M' ? 'Male' : 'Female';
        acc[gender] = acc[gender] || [];
        acc[gender].push(row);
        return acc;
    }, {});

    const byUnit = data.reduce((acc, row) => {
        const unit = row.unitname;
        acc[unit] = acc[unit] || [];
        acc[unit].push(row);
        return acc;
    }, {});

    return { byGender, byUnit };
}

// Generate Excel report
async function generateReport(data, filePath) {
    const workbook = new ExcelJS.Workbook();
    const genderSheet = workbook.addWorksheet('By Gender');
    const unitSheet = workbook.addWorksheet('By Unit');

    // Add data by gender
    Object.entries(data.byGender).forEach(([gender, rows]) => {
        genderSheet.addRow([gender]);
        rows.forEach((row) => genderSheet.addRow(Object.values(row)));
        genderSheet.addRow([]); // Empty row for separation
    });

    // Add data by unit
    Object.entries(data.byUnit).forEach(([unit, rows]) => {
        unitSheet.addRow([unit]);
        rows.forEach((row) => unitSheet.addRow(Object.values(row)));
        unitSheet.addRow([]); // Empty row for separation
    });

    await workbook.xlsx.writeFile(filePath);
    console.log(`Report generated: ${filePath}`);
}

// Main function
async function main() {
    const inputFile = 'PersonalExcel.xlsx';  // Ensure this file exists
    const outputFile = 'employee_report.xlsx';

    try {
        const data = await loadExcel(inputFile);
        const processedData = processData(data);
        await generateReport(processedData, outputFile);
    } catch (err) {
        console.error('Error:', err);
    }
}

// Run the script
main();


// import pandas as pd

// # Sample Data for GSTR-2B and Go GST Purchase/Service Report
// gstr2b_data = {
//     "GSTIN": [
// "29AAICA3918J1ZE",
// "32AAGFL9125M1Z3",
// "32AACCC5591P1ZA",
// "32AACCE4671N1ZH",
// "32AAKFJ9784C1Z5",
// "32AANFK2015N1ZC",
// "32AAAAK5375M4ZD",
// "32AAAAK5375M4ZD",
// "32AAAAK5375M4ZD",
// "32AAAAK5375M4ZD",
// "32AADFK1480D1ZW",
// "32AAFCA6562R1Z8"
// "32ALFPP3798J1ZN",
// "32AARFR5296E1ZV",
// "32AARFR5296E1ZV",
// "32AARFR5296E1ZV"
// "32AARFR5296E1ZV",
// "32AARFR5296E1ZV",
// "32AARFR5296E1ZV",
// "32AAATM9488J1Z3",
// "32AAATM9488J1Z3",
// "32AAATM9488J1Z3",
// "32AAATM9488J1Z3",
// "32AACFO3414C1Z1"
// "32AIMPJ6311K1ZB",
// "32AIMPJ6311K1ZB",
// "32AIMPJ6311K1ZB",
// "32AIMPJ6311K1ZB",
// "32AIMPJ6311K1ZB",
// "32AABFL1849D1ZS",
// "33AACCE4671N1ZF",
// "32ABDFS9372D1Z7",
//     ],
// }




// go_gst_data = {
//     "GSTIN": [
// "32AFYPC1510J1ZO",
// "32AAAAK5375M3ZE",
// "32AFYPC1510J1ZO",
// "32ABDFS9372D1Z7",
// "32LEUPS1327D1Z9",
// "32AFYPC1510J1ZO",
// "32AAAAN3982H2ZM",
// "32AARFT8404A1ZD",
// "32AADFP4734B1ZS",
// "32AADFP4734B1ZS",
// "32AAACZ7977Q1ZD",
// "32AAFFR2352H1ZI",
// "32AAFFR2352H1ZI",
// "32AAHCC9798D1ZD",
// "32AAHCC9798D1ZD",
// "32AAFFR2352H1ZI",
// "32AAATM9488J1Z3",
// "32AABCB5576G5ZQ",
// "32AAAAM1011G2ZH",
// "32AAAAM1011G2ZH",
// "32AAAAM1011G2ZH",
// "32AAAAM1011G2ZH",
// "32AAAAM1011G2ZH",
// "32AAAAM1011G2ZH",
// "32AATFA9615Q1ZQ",
// "32AATFA9615Q1ZQ",
// "32AABCB5576G5ZQ",
// "32AAATM9488J1Z3",
// "32AABCK7612C1Z4",
// "32AABCK7612C1Z4",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32AABCP3805G1ZW",
// "32ALBPR0528A1ZU",
// "32ALBPR0528A1ZU",
// "32AAATM9488J1Z3",
// "32AACFV1489L1ZW",
// "32CBYPP8451P1ZI",
// "32AAFFV3113Q1Z2",
// "32ALBPR0528A1ZU",
// "32AAOFA3123E1Z1",
// "32AAOFA3123E1Z1",
// "32AANFP0514H1ZK",
// "32AAIFM5969L1ZQ",
// "32AAIFM5969L1ZQ",
// "32AACF03414C1Z1",
// "32AABCV3896K1ZY",
// "32AABCV3896K1ZY",
// "32ALFPP3798J1ZN",
// "32AAATM9488J1Z3",
// "32ADEPR9133J1ZE",
// "32AA ACC1206D2ZO",
// "32AAATM9488J1Z3",
// "32CCUPP6022R1ZU",
// "32BEPPG1387E2ZL",
// "32BEPPG1387E2ZL",
// "32CCUPP6022R1ZU",
// "32AADFP4734B1ZS",
// "32ALBPR0528A1ZU",
// "32ALBPR0528A1ZU",
// "32AHDPK2803L1ZM",
// "32AHDPK2803L1ZM",
// "32AACCG5704A1ZE",
// "32ARUPM6535D1ZN",
// "32AACCN6452C1ZW",
// "32AACCN6452C1ZW",
// "32AAFFR2352H1Z1",
// "32AAFFR2352H1Z1",
// "32ADUPV4996K1ZF",
// "32ALBPR0528A1ZU",
// "32ALBPR0528A1ZU",
// "32AABCV3896K1ZY",
// "32AABCV3896K1ZY",
// "32AAQFM3558B1ZE",
// "32AAKFD4334Q1Z6",
// "32ABQFA9908E1ZC",
// "32AAAAC9617L1ZP",
// "32AACF03414C1Z1",
// "32ALBPR0528A1ZU",
// "32AACF03414C1Z1",
// "32AABCB5576G5ZQ",
// "32AAAFE7091D1ZU",
// "32AABCB5576G5ZQ",
// "32AABFW8349K2ZS",
// "32AMZPJ8281H1ZG",
// "32AANFR6386H1Z2",
// "32AAUFN7453L1ZM",
// "32ADUPV4996K1ZF",
// "32EFIPK9998Q1Z6",
// "32AAAAK5375M3ZE",
// "32AAAAC9617L1ZP",

//     ]
// }

// # Convert to DataFrames
// gstr2b_df = pd.DataFrame(gstr2b_data)
// go_gst_df = pd.DataFrame(go_gst_data)

// # Merge DataFrames on GSTIN
// merged_df = pd.merge(
//     gstr2b_df, go_gst_df, on="GSTIN", how="outer", indicator=True
// )

// # Create a new column to indicate if GSTIN values are present in the first DataFrame
// merged_df['Present_in_1st_GSTIN'] = merged_df['GSTIN'].apply(
//     lambda x: 'Present' if x in gstr2b_df['GSTIN'].values else 'Not Present'
// )

// # Generate Reports
// # 1. Items reflected in both GSTR-2B and Go GST Report
// both_reflected = merged_df[merged_df["_merge"] == "both"]

// # 2. Items reflected in GSTR-2B but not in Go GST Report
// only_in_2b = merged_df[merged_df["_merge"] == "left_only"]

// # 3. Items not reflected in GSTR-2B but available in Go GST Report
// only_in_go_gst = merged_df[merged_df["_merge"] == "right_only"]

// # Save Reports to Excel
// with pd.ExcelWriter("GST_Comparison_Report.xlsx") as writer:
//     both_reflected.to_excel(writer, sheet_name="Both_Reflected", index=False)
//     only_in_2b.to_excel(writer, sheet_name="Only_in_2B", index=False)
//     only_in_go_gst.to_excel(writer, sheet_name="Only_in_GoGST", index=False)

// print("Reports generated successfully! Check 'GST_Comparison_Report.xlsx'.")