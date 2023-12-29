// Import required modules
import XLSX from "xlsx";
import path from "path";
import fs from "fs";
import dotenv from "dotenv";

// Check if .env file exists, create one if not
if (!fs.existsSync("./.env"))
  fs.writeFileSync(
    "./.env",
    `# COPY THE ID KEGIATAN & BEARER TOKEN FROM KAMPUS MERDEKA AFTER YOU LOGIN!\n# https://kampusmerdeka.kemdikbud.go.id/activity/status\n\nIDKEGIATAN=\nBEARERTOKEN=`
  );

// Load environment variables from .env file
dotenv.config();

// Retrieve values from environment variables or use default values
const idKegiatan = process.env.IDKEGIATAN || "";
const bearerToken = process.env.BEARERTOKEN || "";

/**
 * Export data to an Excel file.
 *
 * @param {object} dataJSON - JSON data to be exported.
 * @returns {boolean} - Returns true on successful export.
 */
const exportToExcel = (dataJSON) => {
  try {
    // Prepare data for Excel export
    const dataExcel = dataJSON.data.reverse().flatMap((weekly) => {
      let previousValues = [];
      return weekly.daily_report.map((report) => {
        // Format date for report and weekly counter
        const convertDateReport = `Minggu ke-${weekly.counter} | ${new Date(
          weekly.start_date
        ).getDate()} - ${new Date(weekly.end_date).toLocaleDateString("id-ID", {
          day: "numeric",
          month: "long",
          year: "numeric",
        })}`;
        const currentValues = [
          weekly.counter,
          convertDateReport,
          weekly.learned_weekly,
          weekly.reviewed_name,
        ];

        // Check if current values are the same as the previous, nullify if true
        const shouldNullify = currentValues.every(
          (value, index) => value === previousValues[index]
        );
        previousValues = shouldNullify ? currentValues : currentValues;

        // Format report date
        const reportDate = new Date(report.report_date).toLocaleDateString(
          "id-ID",
          { weekday: "long", day: "numeric", month: "long", year: "numeric" }
        );

        // Return formatted data for Excel export
        return shouldNullify
          ? [null, null, null, null, reportDate, report.report]
          : [...currentValues, reportDate, report.report];
      });
    });

    // Set column names for Excel sheet
    const workSheetColumnNames = [
      "No",
      "Tanggal Laporan",
      "Ringkasan Laporan Mingguan",
      "Diterima Oleh",
      "Tanggal Harian",
      "Laporan Harian",
    ];

    // Create Excel workbook and worksheet
    const workBook = XLSX.utils.book_new();
    const workSheetData = [workSheetColumnNames, ...dataExcel];
    const worksheet = XLSX.utils.aoa_to_sheet(workSheetData);

    // Set column widths based on content length
    const colsWidth = Array.from(
      { length: workSheetColumnNames.length },
      (_, index) => {
        let count =
          dataExcel
            .map((data) => data[index])
            .filter((value) => value !== null)
            .reduce((max, str) => Math.max(max, str.length), 0) + 3;

        // Limit column width to 200 characters
        if (count >= 200) {
          count = 200;
        } else if (isNaN(count)) {
          count = 3;
        }

        // Set width for each column
        const data = {
          wch: count,
        };
        return data;
      }
    );

    // Apply column widths to the worksheet
    worksheet["!cols"] = colsWidth;

    // Set merged cells for a cleaner look
    const merge = Array.from(
      { length: (dataExcel.length / 5) * 4 },
      (_, index) => ({
        s: { r: Math.floor(index / 4) * 5 + 1, c: index % 4 },
        e: { r: Math.floor(index / 4) * 5 + 5, c: index % 4 },
      })
    );
    worksheet["!merges"] = merge;

    // Append worksheet to workbook
    XLSX.utils.book_append_sheet(
      workBook,
      worksheet,
      `Logbook MSIB ${dataJSON.data[0].id_reg_penawaran}`
    );

    // Write workbook to file
    XLSX.writeFile(
      workBook,
      path.resolve(`./Logbook MSIB ${dataJSON.data[0].id_reg_penawaran}.xlsx`)
    );
  } catch (error) {
    // Handle errors during API consumption or JSON file reading
    console.log(
      `ERROR OCCURS WHEN YOU WANT TO CONSUME API OR JSON FILE!\n\nPLEASE COPY THE RESPONSE JSON (file.json) OR ID KEGIATAN & BEARER TOKEN (.env) FROM PLATFROM KAMPUS MERDEKA AFTER LOGIN!\nhttps://kampusmerdeka.kemdikbud.go.id/activity/status\n\nCHECK THE ERROR OR ENV FILE OR FILE.JSON CONSTRAINT: ${error.message}`
    );
  }
  return true;
};

// Check if file.json exists, attempt to read and export data
if (fs.existsSync("./file.json")) {
  let checkJSON = false;
  try {
    const dataJSON = JSON.parse(fs.readFileSync("./file.json", "utf8"));
    exportToExcel(dataJSON);
  } catch (error) {
    // Set checkJSON to true if there's an error reading file.json
    checkJSON = true;
  }

  // If file.json is not valid or missing, attempt to fetch data from the API
  if (checkJSON) {
    try {
      fetch(
        `https://api.kampusmerdeka.kemdikbud.go.id/magang/report/allweeks/${idKegiatan}`,
        {
          method: "GET",
          headers: { Authorization: `Bearer ${bearerToken}` },
        }
      )
        .then(function (response) {
          return response.json();
        })
        .then(function (json) {
          exportToExcel(json);
        });
    } catch (error) {
      // Handle errors during API consumption
      console.log(
        `ERROR OCCURS WHEN YOU WANT TO CONSUME API OR JSON FILE!\n\nPLEASE COPY THE RESPONSE JSON (file.json) OR ID KEGIATAN & BEARER TOKEN (.env) FROM PLATFROM KAMPUS MERDEKA AFTER LOGIN!\nhttps://kampusmerdeka.kemdikbud.go.id/activity/status\n\nCHECK THE ERROR OR ENV FILE OR FILE.JSON CONSTRAINT: ${error.message}`
      );
    }
  }
} else {
  // Inform the user if file.json is missing
  console.log(
    "FILE MISSING PLEASE RESTART THE PROGRAM!\n\nPLEASE COPY THE RESPONSE JSON (file.json) OR ID KEGIATAN & BEARER TOKEN (.env) FROM PLATFROM KAMPUS MERDEKA AFTER LOGIN!\nhttps://kampusmerdeka.kemdikbud.go.id/activity/status\n\nPLEASE FILL THE ENV FILE OR FILE.JSON AND RETRY TYPE npm start"
  );
  fs.writeFileSync("./file.json", "");
  fs.writeFileSync(
    "./file.example.txt",
    `PLEASE COPY THE RESPONSE JSON FROM PLATFROM KAMPUS MERDEKA AFTER LOGIN!\nhttps://kampusmerdeka.kemdikbud.go.id/activity/status\n\nEXAMPLE JSON: \nPLEASE GO TO JSON BEUTIFY FOR SEE THE STRUCTURE JSON \nhttps://codebeautify.org/jsonviewer\n\n{"data":[{"id":"xxxx","position_id":"xxxx","uuid_akun":"xxxx","id_reg_penawaran":"xxxx","start_date":"xxxx","end_date":"xxxx","counter":"xxxx","status":"xxxx","learned_weekly":"xxxx","note_from_mentor":"xxxx","reviewed_by":"xxxx","reviewed_name":"xxxx","daily_report":[{"id":"xxxx","weekly_report_id":"xxxx","uuid_akun":"xxxx","id_reg_penawaran":"xxxx","report_date":"xxx","status":"xxxx","feeling_level":"xxxx","report":"xxxx","string_day":"xxx","counter":"xxxx","created_by":"xxxx","updated_by":"xxxx","created_at":"xxx","updated_at":"xxx"},{"id":"xxxx","weekly_report_id":"xxxx","uuid_akun":"xxxx","id_reg_penawaran":"xxxx","report_date":"xxx","status":"xxxx","feeling_level":"xxxx","report":"xxxx","string_day":"xxx","counter":"xxxx","created_by":"xxxx","updated_by":"xxxx","created_at":"xxx","updated_at":"xxx"},{"id":"xxxx","weekly_report_id":"xxxx","uuid_akun":"xxxx","id_reg_penawaran":"xxxx","report_date":"xxx","status":"xxxx","feeling_level":"xxxx","report":"xxxx","string_day":"xxx","counter":"xxxx","created_by":"xxxx","updated_by":"xxxx","created_at":"xxx","updated_at":"xxx"},{"id":"xxxx","weekly_report_id":"xxxx","uuid_akun":"xxxx","id_reg_penawaran":"xxxx","report_date":"xxx","status":"xxxx","feeling_level":"xxxx","report":"xxxx","string_day":"xxx","counter":"xxxx","created_by":"xxxx","updated_by":"xxxx","created_at":"xxx","updated_at":"xxx"},{"id":"xxxx","weekly_report_id":"xxxx","uuid_akun":"xxxx","id_reg_penawaran":"xxxx","report_date":"xxx","status":"xxxx","feeling_level":"xxxx","report":"xxxx","string_day":"xxx","counter":"xxxx","created_by":"xxxx","updated_by":"xxxx","created_at":"xxx","updated_at":"xxx"}],"mentor_notes":[],"updated_at":"xxxx"}],"meta":{"limit":"xxxx","offset":"xxxx","total":"xxxx"}}`
  );
}

console.log(`© ${new Date().getFullYear()} by Danny Jiustian`);