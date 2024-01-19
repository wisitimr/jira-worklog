const dotenv = require("dotenv");
dotenv.config({ path: 'config/config.env' });

const XLSX = require("xlsx");
const dateFns = require("date-fns");
const fs = require("fs");

/* .env file
IN_DIR=/Users/zumo/Downloads
OUT_DIR=export
FILENAME=worklogs_2023-12-01_2023-12-31-2.xlsx
*/

(() => {
    try {
        const fileName = process.env.FILENAME;
        const workbookExcel = XLSX.readFile(`${process.env.IN_DIR}/${fileName}`);
        const sheetName = workbookExcel.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbookExcel.Sheets[sheetName]);

        const worklogSheet = sheetData
            .sort((a, b) => {
                return new Date(a["Started at"]) - new Date(b["Started at"]);
            })
            .map((col) => {
                let timeSpent = col["Time Spent (seconds)"] / 3600;
                const column = new Object();
                column["Name"] = col["Author"];
                column["ISSUEKEY"] = col["Issue Key"];
                column["Subtask Name"] = col["Issue Summary"];
                column["Description"] = col["Comment"];
                column["Date"] = dateFns.format(new Date(col["Started at"]), "dd/MM/yyyy");
                column["Time Spent"] = parseFloat(timeSpent).toFixed(1) + "h";
                return column;
            });

        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(worklogSheet);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
        const dateArray = fileName.split("_")[1].split("-");
        const yearMonth = `${dateArray[0]}_${dateArray[1]}`;
        const fileNameExport = `Report JiraWorkLog - ${yearMonth}`;
        const workDir = `${process.env.OUT_DIR}/${yearMonth}`;
        if (!fs.existsSync(workDir)) {
            fs.mkdirSync(workDir, { recursive: true });
        }
        const exportFile = `${workDir}/${fileNameExport}.xlsx`;
        XLSX.writeFile(workbook, exportFile);
    } catch (error) {
        console.error(error);
    }
})()
