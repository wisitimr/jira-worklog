const dotenv = require("dotenv");
dotenv.config();

const XLSX = require("xlsx");
const fs = require("fs");
const _ = require("lodash");
const moment = require("moment");

(() => {
    try {
        const fileName = process.env.FILE_NAME;
        const workbookExcel = XLSX.readFile(`${process.env.WORKDIR_IN}/${fileName}`);
        const sheetName = workbookExcel.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbookExcel.Sheets[sheetName]);

        const data = sheetData
            .sort((a, b) => {
                return new Date(a["Started at"]) - new Date(b["Started at"]);
            })
            .map((col) => {
                return [col["Parent Summary"].replace("[Payment & Billing] Enhancement and Improvement", "[MyAIS] Backend"), col["Comment"]].join(" : ");
            });

        const date = fileName.split("_")[1];
        console.log(moment(date).format('MMMM YYYY'));
        const workDir = `${process.env.WORKDIR_OUT}`;
        if (!fs.existsSync(workDir)) {
            fs.mkdirSync(workDir, { recursive: true });
        }
        console.log([
            "##################################################################",
            moment(date).format('MMMM YYYY'),
            "##################################################################",
            ..._.uniq(data)].join('\n'));
    } catch (error) {
        console.error(error);
    }
})()
