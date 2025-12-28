import xlsx from "xlsx";

export function parseStudentsExcel(buffer) {
  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  return xlsx.utils.sheet_to_json(sheet, {
    defval: "",   // 🔑 missing cells → ""
    raw: false,
  });
}
