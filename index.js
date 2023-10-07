const readXlsxFile = require("read-excel-file/node");

const file = "sheet.xlsx";

let rows;

const getCleanField = (cell) => {
  let value = cell;
  value = value.replace(/GHG/g, "");
  value = value.replace(/  /g, " ");
  value = value.toLowerCase().trim();

  return value;
};

const checkKeyword = (cell, name) => {
  let value = getCleanField(cell);
  if (value.includes(name)) return true;
  return false;
};

const getCorrespondingValue = (sheet, i, j) => {
  const rightValue = rows[i][j + 1];
  const bottomValue = rows[i + 1][j];

  const validValue =
    typeof rightValue === "number"
      ? rightValue
      : typeof bottomValue === "number"
      ? bottomValue
      : false;

  return validValue;
};

const run = async () => {
  let sheetNames = await readXlsxFile.readSheetNames(file);

  for (let i = 0; i < sheetNames.length; i += 1) {
    const sheet = sheetNames[i];
    rows = await readXlsxFile(file, { sheet });

    let skipThisSheet = false;

    let fossilEmissions,
      biogenicEmissions,
      biogenicRemoval,
      airCraftEmissions,
      landEmissions;

    for (let j = 0; j < rows.length; j += 1) {
      const row = rows[j];

      for (let k = 0; k < row.length; k += 1) {
        let cell = row[k];

        if (!cell) continue;
        cell = String(cell);
        if (checkKeyword(cell, "fossil emissions")) {
          const value = getCorrespondingValue(sheet, j, k);
          fossilEmissions = value;
        } else if (checkKeyword(cell, "biogenic emissions")) {
          const value = getCorrespondingValue(sheet, j, k);
          biogenicEmissions = value;
        } else if (checkKeyword(cell, "biogenic removal")) {
          const value = getCorrespondingValue(sheet, j, k);
          biogenicRemoval = value;
        } else if (checkKeyword(cell, "air craft emissions")) {
          const value = getCorrespondingValue(sheet, j, k);
          airCraftEmissions = value;
        } else if (checkKeyword(cell, "emissions from land")) {
          const value = getCorrespondingValue(sheet, j, k);
          landEmissions = value;
        }

        if (
          fossilEmissions &&
          biogenicEmissions &&
          biogenicRemoval &&
          airCraftEmissions &&
          landEmissions
        ) {
          skipThisSheet = true;
          break;
        }
      }

      if (skipThisSheet) {
        break;
      }
    }

    skipThisSheet = false;

    if (fossilEmissions) {
      console.log(
        `------------------------------------------------------------
  Sheet: ${sheet}
  Fossil emissions: ${fossilEmissions}
  Biogenic emissions: ${biogenicEmissions}
  Biogenic removal: ${biogenicRemoval}
  Air craft emissions: ${airCraftEmissions}
  Emissions from land use change: ${landEmissions}`
      );
    }
  }
  console.log("------------------------------------------------------------");
};

run();
