document.addEventListener("DOMContentLoaded", loadDefaultFile);
document.getElementById("generatePromptsButton").addEventListener("click", generatePrompts);

let bibleText = "";

function loadDefaultFile() {
  fetch("TEXT-PCE.txt")
    .then((response) => response.text())
    .then((text) => {
      bibleText = text;
      document.getElementById("status").textContent = "Bible text loaded successfully.";
    })
    .catch((error) => {
      console.error("Error loaxding the file:", error);
      document.getElementById("status").textContent = "Error loading Bible text.";
    });
}

async function generatePrompts() {
  const promptTemplate = document.getElementById("promptTemplate").value;
  if (!promptTemplate) {
    alert("Please enter a prompt template.");
    return;
  }

  const statusDiv = document.getElementById("status");
  statusDiv.textContent = "Generating prompts...";

  const lines = bibleText.split("\n");
  let excelData = [["Book", "Verse", "Prompt", "URL", "Hyperlink"]];
  let currentBook = "";
  let currentChapter = "";

  const chunkSize = 1000;
  const totalChunks = Math.ceil(lines.length / chunkSize);

  for (let i = 0; i < lines.length; i += chunkSize) {
    const chunk = lines.slice(i, i + chunkSize);

    chunk.forEach((line) => {
      const match = line.match(/^(\w+)\s+(\d+):(\d+)\s+(.*)/);
      if (match) {
        const [, book, chapter, verse, text] = match;
        const fullBookName = bookNames[book] || book;
        const verseReference = `${chapter}:${verse} -`;
        const verseWithReference = `${verseReference} ${text}`;
        let prompt = promptTemplate.replace(/\[verse\]/g, `${fullBookName} ${chapter}:${verse}`);

        if (book !== currentBook || chapter !== currentChapter) {
          if (currentBook !== "") {
            excelData.push(["", "", "", "", ""]);
          }
          currentBook = book;
          currentChapter = chapter;
        }

        excelData.push([fullBookName, verseWithReference, prompt, "", `=HYPERLINK(D${excelData.length + 1},B${excelData.length + 1})`]);
      }
    });

    const progress = Math.round(((i + chunkSize) / lines.length) * 100);
    statusDiv.textContent = `Generating prompts... ${progress}% complete`;

    await new Promise((resolve) => setTimeout(resolve, 0));
  }

  statusDiv.textContent = "Creating Excel file...";
  await new Promise((resolve) => setTimeout(resolve, 0));

  createExcelFile(excelData);
  statusDiv.textContent = "Excel file with prompts generated and downloaded.";
}

function createExcelFile(data) {
  if (typeof XLSX === "undefined") {
    console.error("XLSX library is not loaded");
    return;
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);

  // Apply rich text formatting
  for (let R = 1; R < data.length; R++) {
    for (let C = 2; C <= 3; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cellValue = data[R][C];
      if (cellValue) {
        ws[cellAddress].r = formatRichText(cellValue);
      }
    }
  }

  ws["!cols"] = [
    { wch: 15 }, // Book
    { wch: 60 }, // Verse
    { wch: 80 }, // Prompt
    { wch: 30 }, // URL
    { wch: 60 }, // Hyperlink
  ];

  // Enable text wrapping for all cells
  for (let R = 0; R < data.length; R++) {
    for (let C = 0; C < data[R].length; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellAddress]) ws[cellAddress] = {};
      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s.alignment = { wrapText: true, vertical: "top" };

      // Set the formula for the Hyperlink column
      if (C === 4 && R > 0) {
        ws[cellAddress].f = data[R][C];
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "Bible Verses with Prompts");
  XLSX.writeFile(wb, "bible_verses_with_prompts.xlsx");
}

function formatRichText(text) {
  const parts = text.split(/(\d+:\d+ -)/);
  return parts.map((part, index) => {
    if (index % 2 === 1) {
      return {
        r: part,
        s: { bold: true },
      };
    }
    return { r: part };
  });
}

const bookNames = {
  Ge: "Genesis",
  Ex: "Exodus",
  Le: "Leviticus",
  Nu: "Numbers",
  De: "Deuteronomy",
  Jos: "Joshua",
  Jg: "Judges",
  Ru: "Ruth",
  "1Sa": "1 Samuel",
  "2Sa": "2 Samuel",
  "1Ki": "1 Kings",
  "2Ki": "2 Kings",
  "1Ch": "1 Chronicles",
  "2Ch": "2 Chronicles",
  Ezr: "Ezra",
  Ne: "Nehemiah",
  Es: "Esther",
  Job: "Job",
  Ps: "Psalms",
  Pr: "Proverbs",
  Ec: "Ecclesiastes",
  Ca: "Song of Solomon",
  Isa: "Isaiah",
  Jer: "Jeremiah",
  La: "Lamentations",
  Eze: "Ezekiel",
  Da: "Daniel",
  Ho: "Hosea",
  Joe: "Joel",
  Am: "Amos",
  Ob: "Obadiah",
  Jon: "Jonah",
  Mic: "Micah",
  Na: "Nahum",
  Hab: "Habakkuk",
  Zep: "Zephaniah",
  Hag: "Haggai",
  Zec: "Zechariah",
  Mal: "Malachi",
  Mt: "Matthew",
  Mr: "Mark",
  Lu: "Luke",
  Joh: "John",
  Ac: "Acts",
  Ro: "Romans",
  "1Co": "1 Corinthians",
  "2Co": "2 Corinthians",
  Ga: "Galatians",
  Eph: "Ephesians",
  Php: "Philippians",
  Col: "Colossians",
  "1Th": "1 Thessalonians",
  "2Th": "2 Thessalonians",
  "1Ti": "1 Timothy",
  "2Ti": "2 Timothy",
  Tit: "Titus",
  Phm: "Philemon",
  Heb: "Hebrews",
  Jas: "James",
  "1Pe": "1 Peter",
  "2Pe": "2 Peter",
  "1Jo": "1 John",
  "2Jo": "2 John",
  "3Jo": "3 John",
  Jude: "Jude",
  Re: "Revelation",
};
