import { GoogleGenAI } from "https://esm.run/@google/genai";

let excelText = "";
let ai = null;

// Excel upload
document.getElementById("fileInput").addEventListener("change", (e) => {
  const reader = new FileReader();
  reader.onload = (evt) => {
    const workbook = XLSX.read(evt.target.result, { type: "binary" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    excelText = json
      .map((row, i) => `Row ${i + 1}: ${JSON.stringify(row)}`)
      .join("\n");

    addBot("Excel uploaded. Ask your question.");
  };
  reader.readAsBinaryString(e.target.files[0]);
});

// Ask Gemini
document.getElementById("askBtn").addEventListener("click", async () => {
  const apiKey = document.getElementById("apiKey").value.trim();
  const question = document.getElementById("question").value.trim();

  if (!apiKey) return alert("Paste Gemini API key");
  if (!excelText) return alert("Upload Excel first");
  if (!question) return;

  if (!ai) {
    ai = new GoogleGenAI({ apiKey });
  }

  addUser(question);
  document.getElementById("question").value = "";

  const prompt = `
You are a senior QA analyst.

Analyze the Excel data below and answer clearly and accurately.

Excel Data:
${excelText.slice(0, 12000)}

User Question:
${question}
`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash",
      contents: prompt,
    });

    addBot(response.text.replace(/\n/g, "<br>"));
  } catch (err) {
    console.error(err);
    addBot("Gemini SDK error. Check console.");
  }
});

// Chat helpers
function addUser(text) {
  document.getElementById("chat").innerHTML +=
    `<div class="msg user">You: ${text}</div>`;
  scrollChat();
}

function addBot(text) {
  document.getElementById("chat").innerHTML +=
    `<div class="msg bot">Gemini: ${text}</div>`;
  scrollChat();
}

function scrollChat() {
  const chat = document.getElementById("chat");
  chat.scrollTop = chat.scrollHeight;
}
