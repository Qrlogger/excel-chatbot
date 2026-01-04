let excelText = "";

document.getElementById("fileInput").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
        const workbook = XLSX.read(evt.target.result, { type: "binary" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        // Convert to readable text for Gemini
        excelText = json.map((row, i) => `Row ${i + 1}: ${JSON.stringify(row)}`).join("\n");

        addBot("Excel uploaded successfully. Ask me anything about it.");
    };
    reader.readAsBinaryString(file);
});

async function askGemini() {
    const apiKey = document.getElementById("apiKey").value.trim();
    const question = document.getElementById("question").value.trim();

    if (!apiKey) {
        alert("Please paste your Gemini API key.");
        return;
    }
    if (!excelText) {
        alert("Please upload an Excel file first.");
        return;
    }
    if (!question) return;

    addUser(question);
    document.getElementById("question").value = "";

    const prompt = `
You are a senior QA analyst and product reviewer.

Below is Excel data uploaded by the user (bug reports / test cases / analysis sheet).

Analyze ALL rows carefully and answer the user's question clearly and professionally.
If the question asks for lists, return structured bullet points.
If it asks for expected/actual results, extract them accurately.

Excel Data:
${excelText}

User Question:
${question}
`;

    try {
        const response = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [
                        {
                            parts: [{ text: prompt }]
                        }
                    ]
                })
            }
        );

        const data = await response.json();
        const answer =
            data?.candidates?.[0]?.content?.parts?.[0]?.text ||
            "No response received from Gemini.";

        addBot(answer.replace(/\n/g, "<br>"));
    } catch (err) {
        addBot("Error talking to Gemini API.");
        console.error(err);
    }
}

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
