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

    if (!apiKey) return alert("Paste Gemini API key");
    if (!excelText) return alert("Upload Excel first");
    if (!question) return;

    addUser(question);
    document.getElementById("question").value = "";

    const prompt = `
You are a senior QA analyst.

Analyze the Excel data below and answer clearly.

Excel Data:
${excelText.slice(0, 12000)}

Question:
${question}
`;

    try {
        const res = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [
                        {
                            role: "user",
                            parts: [{ text: prompt }]
                        }
                    ]
                })
            }
        );

        const data = await res.json();
        console.log("Gemini raw response:", data); // 🔴 DEBUG

        if (data.error) {
            addBot("Gemini API Error: " + data.error.message);
            return;
        }

        const answer = data.candidates?.[0]?.content?.parts?.[0]?.text;

        if (!answer) {
            addBot("Gemini returned empty response. Try a smaller Excel or simpler question.");
            return;
        }

        addBot(answer.replace(/\n/g, "<br>"));
    } catch (err) {
        console.error(err);
        addBot("Network or API failure. Check console.");
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
