const PROMPT = `이 이미지는 기아자동차 엔진 가공 품질 검사 성적서(Inspection Report)입니다.

이미지에서 모든 측정 데이터를 추출해주세요.

반드시 아래 JSON 형식으로만 응답하세요. 설명이나 마크다운 없이 JSON만 출력하세요:

{
  "품명": "세타3 BLOCK GDI 2.5",
  "rows": [
    {
      "type": "가공 유형 (예: RH BOSS MILLING, LH BOSS MILLING, D/SHAFT MTG HOLE DR 등)",
      "no": "측정 포인트 번호 (E1, E3, E8, N1 등)",
      "item": "측정 축 (X, Y, Z, D, Y/X, Z/X 중 하나)",
      "actual": 실측값(숫자),
      "nominal": 기준값(숫자),
      "upTol": 상한 공차(숫자),
      "lowTol": 하한 공차(숫자),
      "deviation": 편차값(숫자),
      "evaluation": "평가 문자열",
      "ngOk": "NG 또는 OK"
    }
  ]
}

규칙:
- "No." 열의 ◆E1~E2>, ◆E3> 등에서 포인트 번호만 추출하세요.
- 같은 포인트에 여러 축 측정이 있으면 각각 별도 행으로 만드세요.
- "type"은 섹션 헤더(=== RH BOSS MILLING === 등)에서 가져오세요.
- 빨간색 글씨 = NG입니다.
- 숫자는 소수점 포함 정확히 기록하세요.`;

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { image, mimeType } = req.body;
  if (!image) return res.status(400).json({ error: "image required" });

  try {
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${process.env.GEMINI_API_KEY}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{
            parts: [
              { inline_data: { mime_type: mimeType || "image/png", data: image } },
              { text: PROMPT },
            ],
          }],
          generationConfig: { temperature: 0 },
        }),
      }
    );

    const data = await response.json();
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    const json = JSON.parse(text.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim());
    res.status(200).json(json);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
