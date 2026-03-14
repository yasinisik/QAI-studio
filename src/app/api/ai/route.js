import { NextResponse } from 'next/server'

const SYSTEM_PROMPT = "Sen QA Studio'nun AI asistanisin. Dunyanin en iyi yazilim test muhendislerinden birisin. Uzmanlik alanların: test senaryosu tasarimi, test otomasyonu (Playwright, Selenium, pytest, Jest), API testleri (REST, GraphQL, Postman), performans testleri (JMeter, k6), guvenlik testleri (OWASP), CI/CD entegrasyonu (GitHub Actions), BDD/Gherkin, hata analizi ve root cause tespiti. Yanit verirken: Turkce yaz, sade ve anlasılır ol, somut ornekler ver, adim adim acikla, kod orneklerini kod bloklari icinde goster, kisa tut ama eksik bilgi verme."

export async function POST(req) {
  const { messages } = await req.json()
  const apiKey = process.env.GROQ_API_KEY

  if (!apiKey) {
    return NextResponse.json({ error: 'API anahtari eksik.' }, { status: 500 })
  }

  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + apiKey,
    },
    body: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      max_tokens: 2048,
      temperature: 0.7,
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        ...messages,
      ],
    }),
  })

  const data = await res.json()
  const text = data.choices?.[0]?.message?.content || 'Yanit alinamadi.'
  return NextResponse.json({ text })
}
