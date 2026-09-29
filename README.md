Sanwad (संवाद) — Voice in. Text across. Voice out.
Smart India Hackathon 2026 · Problem Statement 26173 · ISRO Space Applications Centre
A fully offline, open-source emergency walkie-talkie for Android with local Speech-to-Text (STT) and Text-to-Speech (TTS) in 10 Indian languages — built for the links where audio simply cannot fit.
> When the network dies, voice is the last link. Sanwad turns speech into ~100 bytes of text, streams it over Wi-Fi Direct / Bluetooth (and later LoRa / NavIC-ecosystem / HF-VHF packet links), and turns it back into speech on the other phone.
---
Try the interactive prototype
Open `prototype/index.html` in any browser — no build step, no internet needed (single self-contained file).
Demo flow (2 minutes):
Phone A (left): Get started → allow permissions → enter a display name → pick a language
The two phones link over Wi-Fi Direct automatically
Hold the mic on Phone A and speak — Sanwad detects the pause, forms the sentence, and streams it as ~100 B of text (watch the packet chip + packet log — audio never crosses the link)
Phone B receives it and plays it as a voice note
Tick ⚠ Send as alert and talk again — Phone B announces it at highest volume, non-interruptible
Toggle Walkie → Phone to see "off = normal phone" (plain text messaging)
Tap ⓘ on either phone for the on-device architecture & latency budget
Demo helpers
"Demo: tap to talk" checkbox on each phone — one tap sends (no need to hold). Ideal for live judge demos; off by default to show true push-to-talk.
Deep links (open directly on a state): `index.html#shot=linked` · `#shot=voice` · `#shot=alert` · `#shot=phone` · `#demo` (auto-plays the loop after onboarding)
![Voice note loop](screenshots/03-voice-note.png)
	
![Onboarding](screenshots/01-onboarding.png)	![Alert — max volume, non-interruptible](screenshots/04-alert.png)
Onboarding: permissions, name, 10-language picker	Alert mode: highest volume, mute bypassed


![Phone mode](screenshots/05-phone-mode.png)
Walkie off = normal phone + on-device architecture sheet
---
How it maps to the problem statement
PS requirement	Where it is
Lightweight, accurate STT + TTS for 10 Indian languages (hi, gu, mr, kn, ml, ta, te, or, bn, en), running locally	On-device STT/TTS; ⓘ sheet on each phone
STT detects pauses & stoppages, forms sentences	Hold-to-talk flow: "detecting pause" → "forming sentence"
Instantly stream via Wi-Fi/Bluetooth to embedded device or another phone, minimal latency	Wi-Fi Direct / BLE transport; packet log with byte counts; latency budget ≤ 4–5 s mouth-to-ear
TTS converts received text to intelligible speech, played as voice note	Phone B playback with play/progress controls
Alert messages at highest volume, non-interruptible	⚠ alert flow — full-screen, mute bypassed
Two phones with same app, one TTS mode / one STT mode, push-to-talk walkie-talkie	The live loop in the prototype
Turned off = works like a phone	Walkie/Phone toggle
Architecture (as intended for the on-device build)
STT: IndicConformer (AI4Bharat) — int8 TFLite, on-device, < 300 MB, 10 languages; Whisper.cpp as fallback
TTS: espeak-ng baseline (all 10 languages, fully offline) + neural Piper (en/hi) for high-frequency alert phrases — speed + intelligibility prioritised over studio fidelity
Transport: Wi-Fi Direct / BLE today; same text-first design ports to LoRa, NavIC-ecosystem messaging, HF/VHF packet without re-architecture
Efficiency targets: 2 GB RAM footprint, ₹8–10k (Cortex-A53 class) devices, < 5% idle CPU, ₹0 recurring cloud/cellular cost
Latency budget: STT 1.5–2.5 s · link < 0.5 s · TTS < 1.5 s → ≤ 4–5 s mouth-to-ear, RTF < 0.5
Open source only (Apache-2.0 / MIT), 100% offline
Repository layout
```
sanwad/
├── prototype/
│   └── index.html          # interactive 2-phone prototype (open in browser)
├── pitch-deck/
│   └── Sanwad_PS26173_SIHS2026_Deck.pptx   # 7-slide pitch deck
├── screenshots/            # verified prototype captures
├── tools/
│   └── build_deck.py       # regenerates the pitch deck (pip install python-pptx)
├── LICENSE
└── README.md
```
Honest note (prototype)
The in-browser prototype simulates STT/TTS and the phone-to-phone transport (with realistic ~0.5 s link delay). Where the browser supports it, live Web Speech APIs are used automatically — real microphone STT on the sender and real TTS voices on the receiver. All model names, footprints and latency figures on the ⓘ sheet reflect the intended on-device architecture.
Team
	
Team name: (to be filled)	Team ID: (to be filled)
---
Sanwad — zero recurring cloud or cellular cost. Words travel even when data can't.
