# Audio architecture (legacy VB6 client)

This is a quick map of the *current* runtime audio split implemented in `clsAudioEngine`.

## Backend matrix

- **Music OGG**: BASS stream path (`BassBackend_PlayOgg`).
- **Music MIDI**: DirectMusic path (`PlayMidi` / `LoadMidi`).
- **SFX OGG**: BASS sample path (`PlayCompressedSfx`, cached by `GetCompressedSfxSample`).
- **SFX WAV**: DirectSound fallback path (`CreateWavBufferFromFile` + duplicated buffers).
- **Ambient**: legacy DirectSound/WAV path (`PlayAmbient`).

## Entry-point intent

- Use `PlayMusic(filename, ...)` for new music callers.
  - `.ogg` selects BASS stream.
  - `.mid` selects DirectMusic.
  - Same-track requests are idempotent and only refresh volume.
- Use `PlayWav(id, ...)` for SFX callers.
  - It tries `id.ogg` first.
  - On miss/failure it falls back to WAV/DirectSound.
- `PlayMidi(id, ...)` remains for legacy call sites that already work with MIDI ids.

## Why both BASS streams and BASS samples exist

- Streams are used for long-lived, single-instance music playback.
- Samples are used for short SFX so one cached sample can spawn overlapping channels.

## Lifecycle notes

- Compressed SFX sample handles are long-lived caches for runtime reuse.
- They are released at shutdown (`FreeCompressedSfxCache`) before `ShutdownBassAudio`.
