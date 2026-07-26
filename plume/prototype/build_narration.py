#!/usr/bin/env python3
"""Narration v8 :
- voix homme = Tom « expressif » : prosodie plus variée (noise_w élevé),
  pitch abaissé modérément (×0.90) pour rester naturel ;
- studio des voix : échantillons comparatifs (Tom ×3, Pierre ×2, Jessica,
  Siwis) pour que le parent choisisse à l'oreille ;
- ambiance sonore procédurale (grillons + nappe chaude) sous la narration.
"""
import base64
import json
import subprocess
import wave
from pathlib import Path

import lameenc
import numpy as np
import sherpa_onnx

HERE = Path(__file__).parent
PIPER = HERE / "piper" / "piper"
SCENES = json.loads((HERE / "scenes.json").read_text())

LAB_TEXT = (
    "Ce soir-là, dans le jardin, la lumière était toute dorée, comme du "
    "miel. Chut… Regarde, là, sous les feuilles… un hérisson !"
)


def make_tts(model_dir: str, model: str, length: float,
             noise: float = 0.667, noise_w: float = 0.8) -> sherpa_onnx.OfflineTts:
    return sherpa_onnx.OfflineTts(
        sherpa_onnx.OfflineTtsConfig(
            model=sherpa_onnx.OfflineTtsModelConfig(
                vits=sherpa_onnx.OfflineTtsVitsModelConfig(
                    model=str(HERE / model_dir / model),
                    tokens=str(HERE / model_dir / "tokens.txt"),
                    data_dir=str(HERE / model_dir / "espeak-ng-data"),
                    length_scale=length,
                    noise_scale=noise,
                    noise_scale_w=noise_w,
                ),
                num_threads=4,
            ),
        ),
    )


def sherpa_wav(tts: sherpa_onnx.OfflineTts, text: str, out: Path,
               pitch: float = 1.0, sid: int = 0) -> None:
    audio = tts.generate(text, sid=sid)
    pcm = (np.clip(np.array(audio.samples), -1, 1) * 32767).astype(np.int16)
    with wave.open(str(out), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(int(audio.sample_rate * pitch))
        w.writeframes(pcm.tobytes())


def siwis_wav(text: str, out: Path) -> None:
    subprocess.run(
        [str(PIPER), "--model", str(HERE / "fr-siwis-medium.onnx"),
         "--length_scale", "1.22", "--sentence_silence", "0.45",
         "--output_file", str(out)],
        input=text.encode(), check=True, capture_output=True,
    )


def wav_to_mp3(path: Path, kbps: int = 48) -> bytes:
    with wave.open(str(path), "rb") as w:
        rate, channels = w.getframerate(), w.getnchannels()
        pcm = w.readframes(w.getnframes())
    enc = lameenc.Encoder()
    enc.set_bit_rate(kbps)
    enc.set_in_sample_rate(rate)
    enc.set_channels(channels)
    enc.set_quality(2)
    return bytes(enc.encode(pcm)) + bytes(enc.flush())


def data_uri(mp3: bytes) -> str:
    return "data:audio/mpeg;base64," + base64.b64encode(mp3).decode()


# ─── Ambiance : soir d'été procédural (grillons + nappe + souffle) ──────────
def build_ambience(seconds: float = 32.0, rate: int = 22050) -> bytes:
    rng = np.random.default_rng(7)
    n = int(seconds * rate)
    t = np.arange(n) / rate

    # Nappe chaude : accord doux (ré majeur grave), respiration très lente.
    pad = (
        0.30 * np.sin(2 * np.pi * 73.42 * t)   # ré2
        + 0.22 * np.sin(2 * np.pi * 110.0 * t)  # la2
        + 0.16 * np.sin(2 * np.pi * 146.83 * t)  # ré3
    )
    pad *= 0.5 + 0.5 * np.sin(2 * np.pi * t / 16.0 - np.pi / 2)  # 16 s de cycle

    # Souffle du vent : bruit filtré (moyenne glissante).
    wind = rng.standard_normal(n)
    kernel = np.ones(400) / 400
    wind = np.convolve(wind, kernel, mode="same")
    wind /= np.abs(wind).max()
    wind *= 0.5 + 0.5 * np.sin(2 * np.pi * t / 11.0)  # rafales lentes

    # Grillons : trains de stridulations à ~4,1 kHz, modulés à 32 Hz.
    crickets = np.zeros(n)
    carrier = np.sin(2 * np.pi * 4100 * t)
    mod = (np.sin(2 * np.pi * 32 * t) > 0.2).astype(float)
    pos = 1.0
    while pos < seconds - 2.5:
        dur = rng.uniform(0.5, 1.4)
        i0, i1 = int(pos * rate), int((pos + dur) * rate)
        env = np.hanning(i1 - i0)
        crickets[i0:i1] += carrier[i0:i1] * mod[i0:i1] * env * rng.uniform(0.5, 1.0)
        pos += dur + rng.uniform(0.8, 2.2)

    mix = 0.16 * pad + 0.05 * wind + 0.045 * crickets

    # Boucle sans couture : fondu croisé des 2 dernières s sur les 2 premières.
    fade = int(2.0 * rate)
    ramp = np.linspace(0, 1, fade)
    mix[:fade] = mix[:fade] * ramp + mix[-fade:] * (1 - ramp)
    mix = mix[: n - fade]

    pcm = (np.clip(mix, -1, 1) * 32767).astype(np.int16)
    tmp = HERE / "tmp-ambience.wav"
    with wave.open(str(tmp), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(pcm.tobytes())
    mp3 = wav_to_mp3(tmp, kbps=32)
    tmp.unlink()
    return mp3


def main() -> None:
    tom_dir, tom_model = "vits-piper-fr_FR-tom-medium", "fr_FR-tom-medium.onnx"
    upmc_dir, upmc_model = "vits-piper-fr_FR-upmc-medium", "fr_FR-upmc-medium.onnx"

    # Voix du récit : Tom expressif (intonations plus vivantes) + grave modéré.
    tom_expr = make_tts(tom_dir, tom_model, length=1.10, noise=0.75, noise_w=1.05)

    narration: dict[str, dict[str, str]] = {"homme": {}, "femme": {}}
    for key, text in SCENES.items():
        wav = HERE / "tmp-n.wav"
        sherpa_wav(tom_expr, text, wav, pitch=0.90)
        narration["homme"][key] = data_uri(wav_to_mp3(wav))
        siwis_wav(text, wav)
        narration["femme"][key] = data_uri(wav_to_mp3(wav))
        wav.unlink()
        print(f"narration/{key} ok")

    # Studio des voix : mêmes phrases, réglages différents.
    tom_nat = make_tts(tom_dir, tom_model, length=1.15)
    upmc = make_tts(upmc_dir, upmc_model, length=1.15, noise=0.72, noise_w=0.95)
    lab: dict[str, str] = {}
    wav = HERE / "tmp-lab.wav"
    for name, (tts, pitch, sid) in {
        "tom_expressif_grave": (tom_expr, 0.90, 0),
        "tom_naturel": (tom_nat, 1.0, 0),
        "tom_tres_grave": (tom_nat, 0.85, 0),
        "pierre_naturel": (upmc, 1.0, 1),
        "pierre_grave": (upmc, 0.91, 1),
        "jessica": (upmc, 1.0, 0),
    }.items():
        sherpa_wav(tts, LAB_TEXT, wav, pitch=pitch, sid=sid)
        lab[name] = data_uri(wav_to_mp3(wav))
        print(f"lab/{name} ok")
    siwis_wav(LAB_TEXT, wav)
    lab["siwis_femme"] = data_uri(wav_to_mp3(wav))
    wav.unlink()

    ambience = data_uri(build_ambience())
    print(f"ambience: {len(ambience) // 1024} Ko")

    js = (
        "const NARRATION = " + json.dumps(narration) + ";\n"
        + "const VOICELAB = " + json.dumps(lab) + ";\n"
        + "const AMBIENCE = " + json.dumps(ambience) + ";\n"
    )
    (HERE / "narration.js").write_text(js)
    print(f"narration.js: {len(js) // 1024} Ko")


if __name__ == "__main__":
    main()
