#!/usr/bin/env python3
"""Bande-son v2 — retours client : musique audible, niveaux doux, sons réalistes.

Techniques :
- craquement de branche : pré-craquements + rafale de micro-fractures
  (60-100 impulsions en 80 ms) résonnant dans un « corps de bois »
  (combs 130/310 Hz) + choc sourd final ;
- bruissement : bruit « velours » (impulsions éparses lissées), bien plus
  naturel qu'un bruit blanc filtré ;
- merle : motifs à deux notes avec portamento + petite queue de réverbe ;
- gorgées : glouglous graves à formant + gouttelettes ;
- tous les effets normalisés au même niveau crête (0,5) puis mixés bas.

Sortie : soundscape.js (AMBIENCE = musique, SFX = bruitages par scène).
"""
import base64
import json
import wave
from pathlib import Path

import lameenc
import numpy as np

HERE = Path(__file__).parent
RATE = 22050


def t_axis(seconds: float) -> np.ndarray:
    return np.arange(int(seconds * RATE)) / RATE


def silence(seconds: float) -> np.ndarray:
    return np.zeros(int(seconds * RATE))


def place(target: np.ndarray, clip: np.ndarray, at: float, gain: float = 1.0) -> None:
    i0 = int(at * RATE)
    i1 = min(i0 + len(clip), len(target))
    if i1 > i0:
        target[i0:i1] += clip[: i1 - i0] * gain


def lowpass(x: np.ndarray, width: int) -> np.ndarray:
    return np.convolve(x, np.ones(width) / width, mode="same")


def normalize(x: np.ndarray, peak: float = 0.5) -> np.ndarray:
    m = np.abs(x).max()
    return x * (peak / m) if m > 0 else x


def reverb(x: np.ndarray, delay_s: float = 0.09, decay: float = 0.35, taps: int = 4) -> np.ndarray:
    out = np.copy(x)
    d = int(delay_s * RATE)
    for k in range(1, taps + 1):
        shifted = np.zeros_like(x)
        shifted[k * d:] = x[: len(x) - k * d] * (decay ** k)
        out += shifted
    return out


def comb_body(x: np.ndarray, freqs: tuple[float, ...], decay: float = 0.5) -> np.ndarray:
    """Résonance de « corps » : combs bouclés aux fréquences données."""
    out = np.copy(x)
    for f in freqs:
        d = max(1, int(RATE / f))
        buf = np.copy(x)
        for _ in range(6):
            shifted = np.zeros_like(buf)
            shifted[d:] = buf[:-d] * decay
            buf = shifted
            out += buf
    return out


# ─── Effets réalistes ───────────────────────────────────────────────────────

def branch_crack(rng: np.random.Generator) -> np.ndarray:
    """Vrai craquement : ça craque d'abord, puis ça CASSE, puis ça retombe."""
    total = silence(2.2)

    # 1. Pré-craquements : le bois travaille (4 petits clics espacés).
    for k in range(4):
        n = int(rng.uniform(0.006, 0.012) * RATE)
        click = rng.standard_normal(n) * np.exp(-np.linspace(0, 10, n))
        place(total, comb_body(click, (400.0,), 0.4), 0.15 + k * rng.uniform(0.09, 0.16),
              rng.uniform(0.15, 0.3))

    # 2. La fracture : rafale dense de micro-impulsions sur ~90 ms.
    n = int(0.09 * RATE)
    burst = np.zeros(n)
    for _ in range(80):
        i = int(rng.uniform(0, n - 4))
        burst[i:i + 3] += rng.uniform(-1, 1)
    burst *= np.exp(-np.linspace(0, 5, n))
    fracture = comb_body(burst, (130.0, 310.0), 0.55)
    place(total, fracture, 0.82, 1.0)

    # 3. Choc sourd (la branche cède) + petites retombées.
    thud = np.sin(2 * np.pi * 85 * t_axis(0.16)) * np.exp(-t_axis(0.16) * 26)
    place(total, thud, 0.90, 0.8)
    for k in range(3):
        n2 = int(0.01 * RATE)
        debris = rng.standard_normal(n2) * np.exp(-np.linspace(0, 9, n2))
        place(total, debris, 1.1 + k * rng.uniform(0.1, 0.2), rng.uniform(0.08, 0.15))

    return normalize(reverb(total, 0.07, 0.25, 2))


def velvet_rustle(dur: float, density: float, rng: np.random.Generator,
                  swells: int = 2) -> np.ndarray:
    """Feuillage en bruit « velours » : impulsions éparses lissées."""
    n = int(dur * RATE)
    x = np.zeros(n)
    count = int(density * dur)
    for _ in range(count):
        i = int(rng.uniform(0, n - 2))
        x[i] = rng.choice((-1.0, 1.0)) * rng.uniform(0.4, 1.0)
    x = lowpass(x, 3) - lowpass(x, 14)  # timbre feuille sèche
    envelope = np.zeros(n)
    for k in range(swells):
        at = (k + rng.uniform(0.15, 0.45)) * dur / swells
        width = int(rng.uniform(0.6, 1.1) * RATE)
        i0 = int(at * RATE)
        i1 = min(i0 + width, n)
        if i1 > i0:
            envelope[i0:i1] += np.hanning(i1 - i0)
    return normalize(x * envelope)


def blackbird(rng: np.random.Generator) -> np.ndarray:
    """Merle : motifs flûtés à deux notes, portamento, un peu d'air."""
    parts = []
    for _ in range(rng.integers(2, 4)):
        f0 = rng.uniform(1800, 2400)
        f1 = f0 * rng.uniform(1.15, 1.4)
        for fa, fb, dur in ((f0, f1, 0.13), (f1, f0 * rng.uniform(0.85, 1.0), 0.17)):
            t = t_axis(dur)
            freq = fa + (fb - fa) * (t / dur) ** 0.7
            vib = 1 + 0.015 * np.sin(2 * np.pi * 35 * t)
            phase = 2 * np.pi * np.cumsum(freq * vib) / RATE
            syl = np.sin(phase) * np.hanning(len(t))
            parts.append(syl)
            parts.append(silence(rng.uniform(0.04, 0.1)))
        parts.append(silence(rng.uniform(0.25, 0.5)))
    song = np.concatenate(parts)
    return normalize(reverb(song, 0.11, 0.3, 3))


def grass_steps(count: int, rng: np.random.Generator) -> np.ndarray:
    total = silence(0.72 * count + 0.4)
    for k in range(count):
        swish = velvet_rustle(0.16, 900, rng, swells=1)
        thump = np.sin(2 * np.pi * 78 * t_axis(0.07)) * np.exp(-t_axis(0.07) * 55)
        at = 0.2 + k * 0.72 + rng.uniform(-0.04, 0.04)
        place(total, swish, at, 0.7)
        place(total, thump, at + 0.02, 0.5)
    return normalize(total)


def star_chimes() -> np.ndarray:
    total = silence(6.5)
    for k, note in enumerate((1174.66, 1760.0, 1479.98)):
        t = t_axis(2.6)
        tone = (np.sin(2 * np.pi * note / 2 * t)
                + 0.12 * np.sin(2 * np.pi * note * 1.38 * t)) * np.exp(-t * 2.0)
        place(total, tone, 0.5 + k * 1.7, 0.7)
    return normalize(reverb(total, 0.13, 0.35, 3))


def house_sounds(rng: np.random.Generator) -> np.ndarray:
    total = silence(6.0)
    # tintement feutré (vaisselle au loin) : cloche très amortie.
    t = t_axis(0.7)
    clink = np.sin(2 * np.pi * 840 * t) * np.exp(-t * 9)
    place(total, lowpass(clink, 6), 0.6, 0.7)
    # chaise : frottement grave descendant.
    t = t_axis(0.4)
    creak = np.sin(2 * np.pi * (115 - 35 * t / 0.4) * t) * np.hanning(len(t))
    place(total, creak, 2.6, 0.5)
    place(total, lowpass(clink, 6), 4.3, 0.5)
    return normalize(total)


def water_sips(rng: np.random.Generator) -> np.ndarray:
    """Petites gorgées : glouglous graves + gouttelettes."""
    total = silence(5.0)
    for k in range(5):
        # glouglou : sinus descendant avec formant (bouche/gorge).
        dur = rng.uniform(0.07, 0.1)
        t = t_axis(dur)
        f = rng.uniform(300, 380) * (1 - 0.5 * t / dur)
        glug = np.sin(2 * np.pi * np.cumsum(f) / RATE) * np.hanning(len(t))
        place(total, glug, 0.3 + k * rng.uniform(0.28, 0.4), 0.8)
        # gouttelette juste après.
        t2 = t_axis(0.03)
        fr = rng.uniform(1000, 1400)
        drop = np.sin(2 * np.pi * fr * (1 - 0.4 * t2 / 0.03) * t2) * np.hanning(len(t2))
        place(total, drop, 0.36 + k * 0.33, 0.25)
    return normalize(total)


def night_settle(rng: np.random.Generator) -> np.ndarray:
    """Fin : dernier carillon + souffle qui s'éteint."""
    total = silence(8.0)
    t = t_axis(4.0)
    bell = (np.sin(2 * np.pi * 1174.66 * t)
            + 0.3 * np.sin(2 * np.pi * 1174.66 * 2.76 * t)) * np.exp(-t * 1.6)
    place(total, bell, 0.5, 0.7)
    wind = lowpass(rng.standard_normal(int(5.0 * RATE)), 500)
    wind = normalize(wind, 1.0) * np.linspace(0.6, 0.0, int(5.0 * RATE))
    place(total, wind, 2.5, 0.3)
    return normalize(reverb(total, 0.12, 0.3, 2))


# ─── Musique : berceuse audible mais douce ──────────────────────────────────

def music_box_note(freq: float, dur: float = 2.8) -> np.ndarray:
    # Timbre chaud : presque pas d'harmoniques aigus (« ding » gommés),
    # attaque adoucie, chute plus lente.
    t = t_axis(dur)
    tone = (np.sin(2 * np.pi * freq * t)
            + 0.15 * np.sin(2 * np.pi * freq * 2.0 * t)
            + 0.03 * np.sin(2 * np.pi * freq * 4.16 * t))
    envelope = np.exp(-t * 1.7)
    attack = int(0.02 * RATE)
    envelope[:attack] *= np.linspace(0, 1, attack)
    return tone * envelope


def build_music(seconds: float = 46.0) -> np.ndarray:
    rng = np.random.default_rng(11)
    n = int(seconds * RATE)
    t = t_axis(seconds)

    # Mélodie au premier plan (c'était le retour client : on ne l'entendait pas).
    d4, e4, fs4, a4, b4, d5 = 293.66, 329.63, 369.99, 440.0, 493.88, 587.33
    melody = [d4, fs4, a4, b4, a4, fs4, e4, d4,
              fs4, a4, d5, b4, a4, fs4, e4, d4]
    music = np.zeros(n)
    beat = 2.4
    for k, note in enumerate(melody):
        at = 0.6 + k * beat
        if at + 2.4 >= seconds:
            break
        place(music, music_box_note(note), at + rng.uniform(-0.06, 0.06),
              1.0 if k % 4 == 0 else 0.75)
    music = reverb(music, 0.14, 0.3, 3)

    # Nappe discrète dessous.
    pad = (0.5 * np.sin(2 * np.pi * 146.83 * t)
           + 0.35 * np.sin(2 * np.pi * 220.0 * t))
    pad *= 0.55 + 0.45 * np.sin(2 * np.pi * t / 20.0 - np.pi / 2)

    mix = normalize(music, 0.42) + normalize(pad, 0.14)

    # Boucle sans couture.
    fade = int(2.5 * RATE)
    ramp = np.linspace(0, 1, fade)
    mix[:fade] = mix[:fade] * ramp + mix[-fade:] * (1 - ramp)
    return normalize(mix[: n - fade], 0.7)


# ─── Assemblage par scène (nouvelles clés synchronisées) ────────────────────

def build_sfx() -> dict[str, np.ndarray]:
    rng = np.random.default_rng(23)
    return {
        # open1 : le jardin calme — léger feuillage lointain.
        "open1": velvet_rustle(3.2, 1400, rng, swells=2) * 0.5,
        # open2 : « Et puis, soudain… cric, crac ! » — LE craquement, à t=0.
        "open2": branch_crack(rng),
        "amb1": np.concatenate([velvet_rustle(2.6, 2200, rng, swells=2),
                                silence(0.4), blackbird(rng) * 0.8]),
        "amb2": grass_steps(5, rng),
        "amb3": star_chimes(),
        "amb4": house_sounds(rng),
        "victory": np.concatenate([blackbird(rng) * 0.8, silence(0.3),
                                   star_chimes()[: int(3.5 * RATE)] * 0.8]),
        "fear": np.concatenate([velvet_rustle(1.6, 1800, rng, swells=3) * 0.6,
                                silence(0.6), star_chimes()[: int(2.5 * RATE)] * 0.6]),
        # close2 : « et il a bu ! » — gorgées à t=0.
        "close2": water_sips(rng),
        "close3": night_settle(rng),
    }


# ─── Encodage ───────────────────────────────────────────────────────────────

def pcm_to_mp3(pcm: np.ndarray, kbps: int = 40) -> bytes:
    data = (np.clip(pcm, -1, 1) * 32767).astype(np.int16)
    enc = lameenc.Encoder()
    enc.set_bit_rate(kbps)
    enc.set_in_sample_rate(RATE)
    enc.set_channels(1)
    enc.set_quality(2)
    return bytes(enc.encode(data.tobytes())) + bytes(enc.flush())


def data_uri(mp3: bytes) -> str:
    return "data:audio/mpeg;base64," + base64.b64encode(mp3).decode()


# Crêtes finales, gravées dans les mp3 (Safari iOS ignore element.volume).
MUSIC_PEAK = 0.05   # ≈ -25 dB sous la narration : un souffle derrière
SFX_PEAK = 0.10     # ≈ -19 dB : présent mais jamais devant la voix


def main() -> None:
    music = data_uri(pcm_to_mp3(normalize(build_music(), MUSIC_PEAK), kbps=48))
    sfx = {key: data_uri(pcm_to_mp3(normalize(pcm, SFX_PEAK)))
           for key, pcm in build_sfx().items()}
    js = ("const AMBIENCE = " + json.dumps(music) + ";\n"
          + "const SFX = " + json.dumps(sfx) + ";\n")
    (HERE / "soundscape.js").write_text(js)
    print(f"musique: {len(music) // 1024} Ko · sfx: {len(sfx)} · total: {len(js) // 1024} Ko")


if __name__ == "__main__":
    main()
