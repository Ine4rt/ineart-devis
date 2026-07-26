#!/usr/bin/env python3
"""Bande-son de Plume, 100 % synthétisée (gratuite, aucune banque de sons) :

- MUSIQUE : berceuse de boîte à musique (ré majeur pentatonique) sur nappe
  chaude, souffle léger et grillons discrets — boucle sans couture ;
- BRUITAGES par scène, en rapport avec le texte :
    open    → feuilles qui bruissent + « cric, crac » de brindille
    amb1    → rafale douce dans les arbres + chant du merle
    amb2    → petits pas dans l'herbe
    amb3    → carillon d'étoiles (clochettes rares)
    amb4    → bruits de maison feutrés (tintement, chaise)
    victory → merle joyeux + carillon ascendant
    fear    → froissement inquiet puis clochette rassurante
    close   → petites gorgées d'eau + dernier carillon + souffle qui s'éteint

Sortie : soundscape.js (const AMBIENCE, const SFX) à injecter dans le proto.
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


def env_hann(n: int) -> np.ndarray:
    return np.hanning(n)


def place(target: np.ndarray, clip: np.ndarray, at: float, gain: float = 1.0) -> None:
    i0 = int(at * RATE)
    i1 = min(i0 + len(clip), len(target))
    target[i0:i1] += clip[: i1 - i0] * gain


def lowpass(x: np.ndarray, width: int) -> np.ndarray:
    return np.convolve(x, np.ones(width) / width, mode="same")


# ─── Instruments ────────────────────────────────────────────────────────────

def music_box_note(freq: float, dur: float = 2.2, gain: float = 1.0) -> np.ndarray:
    """Note de boîte à musique : fondamentale + partiels cloche, chute exp."""
    t = t_axis(dur)
    tone = (np.sin(2 * np.pi * freq * t)
            + 0.35 * np.sin(2 * np.pi * freq * 2.0 * t)
            + 0.18 * np.sin(2 * np.pi * freq * 4.16 * t))
    envelope = np.exp(-t * 2.6)
    attack = int(0.004 * RATE)
    envelope[:attack] *= np.linspace(0, 1, attack)
    return tone * envelope * gain


def bell(freq: float, dur: float = 2.0, gain: float = 1.0) -> np.ndarray:
    t = t_axis(dur)
    tone = (np.sin(2 * np.pi * freq * t)
            + 0.4 * np.sin(2 * np.pi * freq * 2.76 * t)
            + 0.2 * np.sin(2 * np.pi * freq * 5.4 * t))
    return tone * np.exp(-t * 3.0) * gain


def chirp(f0: float, f1: float, dur: float, gain: float = 1.0) -> np.ndarray:
    """Syllabe d'oiseau : glissando avec vibrato."""
    t = t_axis(dur)
    freq = f0 + (f1 - f0) * t / dur
    vib = 1 + 0.02 * np.sin(2 * np.pi * 40 * t)
    phase = 2 * np.pi * np.cumsum(freq * vib) / RATE
    return np.sin(phase) * env_hann(len(t)) * gain


def blackbird(rng: np.random.Generator, gain: float = 1.0) -> np.ndarray:
    """Petit motif de merle : 3-5 syllabes flûtées."""
    parts = []
    for _ in range(rng.integers(3, 6)):
        f0 = rng.uniform(1700, 2600)
        f1 = f0 * rng.uniform(0.7, 1.35)
        parts.append(chirp(f0, f1, rng.uniform(0.08, 0.16), gain))
        parts.append(silence(rng.uniform(0.05, 0.14)))
    return np.concatenate(parts)


def rustle(dur: float, swells: int, rng: np.random.Generator, gain: float = 1.0) -> np.ndarray:
    """Feuillage : bruit filtré en bande, gonflements successifs."""
    n = int(dur * RATE)
    noise = rng.standard_normal(n)
    band = lowpass(noise, 4) - lowpass(noise, 24)  # ~1-5 kHz
    envelope = np.zeros(n)
    for k in range(swells):
        at = (k + rng.uniform(0.1, 0.5)) * dur / swells
        width = int(rng.uniform(0.5, 0.9) * RATE)
        i0 = int(at * RATE)
        i1 = min(i0 + width, n)
        envelope[i0:i1] += env_hann(i1 - i0)
    return band * envelope * gain


def twig_snap(rng: np.random.Generator, gain: float = 1.0) -> np.ndarray:
    """« cric, crac » : deux claquements secs + petit choc sourd."""
    def snap() -> np.ndarray:
        click = rng.standard_normal(int(0.02 * RATE)) * np.exp(-np.linspace(0, 8, int(0.02 * RATE)))
        knock = np.sin(2 * np.pi * 180 * t_axis(0.07)) * np.exp(-t_axis(0.07) * 40)
        return np.concatenate([click, knock])
    return np.concatenate([snap(), silence(0.22), snap()]) * gain


def footsteps(count: int, rng: np.random.Generator, gain: float = 1.0) -> np.ndarray:
    """Pas feutrés dans l'herbe."""
    total = silence(0.62 * count + 0.3)
    for k in range(count):
        thump = np.sin(2 * np.pi * 95 * t_axis(0.09)) * np.exp(-t_axis(0.09) * 45)
        grass = lowpass(rng.standard_normal(int(0.08 * RATE)), 6) * env_hann(int(0.08 * RATE)) * 0.8
        step = thump + grass[: len(thump)] if len(grass) >= len(thump) else thump
        place(total, step, 0.15 + k * 0.62 + rng.uniform(-0.05, 0.05))
    return total * gain


def droplets(count: int, rng: np.random.Generator, gain: float = 1.0) -> np.ndarray:
    """Petites gorgées / gouttes : pings descendants très courts."""
    total = silence(0.2 * count + 0.4)
    for k in range(count):
        d = chirp(rng.uniform(1100, 1400), rng.uniform(500, 700), 0.045)
        place(total, d, 0.1 + k * rng.uniform(0.16, 0.24))
    return total * gain


# ─── Musique de fond (boucle) ───────────────────────────────────────────────

def build_music(seconds: float = 44.0) -> np.ndarray:
    rng = np.random.default_rng(11)
    n = int(seconds * RATE)
    t = t_axis(seconds)

    # Nappe chaude très discrète (ré, la, ré à l'octave).
    pad = (0.24 * np.sin(2 * np.pi * 73.42 * t)
           + 0.18 * np.sin(2 * np.pi * 110.0 * t)
           + 0.12 * np.sin(2 * np.pi * 146.83 * t))
    pad *= 0.55 + 0.45 * np.sin(2 * np.pi * t / 22.0 - np.pi / 2)

    # Berceuse : pentatonique de ré (ré, mi, fa#, la, si), tempo lent.
    d5, e5, fs5, a5, b5, d6 = 587.33, 659.25, 739.99, 880.0, 987.77, 1174.66
    melody = [d5, fs5, a5, b5, a5, fs5, e5, d5,
              fs5, a5, d6, b5, a5, fs5, e5, d5]
    music = np.zeros(n)
    beat = 2.6  # une note toutes les ~2,6 s : très paisible
    for k, note in enumerate(melody):
        at = 0.8 + k * beat
        if at + 2.4 >= seconds:
            break
        gain = 0.30 if k % 4 == 0 else 0.22
        place(music, music_box_note(note, gain=gain), at + rng.uniform(-0.08, 0.08))

    # Souffle léger + grillons rares (plus discrets qu'avant).
    wind = lowpass(rng.standard_normal(n), 400)
    wind /= np.abs(wind).max()
    wind *= 0.5 + 0.5 * np.sin(2 * np.pi * t / 13.0)
    crickets = np.zeros(n)
    carrier = np.sin(2 * np.pi * 4100 * t)
    mod = (np.sin(2 * np.pi * 32 * t) > 0.2).astype(float)
    pos = 2.0
    while pos < seconds - 3.0:
        dur = rng.uniform(0.4, 0.9)
        i0, i1 = int(pos * RATE), int((pos + dur) * RATE)
        crickets[i0:i1] += carrier[i0:i1] * mod[i0:i1] * env_hann(i1 - i0) * rng.uniform(0.3, 0.6)
        pos += dur + rng.uniform(2.0, 4.5)

    mix = 0.14 * pad + 0.30 * music + 0.035 * wind + 0.030 * crickets

    # Boucle sans couture.
    fade = int(2.5 * RATE)
    ramp = np.linspace(0, 1, fade)
    mix[:fade] = mix[:fade] * ramp + mix[-fade:] * (1 - ramp)
    return mix[: n - fade]


# ─── Bruitages par scène ────────────────────────────────────────────────────

def build_sfx() -> dict[str, np.ndarray]:
    rng = np.random.default_rng(23)
    sfx: dict[str, np.ndarray] = {}

    # open : feuilles + cric-crac + petit froissement (le hérisson).
    s = silence(7.0)
    place(s, rustle(2.2, 2, rng), 0.2, 0.8)
    place(s, twig_snap(rng), 2.8, 0.9)
    place(s, rustle(1.6, 3, rng), 4.4, 0.55)
    sfx["open"] = s

    # amb1 : rafale douce + merle.
    s = silence(7.0)
    place(s, rustle(3.2, 3, rng), 0.2, 0.7)
    place(s, blackbird(rng), 3.8, 0.5)
    sfx["amb1"] = s

    # amb2 : pas dans l'herbe.
    sfx["amb2"] = footsteps(5, rng, 0.9)

    # amb3 : carillon d'étoiles.
    s = silence(6.5)
    for k, note in enumerate([1174.66, 1760.0, 1479.98]):
        place(s, bell(note, gain=0.5), 0.4 + k * 1.6)
    sfx["amb3"] = s

    # amb4 : maison feutrée — tintement + chaise.
    s = silence(6.0)
    place(s, bell(820, dur=1.0, gain=0.35), 0.5)
    creak_t = t_axis(0.35)
    creak = np.sin(2 * np.pi * (120 - 40 * creak_t / 0.35) * creak_t) * env_hann(len(creak_t))
    place(s, creak, 2.4, 0.4)
    place(s, bell(650, dur=0.9, gain=0.3), 4.2)
    sfx["amb4"] = s

    # victory : merle joyeux + carillon ascendant.
    s = silence(6.0)
    place(s, blackbird(rng), 0.3, 0.6)
    for k, note in enumerate([587.33, 739.99, 880.0, 1174.66]):
        place(s, bell(note, gain=0.45), 2.2 + k * 0.5)
    sfx["victory"] = s

    # fear : froissement inquiet, puis clochette qui rassure.
    s = silence(7.0)
    place(s, rustle(1.8, 4, rng), 0.3, 0.6)
    place(s, bell(587.33, gain=0.4), 3.6)
    place(s, bell(880.0, gain=0.4), 4.6)
    sfx["fear"] = s

    # close : gorgées + dernier carillon + souffle qui s'éteint.
    s = silence(9.0)
    place(s, droplets(6, rng), 0.4, 0.8)
    place(s, rustle(1.4, 2, rng), 2.6, 0.4)
    place(s, bell(1174.66, dur=3.0, gain=0.4), 5.0)
    dying = lowpass(rng.standard_normal(int(3.0 * RATE)), 400)
    dying /= np.abs(dying).max()
    place(s, dying * np.linspace(0.5, 0.0, len(dying)), 5.5, 0.12)
    sfx["close"] = s

    return sfx


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


def main() -> None:
    music = data_uri(pcm_to_mp3(build_music(), kbps=48))
    sfx = {key: data_uri(pcm_to_mp3(pcm)) for key, pcm in build_sfx().items()}
    js = ("const AMBIENCE = " + json.dumps(music) + ";\n"
          + "const SFX = " + json.dumps(sfx) + ";\n")
    (HERE / "soundscape.js").write_text(js)
    total = len(js) // 1024
    print(f"musique: {len(music) // 1024} Ko · sfx: {len(sfx)} scènes · soundscape.js: {total} Ko")


if __name__ == "__main__":
    main()
