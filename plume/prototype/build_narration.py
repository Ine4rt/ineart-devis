#!/usr/bin/env python3
"""Narration v9 — le rythme du conteur :
- voix homme = Pierre naturel (choix du client) ;
- chaque scène est générée PHRASE PAR PHRASE, avec des silences calibrés
  entre les phrases (0,7 à 1,4 s) : la voix respire, les images ont le
  temps de se poser ;
- même traitement pour la voix femme (Siwis) ;
- studio des voix et ambiance conservés.
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
RATE = 22050

# Chaque scène = liste de (phrase, pause_après_en_secondes).
SCENES: dict[str, list[tuple[str, float]]] = {
    "test": [
        ("Bonsoir…", 0.9),
        ("Installe-toi bien sous ta couverture.", 0.8),
        ("Voilà…", 0.9),
        ("L'histoire va bientôt commencer.", 0.0),
    ],
    "open1": [
        ("Ce soir-là, dans le jardin, la lumière était toute dorée… comme du miel.", 1.1),
        ("Ton ami t'attendait près de la porte, comme tous les soirs.", 0.0),
    ],
    "open2": [
        ("Et puis, soudain…", 0.9),
        ("cric… crac !", 0.8),
        ("La haie s'est mise à bouger !", 1.0),
        ("Ton ami a dressé une oreille.", 1.0),
        ("Chut…", 1.1),
        ("Regarde, là, sous les feuilles…", 0.9),
        ("Un hérisson !", 0.9),
        ("Un tout petit hérisson, rond comme une brosse, avec deux petits yeux noirs comme des myrtilles.", 1.0),
        ("Il avait perdu son chemin.", 0.0),
    ],
    "amb1": [
        ("Le vent du soir soufflait tout doucement dans les feuilles…", 1.0),
        ("comme un long chuchotement.", 1.2),
        ("On aurait dit que le jardin murmurait une berceuse.", 1.1),
        ("Même le merle, sur sa branche, chantait sa toute dernière chanson.", 0.9),
        ("Celle qu'il garde pour le dodo.", 0.0),
    ],
    "amb2": [
        ("Ton ami marchait devant, fier comme un capitaine.", 1.0),
        ("Un pas… deux pas… trois pas…", 0.8),
        ("et hop !", 0.8),
        ("Il se retournait, pour vérifier que tu suivais bien.", 1.0),
        ("C'était sa façon de dire : viens !", 0.7),
        ("Je connais le chemin.", 0.0),
    ],
    "amb3": [
        ("Dans le ciel, la première étoile venait de s'allumer.", 1.0),
        ("Puis une deuxième…", 0.9),
        ("Puis une troisième…", 1.0),
        ("Elles clignotaient doucement, comme pour dire bonsoir.", 1.1),
        ("Ce sont les mêmes étoiles qui veillent sur toi, chaque nuit…", 0.8),
        ("sans jamais oublier.", 0.0),
    ],
    "amb4": [
        ("Dans la maison, les fenêtres faisaient des petits carrés de lumière dorée sur l'herbe.", 1.0),
        ("On entendait les petits bruits du soir…", 0.9),
        ("une casserole…", 0.7),
        ("une chaise…", 0.7),
        ("une voix douce.", 1.1),
        ("Les bruits d'une maison qui va bientôt s'endormir.", 0.0),
    ],
    "victory": [
        ("Et tu sais quoi ?", 0.9),
        ("Dans le jardin, tout le monde était déjà au courant…", 0.9),
        ("Aujourd'hui, tu as réussi quelque chose de difficile.", 1.0),
        ("Bravo !", 0.9),
        ("Le merle l'a chanté au pommier…", 0.8),
        ("Le pommier l'a murmuré aux étoiles…", 1.0),
        ("Et les étoiles, ce soir, brillaient un petit peu plus fort.", 0.9),
        ("Rien que pour toi.", 0.0),
    ],
    "fear": [
        ("Le petit hérisson, lui, était tout inquiet… roulé en boule, comme un caillou.", 1.1),
        ("Alors tu t'es approché doucement… tout doucement…", 1.0),
        ("et tu lui as dit :", 0.7),
        ("Tu sais, petit hérisson… moi aussi, parfois, j'ai un peu peur.", 1.1),
        ("Mais quand on y va quand même… après, on est drôlement fier.", 1.0),
        ("Et presque toujours, c'était moins difficile qu'on croyait.", 1.1),
        ("Alors le hérisson a sorti le tout petit bout de son nez.", 0.9),
        ("Ça voulait dire : d'accord…", 0.8),
        ("J'essaierai.", 0.0),
    ],
    "close1": [
        ("Alors, sans faire de bruit, tu as posé un petit bol d'eau près de la haie.", 1.0),
        ("Le hérisson s'est approché…", 0.0),
    ],
    "close2": [
        ("Et il a bu !", 0.9),
        ("À toutes petites gorgées.", 0.0),
    ],
    "close3": [
        ("Puis la nuit est tombée pour de bon…", 1.0),
        ("et il a disparu sous les feuilles, pour faire dodo, lui aussi.", 1.2),
        ("Est-ce qu'il reviendra demain ?", 1.1),
        ("Chut…", 1.0),
        ("C'est le secret du jardin.", 1.2),
        ("Ferme les yeux…", 1.4),
        ("Bonne nuit.", 0.0),
    ],
}

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


def sherpa_pcm(tts: sherpa_onnx.OfflineTts, text: str, sid: int = 0,
               expected_rate: int | None = RATE) -> np.ndarray:
    audio = tts.generate(text, sid=sid)
    if expected_rate is not None:
        assert audio.sample_rate == expected_rate, audio.sample_rate
    return np.clip(np.array(audio.samples), -1, 1), audio.sample_rate


def siwis_pcm(text: str) -> np.ndarray:
    tmp = HERE / "tmp-siwis.wav"
    subprocess.run(
        [str(PIPER), "--model", str(HERE / "fr-siwis-medium.onnx"),
         "--length_scale", "1.22", "--output_file", str(tmp)],
        input=text.encode(), check=True, capture_output=True,
    )
    with wave.open(str(tmp), "rb") as w:
        assert w.getframerate() == RATE
        pcm = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16)
    tmp.unlink()
    return pcm.astype(np.float64) / 32767


PAUSE_SCALE = 0.6  # retour client : pauses après les points trop longues


def scene_pcm(fragments: list[tuple[str, float]], synth) -> np.ndarray:
    parts: list[np.ndarray] = []
    for text, pause in fragments:
        parts.append(synth(text))
        if pause > 0:
            parts.append(np.zeros(int(pause * PAUSE_SCALE * RATE)))
    return np.concatenate(parts)


def pcm_to_mp3(pcm: np.ndarray, kbps: int = 48) -> bytes:
    data = (np.clip(pcm, -1, 1) * 32767).astype(np.int16)
    enc = lameenc.Encoder()
    enc.set_bit_rate(kbps)
    enc.set_in_sample_rate(RATE)
    enc.set_channels(1)
    enc.set_quality(2)
    return bytes(enc.encode(data.tobytes())) + bytes(enc.flush())


def data_uri(mp3: bytes) -> str:
    return "data:audio/mpeg;base64," + base64.b64encode(mp3).decode()


def build_ambience(seconds: float = 32.0) -> bytes:
    rng = np.random.default_rng(7)
    n = int(seconds * RATE)
    t = np.arange(n) / RATE
    pad = (0.30 * np.sin(2 * np.pi * 73.42 * t)
           + 0.22 * np.sin(2 * np.pi * 110.0 * t)
           + 0.16 * np.sin(2 * np.pi * 146.83 * t))
    pad *= 0.5 + 0.5 * np.sin(2 * np.pi * t / 16.0 - np.pi / 2)
    wind = np.convolve(rng.standard_normal(n), np.ones(400) / 400, mode="same")
    wind /= np.abs(wind).max()
    wind *= 0.5 + 0.5 * np.sin(2 * np.pi * t / 11.0)
    crickets = np.zeros(n)
    carrier = np.sin(2 * np.pi * 4100 * t)
    mod = (np.sin(2 * np.pi * 32 * t) > 0.2).astype(float)
    pos = 1.0
    while pos < seconds - 2.5:
        dur = rng.uniform(0.5, 1.4)
        i0, i1 = int(pos * RATE), int((pos + dur) * RATE)
        crickets[i0:i1] += carrier[i0:i1] * mod[i0:i1] * np.hanning(i1 - i0) * rng.uniform(0.5, 1.0)
        pos += dur + rng.uniform(0.8, 2.2)
    mix = 0.16 * pad + 0.05 * wind + 0.045 * crickets
    fade = int(2.0 * RATE)
    ramp = np.linspace(0, 1, fade)
    mix[:fade] = mix[:fade] * ramp + mix[-fade:] * (1 - ramp)
    return pcm_to_mp3(mix[: n - fade], kbps=32)


def main() -> None:
    upmc_dir, upmc_model = "vits-piper-fr_FR-upmc-medium", "fr_FR-upmc-medium.onnx"
    tom_dir, tom_model = "vits-piper-fr_FR-tom-medium", "fr_FR-tom-medium.onnx"

    # La voix du conteur : Pierre naturel (réglage validé au studio des voix).
    pierre = make_tts(upmc_dir, upmc_model, length=1.15, noise=0.72, noise_w=0.95)

    narration: dict[str, dict[str, str]] = {"homme": {}, "femme": {}}
    plain: dict[str, str] = {}
    for key, fragments in SCENES.items():
        plain[key] = " ".join(text for text, _ in fragments)
        narration["homme"][key] = data_uri(pcm_to_mp3(
            scene_pcm(fragments, lambda text: sherpa_pcm(pierre, text, sid=1)[0])))
        narration["femme"][key] = data_uri(pcm_to_mp3(
            scene_pcm(fragments, siwis_pcm)))
        print(f"narration/{key} ok")

    # Studio des voix (échantillons inchangés, Pierre naturel = voix actuelle).
    tom_expr = make_tts(tom_dir, tom_model, length=1.10, noise=0.75, noise_w=1.05)
    tom_nat = make_tts(tom_dir, tom_model, length=1.15)
    lab: dict[str, str] = {}
    def lab_wav(tts, pitch, sid):
        pcm, rate = sherpa_pcm(tts, LAB_TEXT, sid=sid, expected_rate=None)
        data = (np.clip(pcm, -1, 1) * 32767).astype(np.int16)
        enc = lameenc.Encoder()
        enc.set_bit_rate(48); enc.set_in_sample_rate(int(rate * pitch))
        enc.set_channels(1); enc.set_quality(2)
        return bytes(enc.encode(data.tobytes())) + bytes(enc.flush())
    for name, (tts, pitch, sid) in {
        "pierre_naturel": (pierre, 1.0, 1),
        "pierre_grave": (pierre, 0.91, 1),
        "tom_expressif_grave": (tom_expr, 0.90, 0),
        "tom_naturel": (tom_nat, 1.0, 0),
        "jessica": (pierre, 1.0, 0),
    }.items():
        lab[name] = data_uri(lab_wav(tts, pitch, sid))
        print(f"lab/{name} ok")
    lab["siwis_femme"] = data_uri(pcm_to_mp3(siwis_pcm(LAB_TEXT)))

    js = ("const NARRATION = " + json.dumps(narration) + ";\n"
          + "const VOICELAB = " + json.dumps(lab) + ";\n"
          + "const AMBIENCE = " + json.dumps(data_uri(build_ambience())) + ";\n")
    (HERE / "narration.js").write_text(js)
    (HERE / "scenes.json").write_text(json.dumps(plain, ensure_ascii=False))
    print(f"narration.js: {len(js) // 1024} Ko")


if __name__ == "__main__":
    main()
