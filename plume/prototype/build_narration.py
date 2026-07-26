#!/usr/bin/env python3
"""Narration v7 : histoire réécrite en ton album jeunesse + voix homme
plus grave (Tom, abaissé d'environ 2 tons par ré-échantillonnage).

- homme : fr_FR-tom (sherpa-onnx), pitch ×0.87 → plus grave et chaleureux
- femme : fr_FR-siwis (piper), inchangée
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

# Le ton : phrases courtes, onomatopées, répétitions, douceur, adresse directe.
SCENES = {
    "test": (
        "Bonsoir… Installe-toi bien sous ta couverture. Voilà. "
        "L'histoire va bientôt commencer."
    ),
    "open": (
        "Ce soir-là, dans le jardin, la lumière était toute dorée, comme du "
        "miel. Ton ami t'attendait près de la porte, comme tous les soirs. "
        "Et puis soudain… cric, crac ! La haie s'est mise à bouger ! Ton ami "
        "a dressé une oreille. Chut… Regarde, là, sous les feuilles… un "
        "hérisson ! Un tout petit hérisson, rond comme une brosse, avec deux "
        "petits yeux noirs comme des myrtilles. Il avait perdu son chemin."
    ),
    "amb1": (
        "Le vent du soir soufflait tout doucement dans les feuilles, comme un "
        "long chuchotement. On aurait dit que le jardin murmurait une "
        "berceuse. Même le merle, sur sa branche, chantait sa toute dernière "
        "chanson. Celle qu'il garde pour le dodo."
    ),
    "amb2": (
        "Ton ami marchait devant, fier comme un capitaine. Un pas, deux pas, "
        "trois pas… et hop ! Il se retournait pour vérifier que tu suivais "
        "bien. C'était sa façon de dire : viens ! Je connais le chemin."
    ),
    "amb3": (
        "Dans le ciel, la première étoile venait de s'allumer. Puis une "
        "deuxième. Puis une troisième. Elles clignotaient doucement, comme "
        "pour dire bonsoir. Ce sont les mêmes étoiles qui veillent sur toi, "
        "chaque nuit, sans jamais oublier."
    ),
    "amb4": (
        "Dans la maison, les fenêtres faisaient des petits carrés de lumière "
        "dorée sur l'herbe. On entendait les petits bruits du soir : une "
        "casserole, une chaise, une voix douce. Les bruits d'une maison qui "
        "va bientôt s'endormir."
    ),
    "victory": (
        "Et tu sais quoi ? Dans le jardin, tout le monde était déjà au "
        "courant : aujourd'hui, tu as réussi quelque chose de difficile. "
        "Bravo ! Le merle l'a chanté au pommier. Le pommier l'a murmuré aux "
        "étoiles. Et les étoiles, ce soir, brillaient un petit peu plus "
        "fort. Rien que pour toi."
    ),
    "fear": (
        "Le petit hérisson, lui, était tout inquiet, roulé en boule comme un "
        "caillou. Alors tu t'es approché doucement, tout doucement, et tu "
        "lui as dit : tu sais, petit hérisson, moi aussi, parfois, j'ai un "
        "peu peur. Mais quand on y va quand même, après, on est drôlement "
        "fier. Et presque toujours, c'était moins difficile qu'on croyait. "
        "Alors le hérisson a sorti le tout petit bout de son nez. Ça voulait "
        "dire : d'accord. J'essaierai."
    ),
    "close": (
        "Alors, sans faire de bruit, tu as posé un petit bol d'eau près de "
        "la haie. Le hérisson s'est approché… et il a bu ! À toutes petites "
        "gorgées. Puis la nuit est tombée pour de bon, et il a disparu sous "
        "les feuilles, pour faire dodo, lui aussi. Est-ce qu'il reviendra "
        "demain ? Chut… C'est le secret du jardin. Ferme les yeux… Bonne "
        "nuit."
    ),
}

PITCH_FACTOR = 0.87  # ≈ -2,4 demi-tons : plus grave, encore naturel

tom = sherpa_onnx.OfflineTts(
    sherpa_onnx.OfflineTtsConfig(
        model=sherpa_onnx.OfflineTtsModelConfig(
            vits=sherpa_onnx.OfflineTtsVitsModelConfig(
                model=str(HERE / "vits-piper-fr_FR-tom-medium/fr_FR-tom-medium.onnx"),
                tokens=str(HERE / "vits-piper-fr_FR-tom-medium/tokens.txt"),
                data_dir=str(HERE / "vits-piper-fr_FR-tom-medium/espeak-ng-data"),
                # Généré presque à vitesse normale : le pitch-down ralentit déjà.
                length_scale=1.04,
            ),
            num_threads=4,
        ),
    ),
)


def tom_wav(text: str, out: Path) -> None:
    audio = tom.generate(text)
    pcm = (np.clip(np.array(audio.samples), -1, 1) * 32767).astype(np.int16)
    with wave.open(str(out), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        # Ré-échantillonnage déclaré plus bas = lecture plus lente et plus
        # grave (chaleur de « bande ralentie »), tempo net ≈ 1,2.
        w.setframerate(int(audio.sample_rate * PITCH_FACTOR))
        w.writeframes(pcm.tobytes())


def siwis_wav(text: str, out: Path) -> None:
    subprocess.run(
        [str(PIPER), "--model", str(HERE / "fr-siwis-medium.onnx"),
         "--length_scale", "1.22", "--sentence_silence", "0.45",
         "--output_file", str(out)],
        input=text.encode(), check=True, capture_output=True,
    )


def wav_to_mp3(path: Path) -> bytes:
    with wave.open(str(path), "rb") as w:
        rate, channels = w.getframerate(), w.getnchannels()
        pcm = w.readframes(w.getnframes())
    enc = lameenc.Encoder()
    enc.set_bit_rate(48)
    enc.set_in_sample_rate(rate)
    enc.set_channels(channels)
    enc.set_quality(2)
    return bytes(enc.encode(pcm)) + bytes(enc.flush())


def main() -> None:
    out: dict[str, dict[str, str]] = {}
    total = 0
    for vname, synth in (("homme", tom_wav), ("femme", siwis_wav)):
        out[vname] = {}
        for key, text in SCENES.items():
            wav = HERE / f"tmp-{vname}-{key}.wav"
            synth(text, wav)
            mp3 = wav_to_mp3(wav)
            wav.unlink()
            out[vname][key] = "data:audio/mpeg;base64," + base64.b64encode(mp3).decode()
            total += len(mp3)
            print(f"{vname}/{key}: {len(mp3) // 1024} Ko")
    (HERE / "narration.js").write_text(
        "const NARRATION = " + json.dumps(out) + ";\n",
    )
    (HERE / "scenes.json").write_text(json.dumps(SCENES, ensure_ascii=False))
    print(f"total: {total // 1024} Ko")


if __name__ == "__main__":
    main()
