"""매뉴얼 영상 음성(한국어 TTS) 만들기 — 인터넷 없이 동작합니다.

준비 (한 번만):
  python3 -m venv .venv && .venv/bin/pip install "sherpa-onnx==1.10.46" soundfile numpy
  curl -L -o ko.tar.bz2 https://github.com/k2-fsa/sherpa-onnx/releases/download/tts-models/vits-mimic3-ko_KO-kss_low.tar.bz2
  tar xjf ko.tar.bz2      # vits-mimic3-ko_KO-kss_low 폴더가 생깁니다

사용:
  .venv/bin/python tts.py <모델폴더> <캐시폴더> <문장목록.json>
  문장목록.json: ["문장1", "문장2", ...]
  캐시폴더에 <md5>.wav 를 만들고, <캐시폴더>/durations.json 에 길이(초)를 적습니다.
"""
import hashlib
import json
import os
import sys

import sherpa_onnx
import soundfile as sf

model_dir, cache_dir, list_path = sys.argv[1:4]
speed = float(os.environ.get('TTS_SPEED', '0.95'))  # 조금 느리고(0.95) 잡음을 줄여야(noise_scale 0.33) 또렷합니다
os.makedirs(cache_dir, exist_ok=True)
dur_path = os.path.join(cache_dir, 'durations.json')
durations = json.load(open(dur_path)) if os.path.exists(dur_path) else {}
texts = json.load(open(list_path, encoding='utf-8'))

todo = [t for t in texts if hashlib.md5(t.encode()).hexdigest() not in durations]
if todo:
    name = [f for f in os.listdir(model_dir) if f.endswith('.onnx')][0]
    cfg = sherpa_onnx.OfflineTtsConfig(model=sherpa_onnx.OfflineTtsModelConfig(
        vits=sherpa_onnx.OfflineTtsVitsModelConfig(model=os.path.join(model_dir, name), lexicon='',
                                                   tokens=os.path.join(model_dir, 'tokens.txt'),
                                                   data_dir=os.path.join(model_dir, 'espeak-ng-data'),
                                                   noise_scale=0.33, noise_scale_w=0.4),
        num_threads=4))
    tts = sherpa_onnx.OfflineTts(cfg)
    for t in todo:
        h = hashlib.md5(t.encode()).hexdigest()
        a = tts.generate(t, sid=0, speed=speed)
        sf.write(os.path.join(cache_dir, h + '.wav'), a.samples, a.sample_rate)
        durations[h] = round(len(a.samples) / a.sample_rate, 3)
        print(f'{durations[h]:5.1f}s  {t}')
    json.dump(durations, open(dur_path, 'w'), ensure_ascii=False, indent=0)
print(f'{len(texts)}개 문장, 새로 만든 음성 {len(todo)}개')
