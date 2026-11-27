# 스마트 판서 매뉴얼 영상 녹화 도구

`smart-board-guide.html`에 들어가는 기능별 영상(`assets/smartboard_guide/video/*.mp4`, 포스터 `*.jpg`)을 만드는 도구입니다.
실제 스마트 판서 화면을 Playwright로 자동 조작하며 녹화하고, 가짜 커서·클릭 표시·자막·단축키 배지를 화면에 함께 띄웁니다.

- `chapters.js` — 장마다 자막과 조작 순서(대본)
- `demo-lib.js` — 커서·자막·강조 표시와 그리기 도우미
- `record.js` — 녹화 후 ffmpeg로 mp4(H.264 + AAC 음성 해설)와 포스터를 만듭니다
- `tts.py` — 인터넷 없이 한국어 음성(sherpa-onnx + mimic3 KSS 모델)을 만듭니다. 준비 방법은 파일 머리말 참고
- `handout.html` / `handout.js` — 단축키 핸드아웃 그림(`assets/smartboard_guide/shortcut-handout.png`)과 A4 PDF
- `samples/` — 시연용 예시 자료(활동지 PDF, 클릭 애니메이션 PPTX, 예시 웹 페이지)

```bash
python3 -m http.server 4321 &                       # 저장소 루트에서
node tools/sb-manual-video/record.js http://localhost:4321 assets/smartboard_guide/video         # 전체
node tools/sb-manual-video/record.js http://localhost:4321 assets/smartboard_guide/video c07-pdf # 한 장만

# 음성 해설까지 넣으려면 (tts.py 준비 후)
TTS_PY=.venv/bin/python TTS_MODEL=vits-mimic3-ko_KO-kss_low TTS_CACHE=/tmp/sb-tts \
  node tools/sb-manual-video/record.js http://localhost:4321 assets/smartboard_guide/video

node tools/sb-manual-video/handout.js   # 단축키 핸드아웃 PNG · PDF
```

대본은 `d.say('자막', '읽어 줄 문장')`으로 적습니다. 읽어 줄 문장은 작은따옴표 문자열이어야 `record.js`가 미리 음성을 만들고,
단축키는 '컨트롤 일', '제트'처럼 소리 나는 대로 적어야 또렷하게 읽습니다. 장 목록을 바꾸면 `smart-board-guide.html`의 `CH`·`GROUPS`도 함께 고칩니다.

스마트 판서 화면이 바뀌면 해당 장만 다시 녹화하면 됩니다. 예시 PPTX는 `samples/make_lesson_pptx.py`(python-pptx)로 다시 만들 수 있습니다.
