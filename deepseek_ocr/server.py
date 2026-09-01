"""Local CPU server for DeepSeek-OCR."""

import base64
import contextlib
import io
import os
import shutil
import tempfile
import threading
import uuid

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from transformers import AutoModel, AutoTokenizer

MODEL_NAME = os.environ.get('DEEPSEEK_OCR_MODEL', 'deepseek-ai/DeepSeek-OCR')
DOC_TO_MARKDOWN_PROMPT = '<image>\n<|grounding|>Convert the document to markdown. '
FREE_OCR_PROMPT = '<image>\nFree OCR. '
RESOLUTION_PRESETS = {
    'tiny': dict(base_size=512, image_size=512, crop_mode=False),
    'small': dict(base_size=640, image_size=640, crop_mode=False),
    'base': dict(base_size=1024, image_size=1024, crop_mode=False),
    'large': dict(base_size=1280, image_size=1280, crop_mode=False),
    'gundam': dict(base_size=1024, image_size=640, crop_mode=True),
}

app = FastAPI(title='DeepSeek-OCR Local CPU Backend')
model = None
tokenizer = None
model_lock = threading.Lock()


class OcrRequest(BaseModel):
    image_b64: str
    mode: str = 'markdown'
    resolution: str = 'gundam'


@app.on_event('startup')
def load_model():
    """Load once at startup so requests do not repeatedly initialize the model."""
    global model, tokenizer
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, trust_remote_code=True)
    model = AutoModel.from_pretrained(
        MODEL_NAME,
        _attn_implementation='eager',
        trust_remote_code=True,
        use_safetensors=True,
    ).eval()


@app.get('/health')
def health():
    if model is None:
        raise HTTPException(503, 'Model is still loading.')
    return {'status': 'ok', 'model': MODEL_NAME, 'device': 'cpu'}


@app.post('/ocr')
def ocr(req: OcrRequest):
    if model is None:
        raise HTTPException(503, 'Model is still loading.')
    if req.resolution not in RESOLUTION_PRESETS:
        raise HTTPException(400, f'Unknown resolution preset: {req.resolution}')

    prompt = DOC_TO_MARKDOWN_PROMPT if req.mode == 'markdown' else FREE_OCR_PROMPT
    preset = RESOLUTION_PRESETS[req.resolution]
    work_dir = tempfile.mkdtemp(prefix='dsocr_')
    image_path = os.path.join(work_dir, f'{uuid.uuid4().hex}.png')

    try:
        with open(image_path, 'wb') as image_file:
            image_file.write(base64.b64decode(req.image_b64))

        stdout_buffer = io.StringIO()
        result_text = ''
        # model.infer writes to stdout; inference must remain serial on CPU.
        with model_lock, contextlib.redirect_stdout(stdout_buffer):
            result_text = model.infer(
                tokenizer,
                prompt=prompt,
                image_file=image_path,
                output_path=work_dir,
                base_size=preset['base_size'],
                image_size=preset['image_size'],
                crop_mode=preset['crop_mode'],
                save_results=True,
            )

        if not isinstance(result_text, str) or not result_text.strip():
            printed = stdout_buffer.getvalue().split('=====================')
            result_text = printed[-1].strip() if len(printed) > 1 else ''

        if not result_text:
            saved_results = []
            for filename in sorted(os.listdir(work_dir)):
                if filename.endswith(('.mmd', '.md', '.txt')):
                    with open(os.path.join(work_dir, filename), encoding='utf-8') as result_file:
                        saved_results.append(result_file.read())
            result_text = '\n'.join(saved_results)

        return {'text': result_text}
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
