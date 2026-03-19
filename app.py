"""
MedAI Suite — PPTX Render Service
Railway deployment: wraps renderer.py as a simple HTTP API
"""

import os
import json
import base64
import traceback
from flask import Flask, request, jsonify
from renderer import build_pptx

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'medai-pptx-renderer'})

@app.route('/render', methods=['POST', 'OPTIONS'])
def render():
    # CORS preflight
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin']  = '*'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        return resp

    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({'error': 'No JSON payload'}), 400

        pptx_bytes = build_pptx(payload)
        b64 = base64.b64encode(pptx_bytes).decode('utf-8')

        resp = jsonify({'pptx': b64, 'size': len(pptx_bytes)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp

    except Exception as e:
        tb = traceback.format_exc()
        print(f"ERROR: {e}\n{tb}")
        resp = jsonify({'error': str(e), 'traceback': tb})
        resp.status_code = 500
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
