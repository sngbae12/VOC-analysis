# -*- coding: utf-8 -*-
from __future__ import annotations

from fastapi.testclient import TestClient
from gradio.routes import App

import voc_analysis_app as app


def test_gradio_home_page_loads():
    demo = app.build_ui()
    fastapi_app = App.create_app(demo)
    with TestClient(fastapi_app) as client:
        response = client.get("/config")
        assert response.status_code == 200
        payload = response.json()
        assert isinstance(payload, dict)
