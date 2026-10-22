"""黒塗りの範囲を決める一時サーバー（mask_server.py）"""
import http.client
import json
import threading
import urllib.error
import urllib.request

import pytest

from saiten_llm.mask import mask_status
from saiten_llm.mask_server import MaskServer
from saiten_llm.workspace import Project


@pytest.fixture
def server(oriented_project):
    srv = MaskServer(Project(oriented_project), timeout_minutes=1)
    thread = threading.Thread(target=srv.serve_until_done, kwargs={"poll": 0.05}, daemon=True)
    thread.start()
    yield srv
    srv.approved.set()
    thread.join(5)


def _get(srv, path, token=True):
    sep = "&" if "?" in path else "?"
    url = f"http://127.0.0.1:{srv.port}{path}" + (f"{sep}t={srv.token}" if token else "")
    return urllib.request.urlopen(url, timeout=10)


def _post(srv, path, body):
    req = urllib.request.Request(
        f"http://127.0.0.1:{srv.port}{path}?t={srv.token}",
        data=json.dumps(body).encode(), headers={"Content-Type": "application/json"})
    return json.loads(urllib.request.urlopen(req, timeout=10).read())


def test_listens_only_on_localhost(server):
    assert server.httpd.server_address[0] == "127.0.0.1"
    assert server.url.startswith("http://127.0.0.1:")


def test_token_is_required(server):
    with pytest.raises(urllib.error.HTTPError) as e:
        _get(server, "/api/state", token=False)
    assert e.value.code == 403
    with pytest.raises(urllib.error.HTTPError) as e:
        _get(server, "/img?i=0&t=nottherighttoken", token=False)
    assert e.value.code == 403


def test_other_host_header_is_refused(server):
    """DNS リバインディング対策: 127.0.0.1 / localhost 以外の Host は断る"""
    conn = http.client.HTTPConnection("127.0.0.1", server.port, timeout=10)
    conn.request("GET", f"/api/state?t={server.token}", headers={"Host": "evil.example"})
    assert conn.getresponse().status == 403


def test_page_and_images(server):
    html = _get(server, "/").read().decode("utf-8")
    assert server.token in html and "黒塗りの範囲" in html
    state = json.loads(_get(server, "/api/state").read())
    assert len(state["pages"]) == 6
    img = _get(server, "/img?i=0&w=300&m=1")
    assert img.headers["Content-Type"] == "image/jpeg"
    assert img.headers["Cache-Control"] == "no-store"
    with pytest.raises(urllib.error.HTTPError) as e:
        _get(server, "/img?i=99")
    assert e.value.code == 404


def test_bad_spec_is_rejected(server):
    with pytest.raises(urllib.error.HTTPError) as e:
        _post(server, "/api/save", {"regions": [[0, 0, "x", 1]]})
    assert e.value.code == 400
    with pytest.raises(urllib.error.HTTPError) as e:
        _post(server, "/api/save", {"regions": [], "extra": {"../../etc/passwd": [[0, 0, 1, 1]]}})
    assert e.value.code == 400


def test_save_then_approve_stops_server(server):
    project = server.project
    spec = {"regions": [[0.24, 0.2, 0.57, 0.255]], "extra": {}}
    assert _post(server, "/api/save", spec)["ok"]
    assert mask_status(Project(project.root)) == "draft"
    assert _post(server, "/api/approve", spec)["ok"]
    assert server.approved.wait(5)
    assert mask_status(Project(project.root)) == "approved"
