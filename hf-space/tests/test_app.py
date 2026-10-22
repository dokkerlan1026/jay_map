DATE = "2026-10-23"


def test_health_and_index(client):
    assert client.get("/health").json() == {"status": "ok"}
    assert client.get("/").status_code == 200


def test_store_workflow(client):
    initial = client.get(f"/api/stores?date={DATE}")
    assert initial.status_code == 200
    assert initial.json()["activeCount"] == 500

    hidden = client.post("/api/stores/toggle-visibility", json={
        "date": DATE, "store_id": 1, "is_visible": False
    })
    assert hidden.status_code == 200

    refreshed = client.get(f"/api/stores?date={DATE}").json()
    assert refreshed["activeCount"] == 499
    assert refreshed["stores"][0]["isVisible"] is False


def test_checkin_toggle(client):
    client.get(f"/api/stores?date={DATE}")
    res = client.post("/api/stores/toggle-checkin", json={
        "date": DATE, "store_id": 2, "is_checked_in": True
    })
    assert res.status_code == 200
    assert res.json()["checkinTime"]
    data = client.get(f"/api/stores?date={DATE}").json()
    assert data["checkedCount"] == 1
    assert data["stores"][1]["isCheckedIn"] is True


def test_save_and_revert_named_version(client):
    client.get(f"/api/stores?date={DATE}")
    saved = client.post("/api/versions/save", json={
        "date": DATE, "version_name": "下午路線", "hidden_ids": [1, 2, 3], "checked_ids": [4]
    }).json()
    assert isinstance(saved["versionId"], int)
    assert saved["activeCount"] == 497

    client.post("/api/stores/load-all-open", params={"date": DATE})
    assert client.get(f"/api/stores?date={DATE}").json()["activeCount"] == 500

    reverted = client.post("/api/versions/revert", json={"date": DATE, "version_id": saved["versionId"]})
    assert reverted.status_code == 200
    data = client.get(f"/api/stores?date={DATE}").json()
    assert data["activeCount"] == 497
    assert data["checkedCount"] == 1

    versions = client.get("/api/versions").json()
    assert any(v["versionName"] == "下午路線" for v in versions["namedList"])


def test_batch_poll_and_export(client):
    client.get(f"/api/stores?date={DATE}")
    client.post("/api/stores/batch-visibility", json={"date": DATE, "hidden_ids": [5, 6]})
    poll = client.get("/api/stores/poll", params={"date": DATE}).json()
    assert sorted(poll["hiddenIds"]) == [5, 6]
    export = client.get("/api/export-excel", params={"date": DATE})
    assert export.status_code == 200
    assert export.content[:2] == b"PK"


def test_revert_missing_version_returns_404(client):
    res = client.post("/api/versions/revert", json={"date": DATE, "version_id": 9999})
    assert res.status_code == 404
