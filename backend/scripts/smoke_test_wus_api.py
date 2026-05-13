#!/usr/bin/env python3

import json
import os
import sys
from typing import Any, Optional
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


BASE_URL = os.getenv("QUOTE_API_BASE_URL", "http://127.0.0.1:8000").rstrip("/")

SCENARIOS = [
    {
        "name": "wus-female-69-face-amount",
        "payload": {
            "productCode": "WUS",
            "financingType": "premium",
            "gender": "female",
            "age": 69,
            "coverageYear": 20,
            "faceAmount": 1000000,
            "ltvRatio": 0.5,
            "dividendOption": "第7年起持續增購保額",
        },
        "expect": {
            "productCode": "WUS",
            "coverageYear": 20,
            "currency": "USD",
        },
    },
    {
        "name": "wus-male-19-premium",
        "payload": {
            "productCode": "WUS",
            "financingType": "premium",
            "gender": "male",
            "age": 19,
            "coverageYear": 20,
            "premium": 563168,
            "ltvRatio": 0.5,
            "dividendOption": "第7年起持續增購保額",
        },
        "expect": {
            "productCode": "WUS",
            "coverageYear": 20,
            "currency": "USD",
        },
    },
]


def request_json(path: str, payload: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    request = Request(
        f"{BASE_URL}{path}",
        data=data,
        headers={"content-type": "application/json"} if payload is not None else {},
        method="POST" if payload is not None else "GET",
    )
    with urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))


def main() -> int:
    try:
        health = request_json("/health")
        print(f"[health] {json.dumps(health, ensure_ascii=False)}")
    except (HTTPError, URLError) as exc:
        print(f"Health check failed: {exc}", file=sys.stderr)
        return 1

    failures = 0
    for scenario in SCENARIOS:
        try:
            response = request_json("/api/v1/quotes", scenario["payload"])
        except (HTTPError, URLError) as exc:
            failures += 1
            print(f"[fail] {scenario['name']}: request error {exc}", file=sys.stderr)
            continue

        for field, expected in scenario["expect"].items():
            actual = response.get(field)
            if actual != expected:
                failures += 1
                print(
                    f"[fail] {scenario['name']}: field {field} expected {expected!r} got {actual!r}",
                    file=sys.stderr,
                )
                break
        else:
            engine = response.get("engine", "<unknown>")
            projected_benefit = response.get("projectedBenefit", "<missing>")
            premium = response.get("premium", "<missing>")
            print(
                f"[pass] {scenario['name']}: engine={engine} premium={premium} projectedBenefit={projected_benefit}"
            )

    if failures:
        print(f"Smoke test completed with {failures} failure(s).", file=sys.stderr)
        return 1

    print("Smoke test passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
