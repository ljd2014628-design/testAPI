# -*- coding: utf-8 -*-
"""本地假接口服务，用于演示和自测。

运行：python tools/mock_server.py   监听 http://127.0.0.1:9911
在工具里把 BaseURL 填成该地址即可跑通全部示例用例。
其中 /coupon/receive 每调用一次库存减 1，用来演示「较上次减少 1」的断言。
"""

from flask import Flask, request, jsonify
app = Flask(__name__)

@app.route("/marketing/promotion/findAll", methods=["POST"])
def f():
    body = request.get_json(force=True, silent=True) or {}
    return jsonify({"code": 0, "data": {"promotionId": "P9527",
                                        "lines": [{"lineId": "30", "couponInstanceId": "CI-888"}],
                                        "echoHeader": request.headers.get("X-Trace-Id"),
                                        "auth": request.headers.get("Authorization"),
                                        "orderNo": body.get("orderNo")}})

@app.route("/coupon/query/<cid>", methods=["GET"])
def q(cid):
    return jsonify({"code": 0, "data": {"id": cid, "status": "USED",
                                        "brandCode": request.args.get("brandCode"),
                                        "trace": request.headers.get("X-Trace-Id"),
                                        "ctype": request.headers.get("Content-Type")}})

@app.route("/marketing/promotion/change", methods=["POST"])
def c():
    return jsonify({"code": 0, "data": {"received": request.get_json(force=True, silent=True),
                                        "tenant": request.headers.get("tenant")}})

@app.route("/big", methods=["POST", "GET"])
def big():
    return jsonify({"code": 0, "data": {"rows": [{"id": i, "name": "记录%d" % i,
                    "desc": "这是一条用于测试长内容折叠的记录" * 3} for i in range(200)]}})

@app.route("/coupon/wallet/list", methods=["POST"])
def wallet():
    return jsonify({"code": 0, "msg": "ok", "data": {"total": 3, "list": [
        {"couponNo": "CP1001", "cuonc": "满100减20", "status": 1, "amount": 20.0},
        {"couponNo": "CP1002", "cuonc": "8折券", "status": 1, "amount": None},
        {"couponNo": "CP1003", "cuonc": "新人券", "status": 0, "amount": 5}],
        "token": request.headers.get("authorization")}})

STOCK = {"remain": 10}

@app.route("/coupon/receive", methods=["POST"])
def receive():
    if STOCK["remain"] > 0:
        STOCK["remain"] -= 1
    return jsonify({"code": 0, "msg": "ok", "data": {"remain": STOCK["remain"],
                                                     "couponNo": "CP%04d" % STOCK["remain"]}})

app.run(port=9911)
