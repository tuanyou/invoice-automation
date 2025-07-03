# -*- coding: utf-8 -*-
from fastapi import FastAPI, Body
from fastapi.responses import FileResponse
import uvicorn
import os
from 为途发票填写自动化 import run  # 假设原始脚本叫做这个

app = FastAPI()

@app.post("/generate-invoice")
def generate_invoice():
    try:
        run()  # 原函数里会批量读取并处理所有sheet

        # 示例：返回最新生成的文件路径（实际可根据生成逻辑动态定位）
        latest_file = get_latest_invoice_file()
        return FileResponse(latest_file, filename=os.path.basename(latest_file))

    except Exception as e:
        return {"status": "error", "message": str(e)}

def get_latest_invoice_file():
    dir_path = r'D:\work\data\发票\为途'
    files = [os.path.join(dir_path, f) for f in os.listdir(dir_path) if f.endswith('.xlsx')]
    if not files:
        raise FileNotFoundError("未找到发票文件")
    return max(files, key=os.path.getctime)

if __name__ == '__main__':
    uvicorn.run(app, host="127.0.0.1", port=8000)