"""图片转 Excel：Streamlit + DeepSeek Vision 的轻量级交付工具。

授权文件格式（UTF-8，每行一条）：
    授权码
授权码首次验证时开始计算 30 天，到期信息保存在 authorized_activations.json。
也兼容“授权码,YYYY-MM-DD”这种固定到期日期格式。

启动：streamlit run app.py
"""

from __future__ import annotations

import base64
import json
import os
import re
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, UnidentifiedImageError


# ------------------------------ 基础配置 ------------------------------
APP_TITLE = "图片转 Excel"
API_URL = "https://api.deepseek.com/chat/completions"
DEFAULT_MODEL = "deepseek-flash"  # DeepSeek 官方视觉模型
AUTH_FILE = os.environ.get("AUTHORIZED_USERS_FILE", "authorized_users.txt")
ACTIVATION_FILE = os.environ.get("AUTHORIZED_ACTIVATIONS_FILE", "authorized_activations.json")
LICENSE_DAYS = 30
MAX_IMAGE_MB = 20  # 留出 Base64 编码后的空间，避免触及接口请求体限制
SUPPORTED_FORMATS = {"PNG": "image/png", "JPEG": "image/jpeg", "WEBP": "image/webp", "GIF": "image/gif"}

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")


def load_authorized_users(path: str) -> dict[str, date | None]:
    """读取本地授权码。格式为“码”或“码,YYYY-MM-DD”。"""
    users: dict[str, date | None] = {}
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            for line_no, raw in enumerate(f, start=1):
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                parts = [part.strip() for part in line.split(",", maxsplit=1)]
                code = parts[0]
                if not code:
                    continue
                expiry = None
                if len(parts) == 2 and parts[1]:
                    try:
                        expiry = date.fromisoformat(parts[1])
                    except ValueError:
                        st.warning(f"授权文件第 {line_no} 行的日期格式无效，应为 YYYY-MM-DD；该条已跳过。")
                        continue
                users[code] = expiry
    except FileNotFoundError:
        pass
    except OSError as exc:
        st.error(f"无法读取授权文件 {path}：{exc}")
    return users


def load_activations() -> dict[str, str]:
    """读取首次激活时间记录。"""
    try:
        with open(ACTIVATION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError):
        st.error(f"无法读取激活记录 {ACTIVATION_FILE}，请检查文件是否损坏。")
        return {}


def save_activations(activations: dict[str, str]) -> None:
    """原子保存激活记录，避免中途写入导致文件损坏。"""
    temp_path = ACTIVATION_FILE + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump(activations, f, ensure_ascii=False, indent=2)
    os.replace(temp_path, ACTIVATION_FILE)


def is_authorized(code: str, users: dict[str, date | None], demo_enabled: bool) -> tuple[bool, str, str | None]:
    """校验授权码；纯授权码自首次使用起有效 30 天。"""
    entered = code.strip()
    if not entered:
        return False, "请输入授权码。", None
    if demo_enabled and entered == "DEMO-2026":
        return True, "演示授权有效。", None
    expiry = users.get(entered, "missing")
    if expiry == "missing":
        return False, "授权码无效，请核对后重试。", None
    if expiry is not None and expiry < date.today():
        return False, f"授权码已于 {expiry.isoformat()} 到期。", None

    activations = load_activations()
    if expiry is None:
        # 文件中只写码时，以首次使用时间起算 30 天；之后重复登录不重置期限。
        activated_at = activations.get(entered)
        if not activated_at:
            activated_at = datetime.now().isoformat(timespec="seconds")
            activations[entered] = activated_at
            try:
                save_activations(activations)
            except OSError as exc:
                return False, f"无法保存授权激活记录：{exc}", None
        try:
            ends_at = datetime.fromisoformat(activated_at) + timedelta(days=LICENSE_DAYS)
        except ValueError:
            return False, "该授权码的激活记录格式异常，请联系服务提供者。", None
        if datetime.now() >= ends_at:
            return False, f"授权码已于 {ends_at:%Y-%m-%d %H:%M} 到期。", None
        return True, f"授权有效至 {ends_at:%Y-%m-%d %H:%M}。", ends_at.isoformat(timespec="seconds")
    ends_at = datetime.combine(expiry + timedelta(days=1), datetime.min.time())
    return True, f"授权有效期至 {expiry.isoformat()}。", ends_at.isoformat(timespec="seconds")


def extract_json_object(text: str) -> dict[str, Any]:
    """提取模型返回的 JSON 对象，兼容 Markdown 代码围栏。"""
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError:
        # 有时模型会在 JSON 前后加说明文字；尝试截取首尾大括号之间的内容。
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start < 0 or end <= start:
            raise ValueError("模型没有返回有效 JSON，请尝试裁剪清晰的表格区域后重试。")
        try:
            parsed = json.loads(cleaned[start : end + 1])
        except json.JSONDecodeError as exc:
            raise ValueError(f"模型返回的 JSON 格式不完整：{exc}") from exc
    if not isinstance(parsed, dict):
        raise ValueError("模型返回内容不是 JSON 对象。")
    return parsed


def parse_image_with_deepseek(image_bytes: bytes, mime_type: str, api_key: str, model: str) -> pd.DataFrame:
    """将图片以内嵌 Base64 形式发送给 DeepSeek，并将 JSON 表格整理为 DataFrame。"""
    encoded = base64.b64encode(image_bytes).decode("ascii")
    data_url = f"data:{mime_type};base64,{encoded}"
    prompt = """请仔细识别图片中的表格，并只返回合法 JSON，不要 Markdown 或解释文字。
JSON 结构必须为：{"columns":["列名1","列名2"],"rows":[["单元格","单元格"],["单元格","单元格"]]}。
要求：
1. 按图片的原始行列顺序完整提取表头和所有可辨认单元格，不要省略空行以外的数据。
2. 合并单元格内容只放在其左上角对应单元格；其他位置填空字符串。
3. 数字、日期、金额尽量忠实保留原文，不要自行推测或改写；无法辨认的单元格写“[无法识别]”。
4. 每一行的单元格数量必须与 columns 数量完全一致；没有清晰表头时可使用“列1”“列2”等。
5. 图片若含多个表格，请将内容按视觉顺序合并成一个表，并确保列结构合理。
"""
    try:
        response = requests.post(
            API_URL,
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": model,
                "messages": [{"role": "user", "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": data_url, "detail": "original"}},
                ]}],
                "stream": False,
                "temperature": 0,
            },
            timeout=(15, 180),
        )
    except requests.Timeout as exc:
        raise RuntimeError("连接 DeepSeek 超时，请稍后重试。") from exc
    except requests.RequestException as exc:
        raise RuntimeError(f"连接 DeepSeek 失败：{exc}") from exc

    if not response.ok:
        try:
            detail = response.json().get("error", {}).get("message", response.text[:500])
        except (ValueError, AttributeError):
            detail = response.text[:500]
        raise RuntimeError(f"DeepSeek API 返回 HTTP {response.status_code}：{detail}")
    try:
        body = response.json()
        content = body["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("DeepSeek 返回内容为空或格式异常。") from exc
    if not isinstance(content, str) or not content.strip():
        raise RuntimeError("模型未返回可读取的文本内容。")

    result = extract_json_object(content)
    columns, rows = result.get("columns"), result.get("rows")
    if not isinstance(columns, list) or not columns:
        raise ValueError("解析结果缺少有效的 columns 列名列表。")
    if not isinstance(rows, list):
        raise ValueError("解析结果缺少有效的 rows 数据列表。")
    columns = [str(item) if item is not None else "" for item in columns]
    normalized_rows: list[list[Any]] = []
    for row in rows:
        if not isinstance(row, list):
            continue
        values = list(row[: len(columns)])
        values.extend([""] * (len(columns) - len(values)))
        normalized_rows.append(["" if value is None else value for value in values])
    if not normalized_rows:
        raise ValueError("图片中没有识别到数据行，请上传更清晰的表格图片。")
    return pd.DataFrame(normalized_rows, columns=columns)


def dataframe_to_excel(df: pd.DataFrame) -> bytes:
    """生成带表头、筛选和冻结首行的整洁 xlsx 文件。"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="识别结果")
        ws = writer.sheets["识别结果"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        header_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(name="微软雅黑", color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name="微软雅黑", size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        for col_idx, column in enumerate(df.columns, start=1):
            values = [str(column)] + ["" if pd.isna(v) else str(v) for v in df[column].head(200)]
            width = min(max(max((len(v) for v in values), default=8) + 2, 10), 36)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        ws.row_dimensions[1].height = 24
    return output.getvalue()


# ------------------------------ 页面界面 ------------------------------
st.title("📊 图片转 Excel")
st.caption("上传表格截图或单据图片，AI 识别后即可预览并下载可编辑的 Excel。")

with st.sidebar:
    st.header("服务设置")
    env_key = os.environ.get("DEEPSEEK_API_KEY", "")
    try:
        env_key = st.secrets.get("DEEPSEEK_API_KEY", env_key)
    except Exception:
        # 未配置 secrets.toml 时，部分 Streamlit 版本会抛出专用异常。
        pass
    api_key = st.text_input(
        "DeepSeek API Key", value=env_key, type="password",
        help="也可通过环境变量 DEEPSEEK_API_KEY 或 Streamlit secrets.toml 预设。密钥仅用于本次服务端请求。",
    ).strip()
    model = st.text_input("视觉模型", value=DEFAULT_MODEL, help="默认使用 DeepSeek 官方支持图片输入的 deepseek-flash。")
    st.divider()
    demo_enabled = st.toggle("启用演示授权码", value=False, help="演示码为 DEMO-2026；正式运营请保持关闭。")
    st.caption(f"授权文件：`{AUTH_FILE}`")

if "authorized" not in st.session_state:
    st.session_state.authorized = False
if st.session_state.authorized and st.session_state.get("auth_expires_at"):
    try:
        if datetime.now() >= datetime.fromisoformat(st.session_state.auth_expires_at):
            st.session_state.authorized = False
            st.session_state.pop("license_message", None)
            st.session_state.pop("auth_expires_at", None)
            st.warning("授权已到期，请重新输入有效授权码。")
    except ValueError:
        st.session_state.authorized = False

if not st.session_state.authorized:
    st.subheader("🔐 输入授权码")
    st.write("完成闲鱼下单后，输入收到的授权码即可开始使用。")
    with st.form("license_form"):
        license_code = st.text_input("授权码", placeholder="请输入授权码", type="password")
        submitted = st.form_submit_button("验证并进入", type="primary")
    if submitted:
        ok, message, auth_expires_at = is_authorized(license_code, load_authorized_users(AUTH_FILE), demo_enabled)
        if ok:
            st.session_state.authorized = True
            st.session_state.license_message = message
            st.session_state.auth_expires_at = auth_expires_at
            st.rerun()
        else:
            st.error(message)
    st.stop()

st.success(st.session_state.get("license_message", "授权有效。"))
if st.button("退出授权"):
    st.session_state.authorized = False
    st.session_state.pop("license_message", None)
    st.session_state.pop("auth_expires_at", None)
    st.rerun()

if not api_key:
    st.warning("请在左侧输入 DeepSeek API Key，或先配置环境变量后再进行图片识别。")

uploaded = st.file_uploader("上传图片", type=["png", "jpg", "jpeg", "webp", "gif"], help=f"支持 PNG、JPG、WEBP、GIF，单张不超过 {MAX_IMAGE_MB} MB。")
if uploaded is not None:
    image_bytes = uploaded.getvalue()
    if len(image_bytes) > MAX_IMAGE_MB * 1024 * 1024:
        st.error(f"图片超过 {MAX_IMAGE_MB} MB，请压缩或裁剪后重新上传。")
        st.stop()
    try:
        image = Image.open(BytesIO(image_bytes))
        image_format = (image.format or "").upper()
        if image_format not in SUPPORTED_FORMATS:
            st.error("图片格式不受支持，请上传 PNG、JPG、WEBP 或 GIF。")
            st.stop()
        mime_type = SUPPORTED_FORMATS[image_format]
        st.image(image, caption=uploaded.name, use_container_width=True)
    except (UnidentifiedImageError, OSError):
        st.error("无法读取该图片，请确认文件完整且格式正确。")
        st.stop()

    if st.button("✨ 开始识别并生成表格", type="primary", disabled=not bool(api_key)):
        with st.spinner("正在调用 DeepSeek 视觉模型识别，请稍候…"):
            try:
                st.session_state.result_df = parse_image_with_deepseek(image_bytes, mime_type, api_key, model.strip() or DEFAULT_MODEL)
                st.session_state.source_name = uploaded.name
            except (RuntimeError, ValueError) as exc:
                st.error(str(exc))

if "result_df" in st.session_state:
    df = st.session_state.result_df
    st.subheader("识别结果预览")
    st.caption(f"共 {len(df)} 行，{len(df.columns)} 列。请下载前检查识别结果；图片模糊、倾斜或遮挡可能导致误识别。")
    st.dataframe(df, use_container_width=True, hide_index=True)
    try:
        excel_data = dataframe_to_excel(df)
        base_name = os.path.splitext(st.session_state.get("source_name", "识别结果"))[0]
        st.download_button(
            "⬇️ 下载 Excel 文件",
            data=excel_data,
            file_name=f"{base_name}_识别结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as exc:
        st.error(f"生成 Excel 文件失败：{exc}")

st.divider()
st.caption("提示：请勿上传包含敏感个人信息、银行卡信息或商业机密的图片。上传图片会发送至 DeepSeek API 进行识别。")
