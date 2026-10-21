from datetime import datetime, timedelta
from hashlib import sha256
from html import escape
from pathlib import Path
from zoneinfo import ZoneInfo
import os

import streamlit as st
from quote_service import QuoteError, read_upload, request_quote, export_results

ROOT = Path(__file__).resolve().parent
st.set_page_config(page_title='Fmile 批量询价', page_icon='🚚', layout='wide')


def show_table(rows):
    # Native Arrow DLLs are blocked by this machine's application-control policy.
    columns = list(dict.fromkeys(k for row in rows for k in row))
    header = ''.join(f'<th>{escape(k)}</th>' for k in columns)
    body = []
    for row in rows:
        cells = []
        for key in columns:
            value = row.get(key, '')
            text = f'{value:,.2f}' if isinstance(value, float) else str(value)
            cells.append(f'<td>{escape(text)}</td>')
        body.append('<tr>' + ''.join(cells) + '</tr>')
    st.markdown('<div style="overflow-x:auto"><table style="width:100%;font-size:14px"><thead><tr>'
        + header + '</tr></thead><tbody>' + ''.join(body) + '</tbody></table></div>', unsafe_allow_html=True)


def setting(name):
    value = os.environ.get(name, '')
    if value:
        return value
    try:
        value = st.secrets.get(name, '')
    except st.errors.StreamlitSecretNotFoundError:
        value = ''
    if value:
        return str(value)
    env = ROOT / '.env'
    if env.exists():
        for line in env.read_text(encoding='utf-8-sig').splitlines():
            if line.startswith(name + '='):
                return line.split('=', 1)[1].strip().strip('\"\'')
    return ''


st.title('Fmile 批量询价')
st.write('上传托盘询价表，比较最低的 4 家卡司，下载回填结果。')
with st.sidebar:
    st.subheader('询价设置')
    today = datetime.now(ZoneInfo('America/Los_Angeles')).date()
    pickup = st.date_input('预计提货日期', value=today + timedelta(days=1), min_value=today)
    st.caption('所有路线使用同一提货日期。')
    st.divider()
    st.write('**报价规则**')
    st.caption('接口原价 · USD · 最低4家不同卡司')
    st.caption('单托盘尺寸：inch；单托盘重量：lbs。总重量按托盘数相乘。')
    st.caption('不添加可选附加服务。Fmile 服务端仍固定加入预约送货 APPT。')
    st.caption('保险无明确承保状态时显示“未明确”，时效缺失显示“未提供”。')

token = setting('FIMILE_API_TOKEN')
if not token:
    st.warning('尚未配置服务端 FIMILE_API_TOKEN。配置后即可询价。')
st.download_button('下载空白模板', (ROOT / '询价模板.xlsx').read_bytes(), '询价模板.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
upload = st.file_uploader('上传询价模板（.xlsx）', type=['xlsx'], help='每行一票，同一行的托盘尺寸和单托盘重量应相同。最多500票、10 MB。')
if upload is None:
    st.info('上传只需 A–G 输入列，H–W 报价结果列由系统自动补齐。')
    st.stop()
content = upload.getvalue()
try:
    workbook, shipments, validation_errors = read_upload(content)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

batch_key = sha256(content + pickup.isoformat().encode()).hexdigest()
if st.session_state.get('batch_key') != batch_key:
    st.session_state.update(batch_key=batch_key, results={}, api_errors={})
all_rows = sorted([s.row for s in shipments] + list(validation_errors))
st.subheader('上传检查')
st.write(f'共 {len(all_rows)} 票，其中 {len(shipments)} 票可询价，{len(validation_errors)} 票需修正。')
if not all_rows:
    st.warning('表格只有表头，请填写数据后重新上传。')

pending = [s for s in shipments if s.row not in st.session_state.results]
label = '开始询价' if not st.session_state.results else '继续询价 / 重试失败行'
if st.button(label, type='primary', disabled=not token or not pending, key='quote'):
    progress = st.progress(0, text='准备询价…')
    for i, shipment in enumerate(pending):
        progress.progress(i / len(pending), text=f'正在询价第 {i + 1}/{len(pending)} 票（Excel第{shipment.row}行）')
        try:
            result = request_quote(shipment, pickup.isoformat(), token)
            st.session_state.results[shipment.row] = result
            st.session_state.api_errors.pop(shipment.row, None)
        except QuoteError as exc:
            st.session_state.api_errors[shipment.row] = str(exc)
            if str(exc).startswith('HTTP 401'):
                st.error('Token 无效或客户已停用，已停止提交后续路线。')
                break
    progress.empty()
    st.rerun()

results = st.session_state.results
errors = {**validation_errors, **st.session_state.api_errors}
if results or errors:
    st.subheader('询价结果')
    st.caption(f'已成功 {len(results)} 票；失败或待修正 {len(errors)} 票。结果只保存在当前浏览器会话中。')
    incomplete = {s.row: '尚未询价' for s in shipments if s.row not in results and s.row not in errors}
    output = export_results(content, results, {**errors, **incomplete})
    st.download_button('下载询价结果 Excel', output, f'{Path(upload.name).stem}_询价结果.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    st.caption('不足4家时其余报价留空；失败行在首组报价中标注原因，原始输入保留。')
    if errors:
        with st.expander(f'查看错误明细（{len(errors)}票）', expanded=False):
            show_table([{'Excel行': row, '问题': message} for row, message in sorted(errors.items())])
