"""
贷后数据看板 - 后端权限管理系统 v3
启动: python server.py
      或双击 启动看板服务器.bat

权限模型：
  admin          — 全部板块 + 全部数据（总负责人）
  region_manager — 全部板块 + 仅自己地区数据（地区负责人）
  group_leader   — 组长板块 + 仅自己小组数据（组长）
  staff          — 受限板块 + 可选数据约束（员工）
"""
import json
import os
import hashlib
import secrets
from datetime import timedelta
from functools import wraps
from flask import Flask, request, jsonify, session, send_from_directory

app = Flask(__name__)
app.secret_key = 'dashboard_v3_2026_fixed_secret_key_loan_collection'  # 固定密钥，重启不丢登录态
app.permanent_session_lifetime = timedelta(hours=8)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, 'users.json')
AUDIT_LOG  = os.path.join(BASE_DIR, 'audit.log')

# ============================================================
# 角色 → 板块/功能 权限配置
# ============================================================
ROLE_PERMISSIONS = {
    'admin': {
        'tabs': [
            'overview', 'ranking', 'today', 'group', 'detail',
            'company-rank', 'region-rank', 'group-rank',
            'first-collection', 'staff-productivity', 'company-productivity'
        ],
        'label':  '👑 总负责人',
        'filter': True,   # 显示筛选栏
        'kpi':    True,   # 显示 KPI 卡片
    },
    'region_manager': {
        'tabs': [
            'overview', 'ranking', 'today', 'group', 'detail',
            'company-rank', 'region-rank', 'group-rank',
            'first-collection', 'staff-productivity', 'company-productivity'
        ],
        'label':  '🏙️ 地区负责人',
        'filter': True,
        'kpi':    True,
    },
    'group_leader': {
        'tabs': [
            'overview', 'ranking', 'today', 'group', 'detail',
            'group-rank', 'staff-productivity'
        ],
        'label':  '⭐ 组长',
        'filter': True,
        'kpi':    True,
    },
    'staff': {
        'tabs': ['overview', 'today', 'detail'],
        'label':  '👤 员工',
        'filter': False,
        'kpi':    False,
    },
}

# ============================================================
# 工具函数
# ============================================================
def load_users():
    with open(USERS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_users(data):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def write_audit(username, action, ip='-'):
    from datetime import datetime
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(AUDIT_LOG, 'a', encoding='utf-8') as f:
        f.write(f'[{ts}] {ip} | {username} | {action}\n')

# ============================================================
# 装饰器
# ============================================================
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return jsonify({'error': '未登录', 'code': 'AUTH_REQUIRED'}), 401
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return jsonify({'error': '未登录'}), 401
        if session.get('role') != 'admin':
            return jsonify({'error': '仅管理员可操作'}), 403
        return f(*args, **kwargs)
    return decorated

# ============================================================
# 页面路由
# ============================================================
@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

# ============================================================
# 认证 API
# ============================================================
@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json(silent=True) or {}
    username = (data.get('username') or '').strip()
    password = (data.get('password') or '').strip()

    if not username or not password:
        return jsonify({'success': False, 'error': '请输入账号和密码'}), 400

    users = load_users()
    user = users.get(username)

    if not user:
        write_audit(username, 'FAIL:用户不存在', request.remote_addr)
        return jsonify({'success': False, 'error': '账号或密码错误'}), 401

    if not user.get('enabled', True):
        write_audit(username, 'FAIL:账号已禁用', request.remote_addr)
        return jsonify({'success': False, 'error': '账号已被禁用'}), 403

    if hash_password(password) != user.get('password', ''):
        write_audit(username, 'FAIL:密码错误', request.remote_addr)
        return jsonify({'success': False, 'error': '账号或密码错误'}), 401

    # 登录成功
    session.permanent = True
    session['user']         = username
    session['role']         = user['role']
    session['display_name'] = user.get('display_name', username)
    session['data_constraint'] = user.get('data_constraint', {'region': None, 'group': None})

    role_cfg = ROLE_PERMISSIONS.get(user['role'], ROLE_PERMISSIONS['staff'])

    write_audit(username, 'LOGIN_OK', request.remote_addr)

    return jsonify({
        'success':      True,
        'username':     username,
        'display_name': user.get('display_name', username),
        'role':         user['role'],
        'permissions':  role_cfg,
        'data_constraint': user.get('data_constraint', {'region': None, 'group': None}),
    })

@app.route('/api/logout', methods=['POST'])
def logout():
    username = session.get('user', '?')
    write_audit(username, 'LOGOUT', request.remote_addr)
    session.clear()
    return jsonify({'success': True})

@app.route('/api/user')
def get_user():
    """前端检查登录状态 + 获取权限 + 数据约束"""
    if 'user' not in session:
        return jsonify({'logged_in': False}), 401

    role = session.get('role', 'staff')
    role_cfg = ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS['staff'])

    return jsonify({
        'logged_in':      True,
        'username':       session['user'],
        'display_name':   session.get('display_name', ''),
        'role':           role,
        'permissions':    role_cfg,
        'data_constraint': session.get('data_constraint', {'region': None, 'group': None}),
    })

# ============================================================
# 用户管理 API (仅管理员)
# ============================================================
@app.route('/api/users', methods=['GET'])
@admin_required
def list_users():
    users = load_users()
    result = []
    for uname, u in users.items():
        dc = u.get('data_constraint', {})
        result.append({
            'username':      uname,
            'display_name':  u.get('display_name', uname),
            'role':          u.get('role', 'staff'),
            'enabled':       u.get('enabled', True),
            'data_constraint': dc,
        })
    return jsonify({'users': result})

@app.route('/api/users', methods=['POST'])
@admin_required
def add_user():
    data = request.get_json(silent=True) or {}
    username     = (data.get('username') or '').strip()
    password     = (data.get('password') or '').strip()
    role         = data.get('role', 'staff')
    display_name = data.get('display_name', username)
    dc           = data.get('data_constraint', {'region': None, 'group': None})

    if not username or not password:
        return jsonify({'success': False, 'error': '账号和密码不能为空'}), 400
    if role not in ROLE_PERMISSIONS:
        return jsonify({'success': False, 'error': '无效的角色'}), 400

    users = load_users()
    if username in users:
        return jsonify({'success': False, 'error': '账号已存在'}), 409

    users[username] = {
        'password':     hash_password(password),
        'role':         role,
        'display_name': display_name,
        'enabled':      True,
        'data_constraint': dc,
    }
    save_users(users)
    write_audit(session['user'], f'ADD_USER:{username}:{role}:{dc.get("region","all")}:{dc.get("group","all")}', request.remote_addr)
    return jsonify({'success': True, 'message': f'用户 {username} 创建成功'})

@app.route('/api/users/<username>', methods=['PUT'])
@admin_required
def update_user(username):
    users = load_users()
    if username not in users:
        return jsonify({'success': False, 'error': '用户不存在'}), 404

    data = request.get_json(silent=True) or {}
    if 'password' in data and data['password']:
        users[username]['password'] = hash_password(data['password'])
    if 'role' in data:
        if data['role'] not in ROLE_PERMISSIONS:
            return jsonify({'success': False, 'error': '无效的角色'}), 400
        users[username]['role'] = data['role']
    if 'display_name' in data:
        users[username]['display_name'] = data['display_name']
    if 'enabled' in data:
        users[username]['enabled'] = data['enabled']
    if 'data_constraint' in data:
        users[username]['data_constraint'] = data['data_constraint']

    save_users(users)
    write_audit(session['user'], f'UPDATE_USER:{username}', request.remote_addr)
    return jsonify({'success': True, 'message': f'用户 {username} 已更新'})

@app.route('/api/users/<username>', methods=['DELETE'])
@admin_required
def delete_user(username):
    if username == session['user']:
        return jsonify({'success': False, 'error': '不能删除自己'}), 400

    users = load_users()
    if username not in users:
        return jsonify({'success': False, 'error': '用户不存在'}), 404

    del users[username]
    save_users(users)
    write_audit(session['user'], f'DEL_USER:{username}', request.remote_addr)
    return jsonify({'success': True, 'message': f'用户 {username} 已删除'})

# ============================================================
# 审计日志 (仅管理员)
# ============================================================
@app.route('/api/audit', methods=['GET'])
@admin_required
def get_audit():
    try:
        with open(AUDIT_LOG, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        return jsonify({'logs': lines[-500:]})
    except FileNotFoundError:
        return jsonify({'logs': []})

# ============================================================
# 静态文件
# ============================================================
@app.route('/<path:path>')
def static_files(path):
    if path in ('users.json', 'audit.log', 'server.py'):
        return jsonify({'error': '禁止访问'}), 403
    return send_from_directory(BASE_DIR, path)

# ============================================================
# 启动
# ============================================================
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='贷后数据看板 v3 - 数据级权限')
    parser.add_argument('--port',  type=int, default=5000)
    parser.add_argument('--host',  default='0.0.0.0')
    parser.add_argument('--debug', action='store_true')
    args = parser.parse_args()

    print(f'''
+==================================================+
|     贷后数据看板 v3 -- 数据级权限版             |
+==================================================+
|  本地:     http://127.0.0.1:{args.port:<5}                   |
|  局域网:   http://服务器IP:{args.port:<5}                   |
+==================================================+
|  [总负责] admin    / admin123  -> 全数据+全板块 |
|  [地区]   zz_mgr   / admin123  -> 仅郑州数据    |
|  [地区]   cd1_mgr  / admin123  -> 仅成都一数据  |
|  [地区]   cd2_mgr  / admin123  -> 仅成都二数据  |
|  [组长]   leader1  / leader123 -> CD-F1小组     |
|  [组长]   leader2  / leader123 -> ZZ-A小组      |
|  [员工]   staff1   / staff123  -> 员工视图      |
+==================================================+
|  用户配置: users.json                           |
|  审计日志: audit.log                            |
|  (请立即修改默认密码!)                          |
+==================================================+
''')
    app.run(host=args.host, port=args.port, debug=args.debug)
