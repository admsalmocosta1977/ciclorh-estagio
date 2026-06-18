import os, json, psycopg2, psycopg2.extras
from flask import Flask, render_template, request, redirect, url_for, flash, g, jsonify, abort
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import date, datetime, timedelta
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ciclorh-dev-troque-em-producao')

login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Faça login para acessar o sistema.'
login_manager.login_message_category = 'warning'

DATABASE_URL = os.environ.get('DATABASE_URL', '')


class _Agente:
    nome = 'CICLO RH – apoio administrativo LTDA'
    cnpj = '32.075.028/0001-84'
    endereco = 'Av. Juracy Magalhães, 3340, Bloco A, sala 1104, Felícia'
    cidade = 'Vitória da Conquista – Bahia'
    email = 'adm.salmocosta@gmail.com'
    representante = 'Ildeflávio dos Santos Silva Maia'
    cargo = 'Sócio Gerente'
    cpf = '915.569.295-87'


AGENTE = _Agente()


class User(UserMixin):
    def __init__(self, id, username, nome, role):
        self.id = str(id)
        self.username = username
        self.nome = nome
        self.role = role

    @property
    def is_admin(self):
        return self.role == 'admin'


@login_manager.user_loader
def load_user(user_id):
    row = _q("SELECT * FROM usuario WHERE id = %s", (user_id,), one=True)
    if row:
        return User(row['id'], row['username'], row['nome'], row['role'])
    return None


def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return decorated


# ─── BANCO DE DADOS ───────────────────────────────────────────────────────────

def _get_conn():
    if 'db' not in g:
        url = DATABASE_URL
        if url.startswith('postgres://'):
            url = url.replace('postgres://', 'postgresql://', 1)
        conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
        conn.autocommit = True
        g.db = conn
    return g.db


@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db and not db.closed:
        db.close()


def _q(sql, params=(), one=False):
    with _get_conn().cursor() as cur:
        cur.execute(sql, params)
        return cur.fetchone() if one else cur.fetchall()


def _run(sql, params=()):
    with _get_conn().cursor() as cur:
        cur.execute(sql, params)


def _ins(sql, params=()):
    with _get_conn().cursor() as cur:
        cur.execute(sql + ' RETURNING id', params)
        return cur.fetchone()['id']


def _log(acao, entidade, entidade_id=None, descricao=''):
    try:
        uid = current_user.id if current_user.is_authenticated else None
        unome = (current_user.nome or current_user.username) if current_user.is_authenticated else 'Sistema'
        _run(
            "INSERT INTO log_auditoria (usuario_id, usuario_nome, acao, entidade, entidade_id, descricao)"
            " VALUES (%s, %s, %s, %s, %s, %s)",
            (uid, unome, acao, entidade, entidade_id, descricao)
        )
        _run("DELETE FROM log_auditoria WHERE created_at < NOW() - INTERVAL '365 days'")
    except Exception:
        pass


def init_db():
    url = DATABASE_URL
    if url.startswith('postgres://'):
        url = url.replace('postgres://', 'postgresql://', 1)
    conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
    conn.autocommit = True
    with conn.cursor() as cur:
        cur.execute('''
        CREATE TABLE IF NOT EXISTS usuario (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            nome TEXT,
            role TEXT DEFAULT 'operador'
        );
        CREATE TABLE IF NOT EXISTS estagiario (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            cpf TEXT UNIQUE NOT NULL,
            rg TEXT,
            data_nascimento TEXT,
            telefone TEXT,
            email TEXT,
            endereco TEXT,
            banco TEXT,
            agencia TEXT,
            conta TEXT,
            obs TEXT
        );
        CREATE TABLE IF NOT EXISTS empresa (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            cnpj TEXT,
            endereco TEXT,
            cidade TEXT DEFAULT \'Vitória da Conquista\',
            telefone TEXT,
            email TEXT,
            ramo TEXT,
            representante TEXT,
            cargo_representante TEXT,
            supervisor_nome TEXT,
            supervisor_cargo TEXT,
            supervisor_registro TEXT
        );
        CREATE TABLE IF NOT EXISTS ie (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            sigla TEXT,
            endereco TEXT,
            cidade TEXT DEFAULT \'Vitória da Conquista\',
            telefone TEXT,
            email TEXT,
            coordenador TEXT,
            coordenador_cargo TEXT
        );
        CREATE TABLE IF NOT EXISTS contrato (
            id SERIAL PRIMARY KEY,
            estagiario_id INTEGER NOT NULL REFERENCES estagiario(id),
            empresa_id INTEGER NOT NULL REFERENCES empresa(id),
            ie_id INTEGER NOT NULL REFERENCES ie(id),
            orientador TEXT DEFAULT \'Salmo Lima Costa\',
            supervisor_nome TEXT,
            supervisor_cargo TEXT,
            supervisor_registro TEXT,
            curso TEXT NOT NULL,
            tipo_estagio TEXT DEFAULT \'Não Obrigatório\',
            area_atuacao TEXT,
            ch_diaria INTEGER DEFAULT 6,
            ch_semanal INTEGER DEFAULT 30,
            data_inicio TEXT NOT NULL,
            data_fim TEXT NOT NULL,
            numero_contrato TEXT,
            bolsa REAL,
            taxa REAL,
            aux_transporte REAL,
            atividades TEXT,
            obs TEXT,
            num_relatorio INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        ''')
        cur.execute("ALTER TABLE contrato ADD COLUMN IF NOT EXISTS jornada TEXT;")
        cur.execute("ALTER TABLE contrato ADD COLUMN IF NOT EXISTS data_encerramento TEXT;")
        cur.execute("ALTER TABLE estagiario ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'ativo';")
        cur.execute("UPDATE estagiario SET status = 'ativo' WHERE status IS NULL;")
        cur.execute("ALTER TABLE empresa ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'ativo';")
        cur.execute("UPDATE empresa SET status = 'ativo' WHERE status IS NULL;")
        cur.execute("ALTER TABLE empresa ADD COLUMN IF NOT EXISTS cpf_representante TEXT;")
        cur.execute("ALTER TABLE estagiario ADD COLUMN IF NOT EXISTS semestre INTEGER;")
        cur.execute("ALTER TABLE estagiario ADD COLUMN IF NOT EXISTS tipo_ensino TEXT DEFAULT 'superior';")
        cur.execute("ALTER TABLE estagiario ADD COLUMN IF NOT EXISTS matricula TEXT;")
        cur.execute("ALTER TABLE ie ADD COLUMN IF NOT EXISTS representante_legal TEXT;")
        cur.execute("ALTER TABLE ie ADD COLUMN IF NOT EXISTS cargo_representante_legal TEXT;")
        cur.execute("ALTER TABLE ie ADD COLUMN IF NOT EXISTS signatario_tce TEXT DEFAULT 'coordenador';")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS ie_professor (
            id SERIAL PRIMARY KEY,
            ie_id INTEGER NOT NULL REFERENCES ie(id) ON DELETE CASCADE,
            nome TEXT NOT NULL,
            cargo TEXT,
            ordem INTEGER DEFAULT 0
        )""")
        cur.execute("ALTER TABLE contrato ADD COLUMN IF NOT EXISTS ie_professor_id INTEGER;")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS aditivo (
            id SERIAL PRIMARY KEY,
            contrato_id INTEGER NOT NULL REFERENCES contrato(id) ON DELETE CASCADE,
            nova_data_fim TEXT,
            clausulas TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS config (
            chave TEXT PRIMARY KEY,
            valor TEXT
        )""")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS log_auditoria (
            id SERIAL PRIMARY KEY,
            created_at TIMESTAMP DEFAULT NOW(),
            usuario_id INTEGER,
            usuario_nome TEXT,
            acao TEXT NOT NULL,
            entidade TEXT NOT NULL,
            entidade_id INTEGER,
            descricao TEXT
        )""")
        for chave in ['seg_seguradora', 'seg_apolice', 'seg_coberturas', 'seg_vigencia']:
            cur.execute(
                "INSERT INTO config (chave, valor) VALUES (%s, '') ON CONFLICT (chave) DO NOTHING",
                (chave,))
        cur.execute("SELECT id FROM usuario WHERE username = 'salmo'")
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO usuario (username, password_hash, nome, role) VALUES (%s, %s, %s, %s)",
                ('salmo', generate_password_hash('ciclorh2026'), 'Salmo Lima Costa', 'admin')
            )
    conn.close()


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def fmt_date(d):
    if not d:
        return ''
    if isinstance(d, str):
        try:
            d = datetime.strptime(d[:10], '%Y-%m-%d').date()
        except Exception:
            return d
    return d.strftime('%d/%m/%Y')


def calcular_status(data_fim_str):
    if not data_fim_str:
        return 'SEM DATA'
    try:
        fim = datetime.strptime(str(data_fim_str)[:10], '%Y-%m-%d').date()
        diff = (fim - date.today()).days
        if diff < 0:
            return 'VENCIDO'
        if diff <= 30:
            return f'VENCE EM {diff} DIAS'
        return 'ATIVO'
    except Exception:
        return 'ATIVO'


def br_currency(valor):
    try:
        v = float(valor or 0)
    except (TypeError, ValueError):
        return '0,00'
    s = '{:,.2f}'.format(v)
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def num_extenso(valor):
    try:
        valor = round(float(valor), 2)
    except (TypeError, ValueError):
        return ''
    inteiro = int(valor)
    centavos = round((valor - inteiro) * 100)
    _un = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove',
           'dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove']
    _dez = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    _cent = ['', 'cem', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos',
             'seiscentos', 'setecentos', 'oitocentos', 'novecentos']

    def dez(n):
        if n < 20:
            return _un[n]
        d, u = _dez[n // 10], _un[n % 10]
        return d + (' e ' + u if u else '')

    def cent(n):
        c, r = n // 100, n % 100
        if c == 1 and r == 0:
            return 'cem'
        base = 'cento' if c == 1 else _cent[c]
        return base + (' e ' + dez(r) if r else '')

    def grupo(n):
        return cent(n) if n >= 100 else dez(n)

    def por_extenso(n):
        if n == 0:
            return 'zero'
        partes = []
        if n >= 1000000:
            m = n // 1000000
            partes.append(grupo(m) + (' milhão' if m == 1 else ' milhões'))
            n %= 1000000
        if n >= 1000:
            m = n // 1000
            partes.append(('um' if m == 1 else grupo(m)) + ' mil')
            n %= 1000
        if n > 0:
            partes.append(grupo(n))
        return ' e '.join(partes)

    partes = []
    if inteiro > 0:
        partes.append(por_extenso(inteiro) + (' real' if inteiro == 1 else ' reais'))
    if centavos > 0:
        partes.append(dez(centavos) + (' centavo' if centavos == 1 else ' centavos'))
    return ' e '.join(partes) if partes else 'zero reais'


def formatar_jornada(jornada_json):
    if not jornada_json:
        return ''
    try:
        j = json.loads(jornada_json)
    except Exception:
        return ''

    DIAS_KEYS = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab']
    DIAS_ACUS = {
        'dom': 'aos domingos', 'seg': 'às segundas-feiras', 'ter': 'às terças-feiras',
        'qua': 'às quartas-feiras', 'qui': 'às quintas-feiras', 'sex': 'às sextas-feiras',
        'sab': 'aos sábados',
    }
    DIAS_DE = {
        'dom': 'domingo', 'seg': 'segunda', 'ter': 'terça', 'qua': 'quarta',
        'qui': 'quinta', 'sex': 'sexta', 'sab': 'sábado',
    }
    DIAS_ATE = {
        'dom': 'domingo', 'seg': 'segunda-feira', 'ter': 'terça-feira', 'qua': 'quarta-feira',
        'qui': 'quinta-feira', 'sex': 'sexta-feira', 'sab': 'sábado',
    }

    def dia_horarios(d):
        h = []
        for periodo in ['mat', 'ves', 'not']:
            ini = d.get(f'{periodo}_ini', '').strip()
            fim = d.get(f'{periodo}_fim', '').strip()
            if ini and fim:
                h.append((ini, fim))
        return tuple(h)

    schedule_groups = {}
    for k in DIAS_KEYS:
        if k not in j:
            continue
        h = dia_horarios(j[k])
        if not h:
            continue
        if h not in schedule_groups:
            schedule_groups[h] = []
        schedule_groups[h].append(k)

    if not schedule_groups:
        return ''

    parts = []
    for horarios, dias in schedule_groups.items():
        dias_sorted = sorted(dias, key=lambda d: DIAS_KEYS.index(d))
        n = len(dias_sorted)
        indices = [DIAS_KEYS.index(d) for d in dias_sorted]
        is_consecutive = (n > 1 and indices[-1] - indices[0] == n - 1)

        if is_consecutive and n >= 2:
            dia_fmt = f'de {DIAS_DE[dias_sorted[0]]} a {DIAS_ATE[dias_sorted[-1]]}'
        elif n == 1:
            dia_fmt = DIAS_ACUS[dias_sorted[0]]
        elif n == 2:
            dia_fmt = f'{DIAS_ACUS[dias_sorted[0]]} e {DIAS_ACUS[dias_sorted[1]]}'
        else:
            dia_fmt = ', '.join(DIAS_ACUS[d] for d in dias_sorted[:-1]) + f' e {DIAS_ACUS[dias_sorted[-1]]}'

        horario_fmt = ' e '.join(f'das {ini} às {fim}' for ini, fim in horarios)
        parts.append(f'{dia_fmt} {horario_fmt}')

    if not parts:
        return ''
    result = '; '.join(parts)
    return result[0].upper() + result[1:] + '.'


def calcular_ch_jornada(jornada_json):
    if not jornada_json:
        return None, None
    try:
        j = json.loads(jornada_json)
    except Exception:
        return None, None
    dias_keys = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab']
    total_semanal = 0.0
    dias_horas = []
    for dia in dias_keys:
        if dia not in j:
            continue
        dd = j[dia]
        horas_dia = 0.0
        for periodo in ['mat', 'ves', 'not']:
            ini = dd.get(f'{periodo}_ini', '').strip()
            fim = dd.get(f'{periodo}_fim', '').strip()
            if ini and fim:
                try:
                    hi = int(ini[:2]) + int(ini[3:5]) / 60
                    hf = int(fim[:2]) + int(fim[3:5]) / 60
                    if hf > hi:
                        horas_dia += hf - hi
                except Exception:
                    pass
        if horas_dia > 0:
            dias_horas.append(horas_dia)
            total_semanal += horas_dia
    if not dias_horas:
        return None, None
    return round(max(dias_horas), 1), round(total_semanal, 1)


def _build_jornada_json():
    dias = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab']
    jornada = {}
    for dia in dias:
        dd = {}
        for periodo in ['mat', 'ves', 'not']:
            ini = request.form.get(f'{dia}_{periodo}_ini', '').strip()
            fim = request.form.get(f'{dia}_{periodo}_fim', '').strip()
            if ini or fim:
                dd[f'{periodo}_ini'] = ini
                dd[f'{periodo}_fim'] = fim
        if dd:
            jornada[dia] = dd
    return json.dumps(jornada, ensure_ascii=False) if jornada else None


def _get_config():
    try:
        rows = _q("SELECT chave, valor FROM config")
        return {r['chave']: r['valor'] or '' for r in rows}
    except Exception:
        return {}


def _fmt_semestre(semestre, tipo_ensino):
    if not semestre:
        return ''
    if tipo_ensino == 'medio':
        return f'{semestre}º Ano'
    return f'{semestre}º Semestre'


_semestre_atualizado = set()


def _atualizar_semestres_auto():
    hoje = date.today()
    if hoje.month not in (1, 7):
        return
    key = (hoje.year, hoje.month)
    if key in _semestre_atualizado:
        return
    _semestre_atualizado.add(key)
    chave_cfg = f'sem_upd_{hoje.year}_{hoje.month:02d}'
    try:
        cfg = _get_config()
        if cfg.get(chave_cfg):
            return
        if hoje.month == 7:
            _run("""UPDATE estagiario SET semestre = LEAST(semestre + 1, 10)
                    WHERE tipo_ensino IN ('superior', 'tecnico')
                    AND semestre IS NOT NULL AND semestre < 10""")
        else:
            _run("""UPDATE estagiario SET semestre = CASE
                        WHEN tipo_ensino = 'medio' THEN LEAST(semestre + 1, 3)
                        ELSE LEAST(semestre + 1, 10)
                    END WHERE semestre IS NOT NULL""")
        _run("INSERT INTO config (chave, valor) VALUES (%s, %s) ON CONFLICT (chave) DO UPDATE SET valor = EXCLUDED.valor",
             (chave_cfg, hoje.isoformat()))
    except Exception:
        pass


@app.before_request
def _antes_da_requisicao():
    _atualizar_semestres_auto()


app.jinja_env.globals.update(fmt_date=fmt_date, calcular_status=calcular_status, fmt_semestre=_fmt_semestre)
app.jinja_env.filters['from_json'] = lambda s: json.loads(s) if s else {}


@app.context_processor
def inject_pending():
    pending = 0
    if current_user.is_authenticated and current_user.is_admin:
        try:
            r1 = _q("SELECT COUNT(*) n FROM estagiario WHERE status='pendente'", one=True)
            r2 = _q("SELECT COUNT(*) n FROM empresa WHERE status='pendente'", one=True)
            pending = (r1['n'] if r1 else 0) + (r2['n'] if r2 else 0)
        except Exception:
            pending = 0
    return {'pending_count': pending}


# ─── AUTH ─────────────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        senha = request.form.get('senha', '')
        row = _q("SELECT * FROM usuario WHERE username = %s", (username,), one=True)
        if row and check_password_hash(row['password_hash'], senha):
            user = User(row['id'], row['username'], row['nome'], row['role'])
            login_user(user, remember=True)
            _log('login', 'sistema', None, f'Login: {username}')
            return redirect(request.args.get('next') or url_for('index'))
        flash('Usuário ou senha incorretos.', 'danger')
    return render_template('auth/login.html')


@app.route('/logout')
@login_required
def logout():
    _log('logout', 'sistema', None, f'Logout: {current_user.username}')
    logout_user()
    return redirect(url_for('login'))


# ─── ADMIN — USUÁRIOS ─────────────────────────────────────────────────────────

@app.route('/admin/usuarios')
@admin_required
def admin_usuarios():
    rows = _q("SELECT id, username, nome, role FROM usuario ORDER BY role DESC, nome")
    return render_template('admin/usuarios.html', usuarios=rows)


@app.route('/admin/usuarios/novo', methods=['GET', 'POST'])
@admin_required
def admin_usuario_novo():
    if request.method == 'POST':
        username = request.form['username'].strip().lower()
        senha = request.form['senha']
        nome = request.form.get('nome', '').strip()
        role = request.form.get('role', 'operador')
        if len(senha) < 6:
            flash('Senha deve ter no mínimo 6 caracteres.', 'danger')
        else:
            try:
                _ins("INSERT INTO usuario (username, password_hash, nome, role) VALUES (%s, %s, %s, %s)",
                     (username, generate_password_hash(senha), nome, role))
                _log('criar', 'usuario', None, f'Criou usuário: {username} ({role})')
                flash(f'Usuário "{username}" criado!', 'success')
                return redirect(url_for('admin_usuarios'))
            except psycopg2.errors.UniqueViolation:
                flash('Nome de usuário já existe.', 'danger')
    return render_template('admin/usuario_form.html', u=None)


@app.route('/admin/usuarios/<int:id>/editar', methods=['GET', 'POST'])
@admin_required
def admin_usuario_editar(id):
    u = _q("SELECT * FROM usuario WHERE id = %s", (id,), one=True)
    if not u:
        abort(404)
    if request.method == 'POST':
        nome = request.form.get('nome', '').strip()
        role = request.form.get('role', 'operador')
        _run("UPDATE usuario SET nome = %s, role = %s WHERE id = %s", (nome, role, id))
        senha = request.form.get('senha', '').strip()
        if senha:
            if len(senha) < 6:
                flash('Senha deve ter no mínimo 6 caracteres.', 'danger')
                return render_template('admin/usuario_form.html', u=u)
            _run("UPDATE usuario SET password_hash = %s WHERE id = %s",
                 (generate_password_hash(senha), id))
        _log('editar', 'usuario', id, f'Editou usuário ID {id} ({u["username"]})')
        flash('Usuário atualizado!', 'success')
        return redirect(url_for('admin_usuarios'))
    return render_template('admin/usuario_form.html', u=u)


@app.route('/admin/usuarios/<int:id>/excluir')
@admin_required
def admin_usuario_excluir(id):
    if str(id) == current_user.id:
        flash('Você não pode excluir sua própria conta.', 'danger')
        return redirect(url_for('admin_usuarios'))
    reg = _q("SELECT username FROM usuario WHERE id = %s", (id,), one=True)
    _run("DELETE FROM usuario WHERE id = %s", (id,))
    _log('excluir', 'usuario', id, f'Excluiu usuário: {reg["username"] if reg else id}')
    flash('Usuário excluído.', 'warning')
    return redirect(url_for('admin_usuarios'))


# ─── ADMIN — IMPORTAR PLANILHA ────────────────────────────────────────────────

@app.route('/admin/importar', methods=['GET', 'POST'])
@admin_required
def admin_importar():
    resultado = None
    if request.method == 'POST':
        arquivo = request.files.get('arquivo')
        if not arquivo or not arquivo.filename.endswith(('.xlsx', '.xlsm')):
            flash('Envie um arquivo .xlsx ou .xlsm válido.', 'danger')
            return render_template('admin/importar.html', resultado=None)
        import tempfile, importar as imp_mod
        with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as tmp:
            arquivo.save(tmp.name)
            resultado = imp_mod.run_from_file(tmp.name, DATABASE_URL)
        os.unlink(tmp.name)
        flash(f'Importação concluída: {resultado["ok"]} contratos importados.', 'success')
    return render_template('admin/importar.html', resultado=resultado)


# ─── DASHBOARD ────────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    total = _q("SELECT COUNT(*) AS n FROM contrato", one=True)['n']
    d30 = (date.today() + timedelta(days=30)).isoformat()
    vencendo = _q("""
        SELECT c.*, e.nome est_nome, emp.nome emp_nome
        FROM contrato c
        JOIN estagiario e ON e.id = c.estagiario_id
        JOIN empresa emp ON emp.id = c.empresa_id
        WHERE c.data_fim <= %s ORDER BY c.data_fim
    """, (d30,))
    recentes = _q("""
        SELECT c.*, e.nome est_nome, emp.nome emp_nome
        FROM contrato c
        JOIN estagiario e ON e.id = c.estagiario_id
        JOIN empresa emp ON emp.id = c.empresa_id
        ORDER BY c.created_at DESC LIMIT 10
    """)
    total_est = _q("SELECT COUNT(*) AS n FROM estagiario", one=True)['n']
    total_emp = _q("SELECT COUNT(*) AS n FROM empresa", one=True)['n']
    total_ie = _q("SELECT COUNT(*) AS n FROM ie", one=True)['n']
    return render_template('index.html', total=total, vencendo=vencendo, recentes=recentes,
                           total_est=total_est, total_emp=total_emp, total_ie=total_ie)


# ─── ESTAGIÁRIOS ──────────────────────────────────────────────────────────────

@app.route('/estagiarios')
@login_required
def estagiarios():
    q = request.args.get('q', '')
    if q:
        rows = _q("""SELECT e.*,
                     (SELECT COUNT(*) FROM contrato WHERE estagiario_id = e.id) qtd_contratos
                     FROM estagiario e
                     WHERE (e.nome ILIKE %s OR e.cpf ILIKE %s) AND e.status = 'ativo' ORDER BY e.nome""",
                  (f'%{q}%', f'%{q}%'))
    else:
        rows = _q("""SELECT e.*,
                     (SELECT COUNT(*) FROM contrato WHERE estagiario_id = e.id) qtd_contratos
                     FROM estagiario e WHERE e.status = 'ativo' ORDER BY e.nome""")
    return render_template('estagiarios/lista.html', estagiarios=rows, q=q)


@app.route('/estagiarios/novo', methods=['GET', 'POST'])
@login_required
def estagiario_novo():
    if request.method == 'POST':
        try:
            _ins("""INSERT INTO estagiario
                    (nome,cpf,rg,data_nascimento,telefone,email,endereco,banco,agencia,conta,obs,
                     tipo_ensino,semestre,matricula)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                 (request.form['nome'], request.form['cpf'],
                  request.form.get('rg'), request.form.get('data_nascimento') or None,
                  request.form.get('telefone'), request.form.get('email'),
                  request.form.get('endereco'), request.form.get('banco'),
                  request.form.get('agencia'), request.form.get('conta'),
                  request.form.get('obs'),
                  request.form.get('tipo_ensino', 'superior'),
                  request.form.get('semestre') or None,
                  request.form.get('matricula') or None))
            _log('criar', 'estagiario', None, f'Criou estagiário: {request.form["nome"]} (CPF: {request.form["cpf"]})')
            flash('Estagiário cadastrado!', 'success')
            return redirect(url_for('estagiarios'))
        except psycopg2.errors.UniqueViolation:
            flash('CPF já cadastrado!', 'danger')
    return render_template('estagiarios/form.html', e=None)


@app.route('/estagiarios/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def estagiario_editar(id):
    e = _q("SELECT * FROM estagiario WHERE id = %s", (id,), one=True)
    if not e:
        abort(404)
    if request.method == 'POST':
        _run("""UPDATE estagiario SET
                nome=%s,cpf=%s,rg=%s,data_nascimento=%s,telefone=%s,email=%s,
                endereco=%s,banco=%s,agencia=%s,conta=%s,obs=%s,
                tipo_ensino=%s,semestre=%s,matricula=%s WHERE id=%s""",
             (request.form['nome'], request.form['cpf'],
              request.form.get('rg'), request.form.get('data_nascimento') or None,
              request.form.get('telefone'), request.form.get('email'),
              request.form.get('endereco'), request.form.get('banco'),
              request.form.get('agencia'), request.form.get('conta'),
              request.form.get('obs'),
              request.form.get('tipo_ensino', 'superior'),
              request.form.get('semestre') or None,
              request.form.get('matricula') or None,
              id))
        _log('editar', 'estagiario', id, f'Editou estagiário: {request.form["nome"]}')
        flash('Atualizado!', 'success')
        return redirect(url_for('estagiarios'))
    return render_template('estagiarios/form.html', e=e)


@app.route('/estagiarios/<int:id>/excluir')
@login_required
def estagiario_excluir(id):
    reg = _q("SELECT nome FROM estagiario WHERE id = %s", (id,), one=True)
    _run("DELETE FROM estagiario WHERE id = %s", (id,))
    _log('excluir', 'estagiario', id, f'Excluiu estagiário: {reg["nome"] if reg else id}')
    flash('Excluído.', 'warning')
    return redirect(url_for('estagiarios'))


# ─── EMPRESAS ─────────────────────────────────────────────────────────────────

@app.route('/empresas')
@login_required
def empresas():
    q = request.args.get('q', '')
    if q:
        rows = _q("""SELECT emp.*,
                     (SELECT COUNT(*) FROM contrato WHERE empresa_id = emp.id) qtd_contratos
                     FROM empresa emp
                     WHERE (emp.nome ILIKE %s OR emp.cnpj ILIKE %s) AND emp.status = 'ativo' ORDER BY emp.nome""",
                  (f'%{q}%', f'%{q}%'))
    else:
        rows = _q("""SELECT emp.*,
                     (SELECT COUNT(*) FROM contrato WHERE empresa_id = emp.id) qtd_contratos
                     FROM empresa emp WHERE emp.status = 'ativo' ORDER BY emp.nome""")
    return render_template('empresas/lista.html', empresas=rows, q=q)


@app.route('/empresas/nova', methods=['GET', 'POST'])
@login_required
def empresa_nova():
    if request.method == 'POST':
        _ins("""INSERT INTO empresa
                (nome,cnpj,endereco,cidade,telefone,email,ramo,
                 representante,cargo_representante,cpf_representante,supervisor_nome,supervisor_cargo,supervisor_registro)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
             (request.form['nome'], request.form.get('cnpj'),
              request.form.get('endereco'), request.form.get('cidade', 'Vitória da Conquista'),
              request.form.get('telefone'), request.form.get('email'),
              request.form.get('ramo'), request.form.get('representante'),
              request.form.get('cargo_representante'), request.form.get('cpf_representante'),
              request.form.get('supervisor_nome'),
              request.form.get('supervisor_cargo'), request.form.get('supervisor_registro')))
        _log('criar', 'empresa', None, f'Criou empresa: {request.form["nome"]}')
        flash('Empresa cadastrada!', 'success')
        return redirect(url_for('empresas'))
    return render_template('empresas/form.html', emp=None)


@app.route('/empresas/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def empresa_editar(id):
    emp = _q("SELECT * FROM empresa WHERE id = %s", (id,), one=True)
    if not emp:
        abort(404)
    if request.method == 'POST':
        _run("""UPDATE empresa SET
                nome=%s,cnpj=%s,endereco=%s,cidade=%s,telefone=%s,email=%s,ramo=%s,
                representante=%s,cargo_representante=%s,cpf_representante=%s,supervisor_nome=%s,
                supervisor_cargo=%s,supervisor_registro=%s WHERE id=%s""",
             (request.form['nome'], request.form.get('cnpj'),
              request.form.get('endereco'), request.form.get('cidade'),
              request.form.get('telefone'), request.form.get('email'),
              request.form.get('ramo'), request.form.get('representante'),
              request.form.get('cargo_representante'), request.form.get('cpf_representante'),
              request.form.get('supervisor_nome'),
              request.form.get('supervisor_cargo'), request.form.get('supervisor_registro'), id))
        _log('editar', 'empresa', id, f'Editou empresa: {request.form["nome"]}')
        flash('Atualizada!', 'success')
        return redirect(url_for('empresas'))
    return render_template('empresas/form.html', emp=emp)


@app.route('/empresas/<int:id>/excluir')
@login_required
def empresa_excluir(id):
    reg = _q("SELECT nome FROM empresa WHERE id = %s", (id,), one=True)
    _run("DELETE FROM empresa WHERE id = %s", (id,))
    _log('excluir', 'empresa', id, f'Excluiu empresa: {reg["nome"] if reg else id}')
    flash('Excluída.', 'warning')
    return redirect(url_for('empresas'))


@app.route('/api/ie_professores/<int:ie_id>')
@login_required
def api_ie_professores(ie_id):
    profs = _q("SELECT id, nome, cargo FROM ie_professor WHERE ie_id = %s ORDER BY ordem, id", (ie_id,))
    return jsonify([dict(p) for p in profs])


@app.route('/api/empresa/<int:id>')
@login_required
def api_empresa(id):
    row = _q("SELECT * FROM empresa WHERE id = %s", (id,), one=True)
    return jsonify(dict(row) if row else {})


# ─── IEs ──────────────────────────────────────────────────────────────────────

@app.route('/ies')
@login_required
def ies():
    rows = _q("""SELECT ie.*,
                 (SELECT COUNT(*) FROM contrato WHERE ie_id = ie.id) qtd_contratos
                 FROM ie ORDER BY ie.nome""")
    return render_template('ies/lista.html', ies=rows)


@app.route('/ies/nova', methods=['GET', 'POST'])
@login_required
def ie_nova():
    if request.method == 'POST':
        ie_id = _ins("""INSERT INTO ie
                (nome,sigla,endereco,cidade,telefone,email,coordenador,coordenador_cargo,
                 representante_legal,cargo_representante_legal,signatario_tce)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
             (request.form['nome'], request.form.get('sigla'),
              request.form.get('endereco'), request.form.get('cidade', 'Vitória da Conquista'),
              request.form.get('telefone'), request.form.get('email'),
              request.form.get('coordenador'), request.form.get('coordenador_cargo'),
              request.form.get('representante_legal'), request.form.get('cargo_representante_legal'),
              request.form.get('signatario_tce', 'coordenador')))
        nomes_p = request.form.getlist('prof_nome[]')
        cargos_p = request.form.getlist('prof_cargo[]')
        for i, np in enumerate(nomes_p):
            cp = cargos_p[i] if i < len(cargos_p) else ''
            if np.strip():
                _ins("INSERT INTO ie_professor (ie_id,nome,cargo,ordem) VALUES (%s,%s,%s,%s)",
                     (ie_id, np.strip(), cp.strip(), i))
        _log('criar', 'ie', ie_id, f'Criou instituição: {request.form["nome"]}')
        flash('Instituição cadastrada!', 'success')
        return redirect(url_for('ies'))
    return render_template('ies/form.html', ie=None, professores=[])


@app.route('/ies/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def ie_editar(id):
    ie = _q("SELECT * FROM ie WHERE id = %s", (id,), one=True)
    if not ie:
        abort(404)
    if request.method == 'POST':
        _run("""UPDATE ie SET nome=%s,sigla=%s,endereco=%s,cidade=%s,telefone=%s,email=%s,
                coordenador=%s,coordenador_cargo=%s,representante_legal=%s,
                cargo_representante_legal=%s,signatario_tce=%s WHERE id=%s""",
             (request.form['nome'], request.form.get('sigla'),
              request.form.get('endereco'), request.form.get('cidade'),
              request.form.get('telefone'), request.form.get('email'),
              request.form.get('coordenador'), request.form.get('coordenador_cargo'),
              request.form.get('representante_legal'), request.form.get('cargo_representante_legal'),
              request.form.get('signatario_tce', 'coordenador'), id))
        _run("DELETE FROM ie_professor WHERE ie_id = %s", (id,))
        nomes_p = request.form.getlist('prof_nome[]')
        cargos_p = request.form.getlist('prof_cargo[]')
        for i, np in enumerate(nomes_p):
            cp = cargos_p[i] if i < len(cargos_p) else ''
            if np.strip():
                _ins("INSERT INTO ie_professor (ie_id,nome,cargo,ordem) VALUES (%s,%s,%s,%s)",
                     (id, np.strip(), cp.strip(), i))
        _log('editar', 'ie', id, f'Editou instituição: {request.form["nome"]}')
        flash('Atualizada!', 'success')
        return redirect(url_for('ies'))
    professores = _q("SELECT * FROM ie_professor WHERE ie_id = %s ORDER BY ordem, id", (id,))
    return render_template('ies/form.html', ie=ie, professores=professores)


@app.route('/ies/<int:id>/excluir')
@login_required
def ie_excluir(id):
    reg = _q("SELECT nome FROM ie WHERE id = %s", (id,), one=True)
    _run("DELETE FROM ie WHERE id = %s", (id,))
    _log('excluir', 'ie', id, f'Excluiu instituição: {reg["nome"] if reg else id}')
    flash('Excluída.', 'warning')
    return redirect(url_for('ies'))


# ─── CONTRATOS ────────────────────────────────────────────────────────────────

@app.route('/contratos')
@login_required
def contratos():
    q = request.args.get('q', '')
    status = request.args.get('status', '')
    sql = """SELECT c.*, e.nome est_nome, emp.nome emp_nome,
             ie.sigla ie_sigla, ie.nome ie_nome,
             COALESCE(
                 c.data_encerramento,
                 (SELECT nova_data_fim FROM aditivo
                  WHERE contrato_id = c.id AND nova_data_fim IS NOT NULL AND nova_data_fim != ''
                  ORDER BY created_at DESC LIMIT 1),
                 c.data_fim
             ) as effective_data_fim
             FROM contrato c
             JOIN estagiario e ON e.id = c.estagiario_id
             JOIN empresa emp ON emp.id = c.empresa_id
             JOIN ie ON ie.id = c.ie_id"""
    params = []
    if q:
        sql += " WHERE (e.nome ILIKE %s OR emp.nome ILIKE %s)"
        params += [f'%{q}%', f'%{q}%']
    sql += ' ORDER BY effective_data_fim'
    rows = _q(sql, params)
    if status:
        filtered = []
        for r in rows:
            st = calcular_status(r['effective_data_fim'])
            if status == 'ativo' and st == 'ATIVO':
                filtered.append(r)
            elif status == 'vencendo' and 'DIAS' in st:
                filtered.append(r)
            elif status == 'vencido' and st == 'VENCIDO':
                filtered.append(r)
            elif status == 'encerrado' and r.get('data_encerramento'):
                filtered.append(r)
        rows = filtered
    return render_template('contratos/lista.html', contratos=rows, q=q, status=status)


@app.route('/contratos/novo', methods=['GET', 'POST'])
@login_required
def contrato_novo():
    if request.method == 'POST':
        ats = '||'.join(request.form.get(f'atividade_{i}', '') for i in range(1, 10))
        _ins("""INSERT INTO contrato
                (estagiario_id,empresa_id,ie_id,orientador,
                 supervisor_nome,supervisor_cargo,supervisor_registro,
                 curso,tipo_estagio,area_atuacao,ch_diaria,ch_semanal,
                 data_inicio,data_fim,numero_contrato,bolsa,taxa,aux_transporte,atividades,obs,
                 jornada,data_encerramento,ie_professor_id)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
             (request.form['estagiario_id'], request.form['empresa_id'], request.form['ie_id'],
              request.form.get('orientador', 'Salmo Lima Costa'),
              request.form.get('supervisor_nome'), request.form.get('supervisor_cargo'),
              request.form.get('supervisor_registro'),
              request.form['curso'], request.form.get('tipo_estagio', 'Não Obrigatório'),
              request.form.get('area_atuacao'),
              request.form.get('ch_diaria', 6), request.form.get('ch_semanal', 30),
              request.form['data_inicio'], request.form['data_fim'],
              request.form.get('numero_contrato'),
              request.form.get('bolsa') or None,
              request.form.get('taxa') or None,
              request.form.get('aux_transporte') or None,
              ats, request.form.get('obs'), _build_jornada_json(),
              request.form.get('data_encerramento') or None,
              request.form.get('ie_professor_id') or None))
        _est = _q("SELECT nome FROM estagiario WHERE id=%s", (request.form['estagiario_id'],), one=True)
        _emp = _q("SELECT nome FROM empresa WHERE id=%s", (request.form['empresa_id'],), one=True)
        _log('criar', 'contrato', None,
             f'Criou contrato: {_est["nome"] if _est else "?"} @ {_emp["nome"] if _emp else "?"}'
             f' ({request.form["data_inicio"]} a {request.form["data_fim"]})')
        flash('Contrato criado!', 'success')
        return redirect(url_for('contratos'))
    estagiarios = _q("SELECT * FROM estagiario WHERE status='ativo' ORDER BY nome")
    empresas_list = _q("SELECT * FROM empresa WHERE status='ativo' ORDER BY nome")
    ies_list = _q("SELECT * FROM ie ORDER BY nome")
    return render_template('contratos/form.html', c=None,
                           estagiarios=estagiarios, empresas=empresas_list, ies=ies_list,
                           aditivos=[])


@app.route('/contratos/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def contrato_editar(id):
    c = _q("SELECT * FROM contrato WHERE id = %s", (id,), one=True)
    if not c:
        abort(404)
    if request.method == 'POST':
        ats = '||'.join(request.form.get(f'atividade_{i}', '') for i in range(1, 10))
        _run("""UPDATE contrato SET
                estagiario_id=%s,empresa_id=%s,ie_id=%s,orientador=%s,
                supervisor_nome=%s,supervisor_cargo=%s,supervisor_registro=%s,
                curso=%s,tipo_estagio=%s,area_atuacao=%s,ch_diaria=%s,ch_semanal=%s,
                data_inicio=%s,data_fim=%s,numero_contrato=%s,bolsa=%s,taxa=%s,
                aux_transporte=%s,atividades=%s,obs=%s,jornada=%s,data_encerramento=%s,
                ie_professor_id=%s WHERE id=%s""",
             (request.form['estagiario_id'], request.form['empresa_id'], request.form['ie_id'],
              request.form.get('orientador'),
              request.form.get('supervisor_nome'), request.form.get('supervisor_cargo'),
              request.form.get('supervisor_registro'),
              request.form['curso'], request.form.get('tipo_estagio'),
              request.form.get('area_atuacao'),
              request.form.get('ch_diaria'), request.form.get('ch_semanal'),
              request.form['data_inicio'], request.form['data_fim'],
              request.form.get('numero_contrato'),
              request.form.get('bolsa') or None,
              request.form.get('taxa') or None,
              request.form.get('aux_transporte') or None,
              ats, request.form.get('obs'), _build_jornada_json(),
              request.form.get('data_encerramento') or None,
              request.form.get('ie_professor_id') or None, id))
        _est2 = _q("SELECT nome FROM estagiario WHERE id=%s", (request.form['estagiario_id'],), one=True)
        _emp2 = _q("SELECT nome FROM empresa WHERE id=%s", (request.form['empresa_id'],), one=True)
        _log('editar', 'contrato', id,
             f'Editou contrato: {_est2["nome"] if _est2 else "?"} @ {_emp2["nome"] if _emp2 else "?"}')
        flash('Contrato atualizado!', 'success')
        return redirect(url_for('contratos'))
    estagiarios = _q("SELECT * FROM estagiario WHERE status='ativo' ORDER BY nome")
    empresas_list = _q("SELECT * FROM empresa WHERE status='ativo' ORDER BY nome")
    ies_list = _q("SELECT * FROM ie ORDER BY nome")
    aditivos = _q("SELECT * FROM aditivo WHERE contrato_id = %s ORDER BY created_at", (id,))
    return render_template('contratos/form.html', c=c,
                           estagiarios=estagiarios, empresas=empresas_list, ies=ies_list,
                           aditivos=aditivos)


@app.route('/contratos/<int:id>/excluir')
@login_required
def contrato_excluir(id):
    info = _q("""SELECT e.nome est, emp.nome emp FROM contrato c
                 JOIN estagiario e ON e.id = c.estagiario_id
                 JOIN empresa emp ON emp.id = c.empresa_id WHERE c.id = %s""", (id,), one=True)
    _run("DELETE FROM contrato WHERE id = %s", (id,))
    desc = f'Excluiu contrato: {info["est"]} @ {info["emp"]}' if info else f'Excluiu contrato ID {id}'
    _log('excluir', 'contrato', id, desc)
    flash('Excluído.', 'warning')
    return redirect(url_for('contratos'))


# ─── DOCUMENTOS ───────────────────────────────────────────────────────────────

def _doc_ctx(id):
    c = _q("SELECT * FROM contrato WHERE id = %s", (id,), one=True)
    if not c:
        return None
    est = _q("SELECT * FROM estagiario WHERE id = %s", (c['estagiario_id'],), one=True)
    emp = _q("SELECT * FROM empresa WHERE id = %s", (c['empresa_id'],), one=True)
    ie = _q("SELECT * FROM ie WHERE id = %s", (c['ie_id'],), one=True)
    cfg = _get_config()
    aditivos = _q("SELECT * FROM aditivo WHERE contrato_id = %s ORDER BY created_at", (id,))

    try:
        ini = datetime.strptime(str(c['data_inicio'])[:10], '%Y-%m-%d').date()
        fim = datetime.strptime(str(c['data_fim'])[:10], '%Y-%m-%d').date()
        meses = (fim.year - ini.year) * 12 + fim.month - ini.month + 1
        ch_total = meses * 4 * (c['ch_semanal'] or 30)
    except Exception:
        ch_total = 0

    enc_str = c.get('data_encerramento')
    try:
        if enc_str:
            enc = datetime.strptime(str(enc_str)[:10], '%Y-%m-%d').date()
            meses_real = (enc.year - ini.year) * 12 + enc.month - ini.month + 1
            ch_total_real = meses_real * 4 * (c['ch_semanal'] or 30)
        else:
            ch_total_real = ch_total
    except Exception:
        ch_total_real = ch_total

    ch_j_diaria, ch_j_semanal = calcular_ch_jornada(c.get('jornada'))

    professor = None
    if c.get('ie_professor_id'):
        professor = _q("SELECT * FROM ie_professor WHERE id = %s", (c['ie_professor_id'],), one=True)

    sig_tce = (ie.get('signatario_tce') or 'coordenador') if ie else 'coordenador'
    if sig_tce == 'representante':
        sig_nome = ie.get('representante_legal', '') if ie else ''
        sig_cargo = ie.get('cargo_representante_legal', '') if ie else ''
        sig_tipo = 'Representante Legal'
    elif sig_tce == 'professor':
        sig_nome = professor['nome'] if professor else ''
        sig_cargo = professor['cargo'] if professor else ''
        sig_tipo = 'Professor Orientador'
    else:
        sig_nome = ie.get('coordenador', '') if ie else ''
        sig_cargo = ie.get('coordenador_cargo', '') if ie else ''
        sig_tipo = 'Coordenador(a) de Estágio'

    d = type('D', (), {
        'id': c['id'],
        'curso': c['curso'],
        'tipo_estagio': c['tipo_estagio'],
        'area_atuacao': c['area_atuacao'] or c['curso'],
        'ch_diaria': c['ch_diaria'] or 6,
        'ch_semanal': c['ch_semanal'] or 30,
        'data_inicio': c['data_inicio'],
        'data_fim': c['data_fim'],
        'data_encerramento': enc_str,
        'numero_contrato': c['numero_contrato'],
        'bolsa': c['bolsa'],
        'bolsa_extenso': num_extenso(c['bolsa']) if c['bolsa'] else '',
        'taxa': c['taxa'],
        'aux_transporte': c['aux_transporte'],
        'aux_transporte_extenso': num_extenso(c['aux_transporte']) if c['aux_transporte'] else '',
        'atividades': c['atividades'],
        'orientador': c['orientador'],
        'supervisor_nome': c['supervisor_nome'],
        'supervisor_cargo': c['supervisor_cargo'],
        'supervisor_registro': c['supervisor_registro'],
        'num_relatorio': c['num_relatorio'] or 1,
        'jornada_texto': formatar_jornada(c['jornada']),
        'ch_diaria_jornada': ch_j_diaria,
        'ch_semanal_jornada': ch_j_semanal,
        'seg_seguradora': cfg.get('seg_seguradora', ''),
        'seg_apolice': cfg.get('seg_apolice', ''),
        'seg_coberturas': cfg.get('seg_coberturas', ''),
        'seg_vigencia': cfg.get('seg_vigencia', ''),
        'est_nome': est['nome'] if est else '',
        'est_cpf': est['cpf'] if est else '',
        'est_rg': est['rg'] if est else '',
        'est_data_nasc': est['data_nascimento'] if est else '',
        'est_telefone': est['telefone'] if est else '',
        'est_email': est['email'] if est else '',
        'est_endereco': est['endereco'] if est else '',
        'est_semestre': _fmt_semestre(est['semestre'], est['tipo_ensino']) if est else '',
        'est_matricula': est['matricula'] if est else '',
        'emp_nome': emp['nome'] if emp else '',
        'emp_cnpj': emp['cnpj'] if emp else '',
        'emp_endereco': emp['endereco'] if emp else '',
        'emp_cidade': emp['cidade'] if emp else 'Vitória da Conquista',
        'emp_telefone': emp['telefone'] if emp else '',
        'emp_email': emp['email'] if emp else '',
        'emp_representante': emp['representante'] if emp else '',
        'emp_cargo_rep': emp['cargo_representante'] if emp else '',
        'emp_cpf_rep': emp.get('cpf_representante', '') if emp else '',
        'ie_nome': ie['nome'] if ie else '',
        'ie_sigla': ie['sigla'] if ie else '',
        'ie_endereco': ie['endereco'] if ie else '',
        'ie_cidade': ie['cidade'] if ie else 'Vitória da Conquista',
        'ie_coordenador': ie['coordenador'] if ie else '',
        'ie_coord_cargo': ie['coordenador_cargo'] if ie else '',
        'ie_representante_legal': ie.get('representante_legal', '') if ie else '',
        'ie_cargo_rep_legal': ie.get('cargo_representante_legal', '') if ie else '',
        'ie_professor_nome': professor['nome'] if professor else '',
        'ie_professor_cargo': professor['cargo'] if professor else '',
        'ie_signatario_nome': sig_nome,
        'ie_signatario_cargo': sig_cargo,
        'ie_signatario_tipo': sig_tipo,
    })()

    return dict(d=d, agente=AGENTE, ch_total=ch_total, ch_total_real=ch_total_real,
                data_hoje=fmt_date(date.today()), fmt_date=fmt_date,
                br_currency=br_currency, aditivos=aditivos)


@app.route('/contratos/<int:id>/tce')
@login_required
def doc_tce(id):
    ctx = _doc_ctx(id)
    if not ctx:
        flash('Contrato não encontrado.', 'danger')
        return redirect(url_for('contratos'))
    return render_template('docs/tce.html', **ctx)


@app.route('/contratos/<int:id>/plano')
@login_required
def doc_plano(id):
    ctx = _doc_ctx(id)
    if not ctx:
        flash('Contrato não encontrado.', 'danger')
        return redirect(url_for('contratos'))
    return render_template('docs/plano.html', **ctx)


@app.route('/contratos/<int:id>/ciencia')
@login_required
def doc_ciencia(id):
    ctx = _doc_ctx(id)
    if not ctx:
        flash('Contrato não encontrado.', 'danger')
        return redirect(url_for('contratos'))
    return render_template('docs/ciencia.html', **ctx)


@app.route('/contratos/<int:id>/tre')
@login_required
def doc_tre(id):
    ctx = _doc_ctx(id)
    if not ctx:
        flash('Contrato não encontrado.', 'danger')
        return redirect(url_for('contratos'))
    return render_template('docs/tre.html', **ctx)


@app.route('/contratos/<int:id>/relatorio')
@login_required
def doc_relatorio(id):
    ctx = _doc_ctx(id)
    if not ctx:
        flash('Contrato não encontrado.', 'danger')
        return redirect(url_for('contratos'))
    return render_template('docs/relatorio.html', **ctx)


@app.route('/contratos/<int:id>/aditivo')
@login_required
def doc_aditivo(id):
    ctx = _doc_ctx(id)
    if not ctx:
        flash('Contrato não encontrado.', 'danger')
        return redirect(url_for('contratos'))
    ctx['nova_data_fim'] = request.args.get('nova_data_fim', '')
    ctx['aditivo_num'] = request.args.get('aditivo_num', '')
    clausulas_extras = []
    i = 1
    while i <= 10:
        titulo = request.args.get(f'clausula_{i}_titulo', '').strip()
        texto = request.args.get(f'clausula_{i}_texto', '').strip()
        if not titulo and not texto:
            break
        clausulas_extras.append({'titulo': titulo, 'texto': texto})
        i += 1
    ctx['clausulas_extras'] = clausulas_extras
    return render_template('docs/aditivo.html', **ctx)


@app.route('/contratos/<int:id>/aditivo/registrar', methods=['POST'])
@login_required
def doc_aditivo_registrar(id):
    c = _q("SELECT id FROM contrato WHERE id = %s", (id,), one=True)
    if not c:
        return jsonify({'ok': False}), 404
    data = request.get_json(silent=True) or {}
    nova_data_fim = data.get('nova_data_fim') or None
    clausulas = data.get('clausulas', [])
    _ins("INSERT INTO aditivo (contrato_id, nova_data_fim, clausulas) VALUES (%s, %s, %s)",
         (id, nova_data_fim, json.dumps(clausulas, ensure_ascii=False)))
    numero = _q("SELECT COUNT(*) as n FROM aditivo WHERE contrato_id = %s", (id,), one=True)['n']
    _info_adi = _q("""SELECT e.nome est, emp.nome emp FROM contrato ct
                      JOIN estagiario e ON e.id=ct.estagiario_id
                      JOIN empresa emp ON emp.id=ct.empresa_id WHERE ct.id=%s""", (id,), one=True)
    _log('criar', 'aditivo', id,
         f'Gerou {numero}º aditivo: {_info_adi["est"]} @ {_info_adi["emp"]}' if _info_adi
         else f'Gerou {numero}º aditivo do contrato ID {id}')
    return jsonify({'ok': True, 'numero': numero})


@app.route('/api/check-vinculo/<int:est_id>/<int:emp_id>')
@login_required
def api_check_vinculo(est_id, emp_id):
    exclude_id = request.args.get('exclude', type=int)
    sql = "SELECT data_inicio, data_encerramento, data_fim FROM contrato WHERE estagiario_id = %s AND empresa_id = %s"
    params = [est_id, emp_id]
    if exclude_id:
        sql += " AND id != %s"
        params.append(exclude_id)
    cts = _q(sql, params)
    total_dias = 0
    for ct in cts:
        try:
            ini = datetime.strptime(str(ct['data_inicio'])[:10], '%Y-%m-%d').date()
            fim_str = ct['data_encerramento'] or ct['data_fim']
            fim = datetime.strptime(str(fim_str)[:10], '%Y-%m-%d').date()
            total_dias += (fim - ini).days
        except Exception:
            pass
    anos = round(total_dias / 365.25, 2)
    return jsonify({'total_dias': total_dias, 'anos': anos,
                    'limite_atingido': total_dias >= 730, 'aviso': total_dias >= 600})


# ─── ADMIN — CONFIGURAÇÕES ────────────────────────────────────────────────────

@app.route('/admin/config', methods=['GET', 'POST'])
@admin_required
def admin_config():
    if request.method == 'POST':
        for chave in ['seg_seguradora', 'seg_apolice', 'seg_coberturas', 'seg_vigencia']:
            valor = request.form.get(chave, '').strip()
            _run("INSERT INTO config (chave, valor) VALUES (%s, %s) ON CONFLICT (chave) DO UPDATE SET valor = EXCLUDED.valor",
                 (chave, valor))
        _log('editar', 'config', None, 'Editou configurações de seguro')
        flash('Configurações salvas!', 'success')
        return redirect(url_for('admin_config'))
    cfg = _get_config()
    return render_template('admin/config.html', cfg=cfg)


# ─── RELATÓRIOS ───────────────────────────────────────────────────────────────

@app.route('/relatorio/vencimentos')
@login_required
def relatorio_vencimentos():
    data_ate = request.args.get('data_ate', '')
    base_sql = """
        SELECT c.*, e.nome est_nome, emp.nome emp_nome, ie.sigla ie_sigla,
               COALESCE(
                   c.data_encerramento,
                   (SELECT nova_data_fim FROM aditivo
                    WHERE contrato_id = c.id AND nova_data_fim IS NOT NULL AND nova_data_fim != ''
                    ORDER BY created_at DESC LIMIT 1),
                   c.data_fim
               ) as effective_data_fim
        FROM contrato c
        JOIN estagiario e ON e.id = c.estagiario_id
        JOIN empresa emp ON emp.id = c.empresa_id
        JOIN ie ON ie.id = c.ie_id
    """
    if data_ate:
        sql = f"WITH eff AS ({base_sql}) SELECT * FROM eff WHERE effective_data_fim <= %s ORDER BY effective_data_fim"
        rows = _q(sql, (data_ate,))
    else:
        sql = f"WITH eff AS ({base_sql}) SELECT * FROM eff ORDER BY effective_data_fim"
        rows = _q(sql)
    return render_template('relatorio/vencimentos.html', contratos=rows, data_ate=data_ate)


@app.route('/relatorio/estagiarios')
@login_required
def relatorio_estagiarios():
    empresa_id = request.args.get('empresa_id', '')
    ie_id = request.args.get('ie_id', '')
    curso = request.args.get('curso', '').strip()

    sql = """
        SELECT c.*, e.nome est_nome, e.cpf est_cpf, e.semestre est_semestre,
               e.tipo_ensino est_tipo_ensino, e.matricula est_matricula,
               emp.nome emp_nome, ie.nome ie_nome, ie.sigla ie_sigla
        FROM contrato c
        JOIN estagiario e ON e.id = c.estagiario_id
        JOIN empresa emp ON emp.id = c.empresa_id
        JOIN ie ON ie.id = c.ie_id
        WHERE 1=1
    """
    params = []
    if empresa_id:
        sql += " AND c.empresa_id = %s"
        params.append(empresa_id)
    if ie_id:
        sql += " AND c.ie_id = %s"
        params.append(ie_id)
    if curso:
        sql += " AND c.curso ILIKE %s"
        params.append(f'%{curso}%')
    sql += " ORDER BY emp.nome, e.nome"

    contratos = _q(sql, params)
    empresas = _q("SELECT id, nome FROM empresa WHERE status='ativo' ORDER BY nome")
    ies = _q("SELECT id, nome FROM ie ORDER BY nome")
    total_taxa = sum(c['taxa'] or 0 for c in contratos)
    total_bolsa = sum(c['bolsa'] or 0 for c in contratos)
    empresa_sel = _q("SELECT nome FROM empresa WHERE id = %s", (empresa_id,), one=True) if empresa_id else None

    return render_template('relatorio/estagiarios.html',
                           contratos=contratos, empresas=empresas, ies=ies,
                           empresa_id=empresa_id, ie_id=ie_id, curso=curso,
                           total_taxa=total_taxa, total_bolsa=total_bolsa,
                           empresa_sel=empresa_sel, agente=AGENTE,
                           fmt_date=fmt_date, data_hoje=date.today())


# ─── CADASTRO PÚBLICO ────────────────────────────────────────────────────────

@app.route('/cadastro/estagiario', methods=['GET', 'POST'])
def cadastro_estagiario():
    if request.method == 'POST':
        try:
            _ins("""INSERT INTO estagiario
                    (nome,cpf,rg,data_nascimento,telefone,email,endereco,obs,status)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,'pendente')""",
                 (request.form['nome'], request.form['cpf'],
                  request.form.get('rg'), request.form.get('data_nascimento') or None,
                  request.form.get('telefone'), request.form.get('email'),
                  request.form.get('endereco'), request.form.get('obs')))
            return render_template('cadastro/sucesso.html', tipo='estagiário')
        except psycopg2.errors.UniqueViolation:
            flash('Este CPF já está cadastrado no sistema.', 'danger')
    return render_template('cadastro/estagiario.html')


@app.route('/cadastro/empresa', methods=['GET', 'POST'])
def cadastro_empresa():
    if request.method == 'POST':
        _ins("""INSERT INTO empresa
                (nome,cnpj,endereco,cidade,telefone,email,ramo,
                 representante,cargo_representante,cpf_representante,supervisor_nome,supervisor_cargo,status)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'pendente')""",
             (request.form['nome'], request.form.get('cnpj'),
              request.form.get('endereco'), request.form.get('cidade', 'Vitória da Conquista'),
              request.form.get('telefone'), request.form.get('email'),
              request.form.get('ramo'), request.form.get('representante'),
              request.form.get('cargo_representante'), request.form.get('cpf_representante'),
              request.form.get('supervisor_nome'), request.form.get('supervisor_cargo')))
        return render_template('cadastro/sucesso.html', tipo='empresa')
    return render_template('cadastro/empresa.html')


# ─── ADMIN — PENDENTES ────────────────────────────────────────────────────────

@app.route('/admin/pendentes')
@admin_required
def admin_pendentes():
    estagiarios_p = _q("SELECT * FROM estagiario WHERE status='pendente' ORDER BY nome")
    empresas_p = _q("SELECT * FROM empresa WHERE status='pendente' ORDER BY nome")
    return render_template('admin/pendentes.html', estagiarios=estagiarios_p, empresas=empresas_p)


@app.route('/admin/pendentes/estagiario/<int:id>/aprovar')
@admin_required
def aprovar_estagiario(id):
    reg = _q("SELECT nome FROM estagiario WHERE id = %s", (id,), one=True)
    _run("UPDATE estagiario SET status='ativo' WHERE id=%s", (id,))
    _log('aprovar', 'estagiario', id, f'Aprovou estagiário: {reg["nome"] if reg else id}')
    flash('Estagiário aprovado com sucesso!', 'success')
    return redirect(url_for('admin_pendentes'))


@app.route('/admin/pendentes/empresa/<int:id>/aprovar')
@admin_required
def aprovar_empresa(id):
    reg = _q("SELECT nome FROM empresa WHERE id = %s", (id,), one=True)
    _run("UPDATE empresa SET status='ativo' WHERE id=%s", (id,))
    _log('aprovar', 'empresa', id, f'Aprovou empresa: {reg["nome"] if reg else id}')
    flash('Empresa aprovada com sucesso!', 'success')
    return redirect(url_for('admin_pendentes'))


@app.route('/admin/pendentes/estagiario/<int:id>/rejeitar')
@admin_required
def rejeitar_estagiario(id):
    reg = _q("SELECT nome FROM estagiario WHERE id = %s AND status='pendente'", (id,), one=True)
    _run("DELETE FROM estagiario WHERE id=%s AND status='pendente'", (id,))
    _log('rejeitar', 'estagiario', id, f'Rejeitou estagiário: {reg["nome"] if reg else id}')
    flash('Cadastro rejeitado e removido.', 'warning')
    return redirect(url_for('admin_pendentes'))


@app.route('/admin/pendentes/empresa/<int:id>/rejeitar')
@admin_required
def rejeitar_empresa(id):
    reg = _q("SELECT nome FROM empresa WHERE id = %s AND status='pendente'", (id,), one=True)
    _run("DELETE FROM empresa WHERE id=%s AND status='pendente'", (id,))
    _log('rejeitar', 'empresa', id, f'Rejeitou empresa: {reg["nome"] if reg else id}')
    flash('Cadastro rejeitado e removido.', 'warning')
    return redirect(url_for('admin_pendentes'))


# ─── ADMIN — LOG DE AUDITORIA ─────────────────────────────────────────────────

@app.route('/admin/log')
@admin_required
def admin_log():
    periodo = request.args.get('periodo', '30')
    entidade = request.args.get('entidade', '')
    uid_filtro = request.args.get('usuario_id', '')
    try:
        dias = int(periodo)
    except ValueError:
        dias = 30
    dt_from = datetime.now() - timedelta(days=dias)
    sql = """SELECT l.*, TO_CHAR(l.created_at, 'DD/MM/YYYY HH24:MI:SS') as dt_fmt
             FROM log_auditoria l WHERE l.created_at >= %s"""
    params = [dt_from]
    if entidade:
        sql += " AND l.entidade = %s"
        params.append(entidade)
    if uid_filtro:
        sql += " AND l.usuario_id = %s"
        params.append(uid_filtro)
    sql += " ORDER BY l.created_at DESC LIMIT 500"
    logs = _q(sql, params)
    usuarios = _q("SELECT id, COALESCE(nome, username) as nome FROM usuario ORDER BY nome")
    return render_template('admin/log.html', logs=logs, periodo=periodo,
                           entidade=entidade, usuario_id=uid_filtro, usuarios=usuarios)


# ─── ERROS ────────────────────────────────────────────────────────────────────

@app.errorhandler(403)
def forbidden(e):
    return render_template('erro.html', codigo=403,
                           msg='Acesso negado. Apenas o administrador pode acessar esta área.'), 403


@app.errorhandler(404)
def not_found(e):
    return render_template('erro.html', codigo=404, msg='Página não encontrada.'), 404


init_db()

if __name__ == '__main__':
    print('\n' + '=' * 50)
    print('  CICLO RH — Sistema de Estágio')
    print('  Acesse: http://localhost:5000')
    print('=' * 50 + '\n')
    app.run(debug=False, port=5000)
