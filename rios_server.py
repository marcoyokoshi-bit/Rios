"""
RIOS v1.12.0 - Restaurant Intelligence OS
"""
import json, re, os, sys, glob, base64, tempfile, shutil
from datetime import datetime

try:
    from flask import Flask, request, jsonify, send_from_directory, send_file
    from flask_cors import CORS
    import anthropic
except ImportError:
    os.system(f"{sys.executable} -m pip install flask flask-cors anthropic python-dotenv psycopg2-binary python-docx Pillow --quiet")
    from flask import Flask, request, jsonify, send_from_directory, send_file
    from flask_cors import CORS
    import anthropic

try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    os.system(f"{sys.executable} -m pip install psycopg2-binary --quiet")
    import psycopg2
    import psycopg2.extras

try:
    import docx  # noqa: F401  (garante que python-docx está instalado)
    from PIL import Image  # noqa: F401
except ImportError:
    os.system(f"{sys.executable} -m pip install python-docx Pillow --quiet")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rios_pop_docx import build_pop_docx, docx_filename  # noqa: E402

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
app        = Flask(__name__)
CORS(app)
API_KEY    = os.environ.get("ANTHROPIC_API_KEY", "").strip()
PORT       = int(os.environ.get("RIOS_PORT", 8765))
MAX_UPLOAD = 10 * 1024 * 1024
RULES_FILE = os.path.join(BASE_DIR, "rios_rules.json")
AI_PROVIDER = os.environ.get("RIOS_AI_PROVIDER", "anthropic").lower()
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

# ── Banco de dados (Supabase / Postgres) ─────────────────────────────────────
# Conexão criada sob demanda (não mantém conexão global aberta).
# DATABASE_URL fica só no .env local — nunca aparece em código ou logs.
def get_db():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL não configurada no .env")
    return psycopg2.connect(DATABASE_URL, connect_timeout=8)

TASK_MODELS = {
    "ingredient": "claude-haiku-4-5-20251001",
    "nf":         "claude-sonnet-4-6",
    "pdf_nf":     "claude-sonnet-4-6",
    "preparo":    "claude-sonnet-4-6",
    "descricao":  "claude-haiku-4-5-20251001",
    "alergenos":  "claude-haiku-4-5-20251001",
    "preco":      "claude-haiku-4-5-20251001",
    "receita":    "claude-sonnet-4-6",
    "analise":    "claude-opus-4-8",
    "insights":   "claude-sonnet-4-6",
    "qc_visual":  "claude-sonnet-4-6",
    "escala":     "claude-sonnet-4-6",
    "importar_receita":  "claude-sonnet-4-6",
    "estimar_pendencia": "claude-haiku-4-5-20251001",
    "buffet_analise":    "claude-sonnet-4-6",
    "default":    "claude-sonnet-4-6",
}

def call_ai(task, messages, system=None, max_tokens=2048):
    if not API_KEY:
        raise ValueError("API Key nao configurada")
    model = TASK_MODELS.get(task, TASK_MODELS["default"])
    client = anthropic.Anthropic(api_key=API_KEY)
    kwargs = {"model": model, "max_tokens": max_tokens, "messages": messages}
    if system:
        kwargs["system"] = system
    return client.messages.create(**kwargs).content[0].text

CHAT_PROMPTS = {
    "preparo": (
        "Voce e o Chef Marco A Souza, chef e consultor gastronomico. "
        "Gere um modo de preparo profissional. "
        'Retorne SOMENTE JSON: {"mise":"...","steps":["..."],"final":"..."}'
    ),
    "descricao": (
        "Crie uma descricao elegante para cardapio. "
        'Retorne SOMENTE JSON: {"descricao":"..."}'
    ),
    "alergenos": (
        "Identifique alergenos (ANVISA/RDC 26/2015). "
        'Retorne SOMENTE JSON: {"alergenos":["..."],"observacao":"..."}'
    ),
    "preco": (
        "Sugira precificacao profissional (CMV 25-35%, markup 2.5x-4x). "
        'Retorne SOMENTE JSON: {"preco_sugerido":39.90,"justificativa":"...","faixa_min":32.0,"faixa_max":48.0,"estrategia":"..."}'
    ),
    "receita": (
        "Sugira receita padrao completa para o prato. "
        'Retorne SOMENTE JSON: {"ingredientes":[{"nome":"...","quantidade":"...","unidade":"..."}],"utensilios":["..."],"descricao":"...","observacao":"..."}'
    ),
    "analise": (
        "Analise critica e profissional da ficha tecnica. Nota 0-10. "
        'Retorne SOMENTE JSON: {"nota":8.5,"pontos_fortes":["..."],"pontos_melhoria":["..."],"alerta_financeiro":"...","recomendacao":"..."}'
    ),
}

# ── Importar Receita (texto bruto -> ficha(s) tecnica(s) estruturada(s)) ─────
# Nao usa CHAT_PROMPTS/build_context_text porque o contexto de entrada aqui e'
# texto livre colado pelo usuario, nao os campos de uma ficha ja preenchida.
IMPORT_RECEITA_SYSTEM = (
    "Voce e o Chef Marco A Souza, chef e consultor gastronomico especializado em fichas tecnicas de producao. "
    "Voce recebe um texto bruto colado pelo usuario, que pode conter uma ou varias receitas em formato informal "
    "(lista de caderno, ficha antiga, anotacao solta, cardapio, receita de familia, etc). Transforme CADA receita "
    "encontrada no texto em uma ficha tecnica estruturada, no padrao profissional de fichas tecnicas de producao.\n\n"
    "REGRA CRITICA: nunca invente um valor numerico (peso, rendimento, tempo de preparo/coccao, fator de correcao, "
    "quantidade de ingrediente) que nao esteja no texto original e que nao possa ser deduzido com confianca "
    "razoavel a partir de uma conversao padrao de medida caseira (ex.: '1 xicara de arroz' pode virar gramas; "
    "'tempero a gosto' NAO deve virar um peso em gramas). Quando um dado nao estiver claro ou ausente no texto, "
    "NAO chute — deixe o campo vazio (string vazia ou 0) e registre uma pendencia em pendencias explicando "
    "exatamente o que falta e por que. Isso vale tanto para campos gerais da ficha (rendimento, tempo, peso da "
    "porcao, categoria) quanto para ingredientes especificos (cite o nome do ingrediente na pendencia).\n\n"
    "Categorias validas para o campo cat: Entrada, Prato Principal, Guarnicao, Sobremesa, Bebida, Molho, Outro.\n\n"
    "Retorne SOMENTE JSON no formato: "
    '{"fichas":[{"nome":"...","cat":"...","rend":"...","pporcao":"...","tprep":"...","tcook":"...","equip":"...",'
    '"alerg":"...","desc":"...","mise":"...","final":"...","steps":["..."],'
    '"ings":[{"nome":"...","pb":"...","un":"g|kg|ml|L|un|cx|dz|pct","fc":"1.000"}],'
    '"pendencias":[{"campo":"...","descricao":"..."}]}]}'
)

ESTIMAR_PENDENCIA_SYSTEM = (
    "Voce e o Chef Marco A Souza, chef e consultor gastronomico. O usuario esta preenchendo uma ficha tecnica "
    "importada de uma receita e pediu para voce estimar, com sua experiencia profissional, um valor razoavel "
    "para um dado especifico que a receita original nao deixou claro. De sua melhor estimativa profissional, "
    "curta e objetiva, e deixe explicito que e uma ESTIMATIVA — nao um dado confirmado pela receita original.\n\n"
    'Retorne SOMENTE JSON: {"valor_estimado":"...","justificativa":"..."}'
)

BUFFET_ANALISE_SYSTEM = (
    "Voce e o Chef Marco A Souza, chef e consultor gastronomico especializado em engenharia de cardapio "
    "para restaurantes de comida a Kg (buffet por peso). Voce recebe a composicao de um buffet montado "
    "(pratos, grupo proteico de cada um, Kg produzido, custo por Kg, CMV individual e geral, e os alertas "
    "automaticos ja calculados pelo sistema — repeticao de proteina e pratos com custo acima da media). "
    "Sua tarefa e dar um parecer estrategico curto e pratico: avaliar o equilibrio do mix (variedade de "
    "proteinas, cores, texturas, guarnicoes), comentar se o CMV geral esta saudavel, e sugerir combinacoes "
    "ou trocas concretas quando fizer sentido (ex: trocar uma proteina bovina repetida por uma opcao de "
    "ave ou peixe, ajustar Kg produzido de um prato caro, etc). Seja direto, sem promessas magicas e sem "
    "recalcular numeros — use os numeros que ja foram enviados.\n\n"
    'Retorne SOMENTE JSON: {"nota_equilibrio":7.5,"pontos_fortes":["..."],"pontos_atencao":["..."],'
    '"sugestoes_combinacao":["..."],"recomendacao_final":"..."}'
)

INSIGHTS_SYSTEM = (
    "Voce e o RIOS Cortex, sistema de inteligencia gastronomica. "
    "Analise dados de QC e gere insights estrategicos acionaveis. "
    "Seja direto e especifico. Retorne SOMENTE JSON valido."
)

def parse_json(text, task):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        return jsonify({"success": True, "data": json.loads(text), "task": task})
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if m:
            try:
                return jsonify({"success": True, "data": json.loads(m.group()), "task": task})
            except Exception:
                pass
    return jsonify({"success": False, "raw": text, "error": "Nao foi possivel parsear JSON."})

def check_upload_size(b64):
    return len(b64.encode("utf-8")) <= MAX_UPLOAD

def build_context_text(ctx):
    lines = []
    if ctx.get("nome"):  lines.append(f"Prato: {ctx['nome']}")
    if ctx.get("cat"):   lines.append(f"Categoria: {ctx['cat']}")
    if ctx.get("rend"):  lines.append(f"Rendimento: {ctx['rend']} porcao(oes)")
    if ctx.get("ings"):
        ings_txt = ", ".join([
            f"{i.get('nome','?')} ({i.get('pb','?')} {i.get('un','un')})"
            for i in ctx["ings"] if i.get("nome")
        ])
        lines.append(f"Ingredientes: {ings_txt}")
    for k, lbl in [("custo_ing","Custo ing/porcao R$"),("custo_total","Custo total/porcao R$"),
                   ("preco","Preco venda R$"),("cmv","CMV meta %"),("oh","Overhead %"),("mo","Mao obra %")]:
        if ctx.get(k): lines.append(f"{lbl}: {ctx[k]}")
    if ctx.get("steps"): lines.append(f"Passos: {'; '.join(ctx['steps'])}")
    if ctx.get("mise"):  lines.append(f"Mise: {ctx['mise']}")
    if ctx.get("alerg"): lines.append(f"Alergenos: {ctx['alerg']}")
    return "\n".join(lines) if lines else "Nenhum dado fornecido."

BACKUP_DIR = os.path.join(BASE_DIR, "backups")
os.makedirs(BACKUP_DIR, exist_ok=True)
MAX_BACKUPS = 30
QC_DIR = os.path.join(BASE_DIR, "qc_registros")
os.makedirs(QC_DIR, exist_ok=True)

def load_rules():
    try:
        if os.path.exists(RULES_FILE):
            with open(RULES_FILE, encoding="utf-8") as fh:
                return json.load(fh)
    except Exception:
        pass
    return {"metas": {}, "regras": []}

def save_rules(data):
    data["atualizado"] = datetime.now().strftime("%Y-%m-%d")
    with open(RULES_FILE, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

def check_rules_logic(metrics, config):
    OPS = {"<": lambda a,b: a<b, "<=": lambda a,b: a<=b,
           ">": lambda a,b: a>b, ">=": lambda a,b: a>=b, "==": lambda a,b: a==b}
    alertas = []
    for regra in config.get("regras", []):
        if not regra.get("ativa", True): continue
        campo, op_str, limiar = regra.get("campo",""), regra.get("operador","<"), regra.get("limiar",0)
        op_fn = OPS.get(op_str)
        if op_fn is None or campo not in metrics: continue
        valor = metrics[campo]
        if op_fn(valor, limiar):
            msg = regra.get("mensagem","Alerta").replace("{valor}", str(round(valor,1))).replace("{limiar}", str(limiar))
            alertas.append({
                "id": regra["id"], "nome": regra.get("nome",""),
                "mensagem": msg, "severidade": regra.get("severidade","aviso"),
                "icone": regra.get("icone","⚠️"),
                "campo": campo, "valor": valor, "limiar": limiar,
            })
    return alertas

# ── HELPER: serve HTML sem cache ─────────────────────────────────────────────
def serve_html(filename):
    from flask import make_response
    resp = make_response(send_from_directory(BASE_DIR, filename))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"]        = "no-cache"
    resp.headers["Expires"]       = "0"
    return resp

# ── ROTAS ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return serve_html("RIOS_Home.html")

@app.route("/fichas")
def fichas():
    return serve_html("RIOS_FichaTecnica.html")

@app.route("/qc")
def qc_visual():
    return serve_html("RIOS_QC_Visual.html")

@app.route("/historico")
def qc_historico():
    return serve_html("RIOS_QC_Historico.html")

@app.route("/insights")
def insights_page():
    return serve_html("RIOS_Insights.html")

@app.route("/escala")
def escala_page():
    return serve_html("RIOS_Escala.html")

@app.route("/pop")
def pop_page():
    return serve_html("RIOS_POP.html")

@app.route("/buffet")
def buffet_page():
    return serve_html("RIOS_CardapioBuffet.html")

@app.route("/produtos")
def produtos_page():
    return serve_html("RIOS_Produtos.html")

@app.route("/api/status")
def api_status():
    key_ok = bool(API_KEY and API_KEY.startswith("sk-ant"))
    return jsonify({"status":"ok","version":"1.13.0","key_configured":key_ok,"ai_provider":AI_PROVIDER,
                     "database_configured": bool(DATABASE_URL)})

@app.route("/api/pop/docx", methods=["POST"])
def pop_docx_route():
    """Gera o arquivo .docx do POP no padrão definido no Prompt Mestre
    (capa técnica, histórico de revisão, marca-d'água, cabeçalho institucional
    repetido, seções numeradas, fluxograma em página própria)."""
    tmp_dir = None
    try:
        data = request.json or {}
        pop = data.get("pop") or {}
        cliente = data.get("cliente") or {}
        logo_b64 = data.get("logoBase64") or ""
        if not pop.get("titulo"):
            return jsonify({"error": "POP sem título — preencha antes de gerar o Word."}), 400

        tmp_dir = tempfile.mkdtemp(prefix="rios_pop_")
        logo_path = None
        if logo_b64:
            try:
                header, b64data = logo_b64.split(",", 1) if "," in logo_b64 else ("", logo_b64)
                raw = base64.b64decode(b64data)
                logo_path = os.path.join(tmp_dir, "_logo_cliente.png")
                with open(logo_path, "wb") as fh:
                    fh.write(raw)
            except Exception:
                logo_path = None

        out_name = docx_filename(pop)
        out_path = os.path.join(tmp_dir, out_name)
        build_pop_docx(pop, cliente, out_path, logo_path=logo_path,
                        color_hex=cliente.get("corHex"), tmp_dir=tmp_dir)

        # Não apaga a pasta temporária aqui: send_file pode enviar em streaming e o
        # arquivo precisa existir até a resposta terminar. Fica em tmp do SO
        # (Windows limpa periodicamente) — cada geração usa poucos KB.
        return send_file(out_path, as_attachment=True, download_name=out_name,
                          mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        return jsonify({"error": f"Falha ao gerar Word: {e}"}), 500

@app.route("/api/db/status")
def db_status():
    if not DATABASE_URL:
        return jsonify({"connected": False, "error": "DATABASE_URL não configurada no .env"}), 200
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("select current_database(), current_user;")
        db_name, db_user = cur.fetchone()
        cur.execute("""
            select table_name from information_schema.tables
            where table_schema='public' order by table_name;
        """)
        tables = [r[0] for r in cur.fetchall()]
        conn.close()
        return jsonify({"connected": True, "database": db_name, "user": db_user, "tables": tables})
    except Exception as e:
        return jsonify({"connected": False, "error": str(e)}), 200

@app.route("/api/backup", methods=["POST"])
def backup_save():
    try:
        data = request.json or {}
        if not data: return jsonify({"error":"Sem dados"}), 400
        ts    = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        fname = f"rios_backup_{ts}.json"
        with open(os.path.join(BACKUP_DIR, fname), "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        all_b = sorted(glob.glob(os.path.join(BACKUP_DIR, "rios_backup_*.json")))
        for old in all_b[:-MAX_BACKUPS]:
            try: os.remove(old)
            except: pass
        return jsonify({"success":True,"file":fname,"fichas":len(data.get("fichas",[]))})
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/backup/list")
def backup_list():
    try:
        files  = sorted(glob.glob(os.path.join(BACKUP_DIR,"rios_backup_*.json")), reverse=True)
        result = []
        for fp in files[:20]:
            try:
                with open(fp, encoding="utf-8") as fh: d = json.load(fh)
                nf = len(d.get("fichas",[]))
            except: nf = "?"
            result.append({"file":os.path.basename(fp),"size_kb":round(os.path.getsize(fp)/1024,1),"fichas":nf})
        return jsonify({"backups":result})
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/backup/restore/<filename>")
def backup_restore(filename):
    try:
        if not filename.startswith("rios_backup_") or not filename.endswith(".json"):
            return jsonify({"error":"Arquivo invalido"}), 400
        fp = os.path.join(BACKUP_DIR, filename)
        if not os.path.exists(fp): return jsonify({"error":"Nao encontrado"}), 404
        with open(fp, encoding="utf-8") as fh: data = json.load(fh)
        return jsonify({"success":True,"data":data})
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/vision", methods=["POST"])
def vision():
    try:
        if not API_KEY: return jsonify({"error":"API Key ausente"}), 401
        data      = request.json or {}
        image_b64 = data.get("image","")
        media_tp  = data.get("media_type","image/jpeg")
        task      = data.get("task")
        if task not in ("ingredient","nf","pdf_nf"):
            return jsonify({"error":f"Tarefa desconhecida: {task}"}), 400
        if not check_upload_size(image_b64):
            return jsonify({"error":"Arquivo muito grande (limite 10MB)"}), 413
        if task == "ingredient":
            prompt = 'Identifique o produto. Retorne SOMENTE JSON: {"nome":"...","unidade":"g/kg/ml/L/un","observacao":"..."}'
        else:
            prompt = 'Extraia itens da NF/cupom. Retorne SOMENTE JSON: {"itens":[{"nome":"...","quantidade":0.5,"unidade":"kg","preco_unitario":12.90}]}'
        model  = TASK_MODELS.get(task, "claude-sonnet-4-6")
        client = anthropic.Anthropic(api_key=API_KEY)
        if task == "pdf_nf":
            content = [{"type":"document","source":{"type":"base64","media_type":"application/pdf","data":image_b64}},{"type":"text","text":prompt}]
        else:
            content = [{"type":"image","source":{"type":"base64","media_type":media_tp,"data":image_b64}},{"type":"text","text":prompt}]
        msg = client.messages.create(model=model, max_tokens=2048, messages=[{"role":"user","content":content}])
        return parse_json(msg.content[0].text, task)
    except anthropic.AuthenticationError:
        return jsonify({"error":"API Key invalida"}), 401
    except anthropic.RateLimitError:
        return jsonify({"error":"Rate limit. Aguarde."}), 429
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/chat", methods=["POST"])
def chat():
    try:
        if not API_KEY: return jsonify({"error":"API Key ausente"}), 401
        data    = request.json or {}
        task    = data.get("task")
        context = data.get("context", {})
        model   = TASK_MODELS.get(task, "claude-sonnet-4-6")
        client  = anthropic.Anthropic(api_key=API_KEY)
        if "messages" in data:
            messages = data["messages"]
            for msg in messages:
                content = msg.get("content")
                if isinstance(content, list):
                    for part in content:
                        if isinstance(part, dict) and part.get("type") == "image":
                            b64 = part.get("source",{}).get("data","")
                            if not check_upload_size(b64):
                                return jsonify({"error":"Imagem muito grande"}), 413
                # content como string simples (texto puro) e' valido e nao precisa de checagem de imagem
            resp = client.messages.create(model=model, max_tokens=2048, messages=messages)
            return jsonify({"content":[{"text":resp.content[0].text}],"task":task})
        if task == "importar_receita":
            texto = (context.get("texto") or "").strip()
            if not texto:
                return jsonify({"error": "Cole o texto da receita antes de importar."}), 400
            user_msg = (f"TEXTO COLADO PELO USUARIO:\n\n{texto}\n\n"
                        "Transforme em ficha(s) tecnica(s) estruturada(s) conforme instruido.")
            resp = client.messages.create(model=model, max_tokens=4096,
                system=IMPORT_RECEITA_SYSTEM, messages=[{"role":"user","content":user_msg}])
            return parse_json(resp.content[0].text, task)
        if task == "estimar_pendencia":
            campo    = context.get("campo","")
            motivo   = context.get("descricao","")
            fichaCtx = context.get("ficha_contexto","")
            user_msg = (f"Prato/contexto conhecido:\n{fichaCtx}\n\n"
                        f"Campo pendente: {campo}\nMotivo da pendencia: {motivo}\n\n"
                        "De sua melhor estimativa profissional para este campo.")
            resp = client.messages.create(model=model, max_tokens=500,
                system=ESTIMAR_PENDENCIA_SYSTEM, messages=[{"role":"user","content":user_msg}])
            return parse_json(resp.content[0].text, task)
        if task == "buffet_analise":
            resumo = context.get("resumo","")
            if not resumo:
                return jsonify({"error": "Monte o buffet antes de pedir a análise da IA."}), 400
            resp = client.messages.create(model=model, max_tokens=1500,
                system=BUFFET_ANALISE_SYSTEM, messages=[{"role":"user","content":resumo}])
            return parse_json(resp.content[0].text, task)
        if task not in CHAT_PROMPTS:
            return jsonify({"error":f"Tarefa desconhecida: {task}"}), 400
        user_msg = f"DADOS DA FICHA TECNICA:\n{build_context_text(context)}\n\nGere o resultado solicitado."
        resp = client.messages.create(model=model, max_tokens=2048,
            system=CHAT_PROMPTS[task], messages=[{"role":"user","content":user_msg}])
        return parse_json(resp.content[0].text, task)
    except anthropic.AuthenticationError:
        return jsonify({"error":"API Key invalida"}), 401
    except anthropic.RateLimitError:
        return jsonify({"error":"Rate limit. Aguarde."}), 429
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/qc/save", methods=["POST"])
def qc_save():
    try:
        data = request.json or {}
        if not data.get("num"): return jsonify({"error":"Numero ausente"}), 400
        ts    = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        num   = re.sub(r"[^A-Za-z0-9\-_]", "", str(data.get("num","X")))
        fname = f"qc_{ts}_pedido{num}.json"
        reg   = {k:v for k,v in data.items() if k not in ("photoSaida","photoRef")}
        reg["arquivo"] = fname
        with open(os.path.join(QC_DIR, fname), "w", encoding="utf-8") as fh:
            json.dump(reg, fh, ensure_ascii=False, indent=2)
        return jsonify({"success":True,"file":fname})
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/qc/history")
def qc_history():
    try:
        files = sorted(glob.glob(os.path.join(QC_DIR,"qc_*.json")), reverse=True)
        result = []
        for fp in files[:100]:
            try:
                with open(fp, encoding="utf-8") as fh: result.append(json.load(fh))
            except: pass
        return jsonify({"registros":result,"total":len(result)})
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/rules")
def rules_get():
    try:
        return jsonify(load_rules())
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/rules", methods=["POST"])
def rules_save_route():
    try:
        data = request.json or {}
        if not data: return jsonify({"error":"Sem dados"}), 400
        save_rules(data)
        return jsonify({"success":True})
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/rules/check", methods=["POST"])
def rules_check():
    try:
        data    = request.json or {}
        metrics = data.get("metrics", {})
        config  = load_rules()
        alertas = check_rules_logic(metrics, config)
        return jsonify({
            "alertas":       alertas,
            "total_alertas": len(alertas),
            "criticos":  sum(1 for a in alertas if a["severidade"]=="critico"),
            "avisos":    sum(1 for a in alertas if a["severidade"]=="aviso"),
            "infos":     sum(1 for a in alertas if a["severidade"]=="info"),
        })
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route("/api/insights/daily", methods=["POST"])
def insights_daily():
    try:
        if not API_KEY: return jsonify({"error":"API Key ausente"}), 401
        data   = request.json or {}
        resumo = data.get("resumo", {})
        linhas = ["DADOS DE QC - ULTIMOS 7 DIAS\n"]
        linhas.append(f"Total analises: {resumo.get('total',0)}")
        linhas.append(f"Aprovados: {resumo.get('aprovados',0)}")
        linhas.append(f"Reprovados: {resumo.get('refazer',0)}")
        linhas.append(f"Aceitos c/ falha: {resumo.get('aceito_falha',0)}")
        linhas.append(f"Taxa aprovacao: {resumo.get('taxa_aprovacao',0):.1f}%")
        linhas.append(f"Score medio: {resumo.get('score_medio',0):.1f}%")
        linhas.append(f"Total fichas tecnicas: {data.get('total_fichas',0)}")
        for d in (data.get("por_dia") or [])[:7]:
            t = d.get("taxa")
            linhas.append(f"  {d.get('data','?')}: {d.get('total',0)} analises, {t:.0f}% aprovacao" if t is not None else f"  {d.get('data','?')}: sem dados")
        for p in (data.get("pratos_problematicos") or [])[:5]:
            linhas.append(f"  {p.get('nome','?')}: {p.get('taxa_falha',0):.0f}% falha ({p.get('reprovados',0)}/{p.get('total',0)})")
        for f in (data.get("falhas_frequentes") or [])[:8]:
            linhas.append(f"  {f.get('item','?')}: {f.get('ocorrencias',0)} ocorrencias")
        prompt = (
            "\n".join(linhas) + "\n\n"
            "Gere analise estrategica concisa e acionavel. "
            'Retorne SOMENTE JSON: {"resumo_executivo":"...","evolucao":"Melhorando|Estavel|Piorando",'
            '"nota_geral":7.5,"pontos_criticos":["..."],'
            '"pratos_atencao":[{"nome":"...","problema":"...","acao":"..."}],'
            '"acoes_imediatas":["..."],"acoes_semana":["..."],"mensagem_chef":"..."}'
        )
        raw = call_ai("insights", [{"role":"user","content":prompt}], system=INSIGHTS_SYSTEM, max_tokens=1500)
        return parse_json(raw, "insights")
    except anthropic.AuthenticationError:
        return jsonify({"error":"API Key invalida"}), 401
    except anthropic.RateLimitError:
        return jsonify({"error":"Rate limit. Aguarde."}), 429
    except Exception as e:
        return jsonify({"error":str(e)}), 500

if __name__ == "__main__":
    print("\n" + "="*56)
    print("  RIOS v1.12.0 - Restaurant Intelligence OS")
    print(f"  AI Provider: {AI_PROVIDER.upper()}")
    print("="*56)
    if API_KEY:
        print(f"\n  Cortex: Ativo [{API_KEY[:12]}...{API_KEY[-4:]}]")
    else:
        print("\n  ATENCAO: API Key nao encontrada no .env")
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        print(f"\n  http://localhost:{PORT}")
        print(f"  http://{ip}:{PORT}  (Wi-Fi)")
    except Exception:
        print(f"\n  http://localhost:{PORT}")
    print("\n  / | /fichas | /qc | /historico | /insights | /escala | /pop | /buffet")
    print("\n  Ctrl+C para encerrar.\n" + "="*56 + "\n")
    app.run(host="0.0.0.0", port=PORT, debug=False)
