"""
Testa a rota real /api/pop/importar_docx do Flask (rios_server.py) contra um
dos documentos de referência do Chef, sem chamar a API da Anthropic de
verdade (este sandbox não tem API key) — só a chamada de rede pra
anthropic.Anthropic é substituída por um stub; toda a leitura do .docx real
(decodificação base64, extração via rios_pop_import.extract_docx_outline,
montagem do prompt) roda de verdade.
"""
import base64, json, sys
from unittest.mock import patch, MagicMock

sys.path.insert(0, '/tmp/rios_produtos')
import rios_server  # noqa: E402

results = []
def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("OK  " if cond else "FAIL"), name, ("-", detail) if detail else "")

rios_server.API_KEY = "sk-ant-fake-key-for-test"

FAKE_POP_JSON = json.dumps({
    "titulo": "Purê de Batata", "codigo": "POP-COZ-001", "setor": "Cozinha", "versao": "1.0",
    "revisoes": [{"descricao": "Revisão inicial", "revisadoPor": "Chef Marco À Souza", "aprovadoPor": "Marco André", "data": "2026-08-31"}],
    "objetivo": "Padronizar o purê de batata.", "aplicacao": "Cozinha e buffet.",
    "referencias": ["Manual de Boas Práticas da unidade."],
    "materiais": [{"grupo": "Ingredientes", "item": "Purê de batata instantâneo Ajinomoto — 500 g", "obs": "Única marca homologada"}],
    "definicoes": [{"termo": "Buffet", "definicao": "Área de exposição e serviço ao cliente."}],
    "responsaveis": [{"papel": "Cozinha / cozinheira", "tarefas": ["Conferir a programação", "Preparar o purê"]}],
    "frequencia": "Diária",
    "etapas": [{"titulo": "Planejamento e solicitação", "destaque": "", "itens": ["Conferir estoque"]}],
    "pcc": [{"etapa": "Solicitação", "risco": "Falta de insumo", "controle": "Solicitar na véspera", "responsavel": "Cozinha"}],
    "registros": ["Peso total", "Tara"],
    "indicadores": [{"controle": "Textura", "criterio": "Homogênea e sem grumos"}],
    "obs": "", "pendencias": []
})

captured = {}
def fake_create(**kwargs):
    captured.update(kwargs)
    resp = MagicMock()
    resp.content = [MagicMock(text=FAKE_POP_JSON)]
    return resp

with patch.object(rios_server.anthropic, "Anthropic") as MockAnthropic:
    instance = MockAnthropic.return_value
    instance.messages.create.side_effect = fake_create

    client = rios_server.app.test_client()

    with open('/tmp/pop_ref/POP_001_Pure.docx', 'rb') as fh:
        raw = fh.read()
    b64 = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + base64.b64encode(raw).decode()

    resp = client.post('/api/pop/importar_docx', json={"docxBase64": b64})
    check("Rota retornou 200", resp.status_code == 200, resp.status_code)
    data = resp.get_json()
    check("Resposta success=True", data and data.get("success") is True, data)
    check("Dados retornados batem com o stub (título)", data and data.get("data", {}).get("titulo") == "Purê de Batata", data)

    # Confere que o texto extraído do .docx real foi de fato mandado pra IA
    sent_msg = captured.get("messages", [{}])[0].get("content", "")
    check("Prompt mandado pra IA contém texto extraído do Word (Objetivo)", "Padronizar a solicitação de insumos" in sent_msg)
    check("Prompt mandado pra IA contém tabela de Materiais extraída", "Purê de batata instantâneo Ajinomoto" in sent_msg)
    check("System prompt correto foi usado (IMPORT_POP_SYSTEM)", captured.get("system") == rios_server.IMPORT_POP_SYSTEM)
    check("Modelo correto foi usado (importar_pop)", captured.get("model") == rios_server.TASK_MODELS["importar_pop"], captured.get("model"))

    # ── Caso de erro: sem arquivo ──
    resp2 = client.post('/api/pop/importar_docx', json={})
    check("Sem docxBase64 -> erro 400", resp2.status_code == 400, resp2.status_code)

    # ── Caso de erro: base64 inválido ──
    resp3 = client.post('/api/pop/importar_docx', json={"docxBase64": "isso-nao-e-base64-valido!!"})
    check("Base64 inválido -> erro 400", resp3.status_code == 400, resp3.status_code)

    # ── estimar_pendencia_pop via /api/chat ──
    FAKE_ESTIMATE = json.dumps({"valor_estimado": "Setor: Cozinha", "justificativa": "Inferido pelo conteúdo do POP."})
    def fake_create2(**kwargs):
        captured['chat_kwargs'] = kwargs
        resp = MagicMock()
        resp.content = [MagicMock(text=FAKE_ESTIMATE)]
        return resp
    instance.messages.create.side_effect = fake_create2
    resp4 = client.post('/api/chat', json={"task": "estimar_pendencia_pop", "context": {"campo": "setor", "descricao": "não encontrado", "pop_contexto": "POP: Teste"}})
    check("estimar_pendencia_pop -> 200", resp4.status_code == 200, resp4.status_code)
    d4 = resp4.get_json()
    check("estimar_pendencia_pop retorna valor_estimado", d4 and d4.get("data", {}).get("valor_estimado") == "Setor: Cozinha", d4)
    check("estimar_pendencia_pop usou o system prompt certo", captured['chat_kwargs'].get("system") == rios_server.ESTIMAR_PENDENCIA_POP_SYSTEM)

print()
n_fail = sum(1 for _, ok, _ in results if not ok)
print(f"{'='*60}\n{len(results)} checagens, {n_fail} falha(s)")
sys.exit(1 if n_fail else 0)
