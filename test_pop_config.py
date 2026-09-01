import json, sys
from playwright.sync_api import sync_playwright

BASE = "/tmp/rios_produtos"
CLIENTE = {"id": "cli_test_001", "nome": "Mestre do Pão", "criadoEm": "2026-08-01T00:00:00.000Z"}
USER = {"nome": "Chef Teste", "email": "teste@x.com", "label": "Consultor", "cliente": CLIENTE}

results = []
def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("OK  " if cond else "FAIL"), name, ("-", detail) if detail else "")

def seed(page, extra=None):
    data = {"rios_user": json.dumps(USER), "rios_pops": json.dumps([])}
    if extra:
        data.update(extra)
    script = ""
    for k, v in data.items():
        script += f"window.localStorage.setItem({json.dumps(k)}, {json.dumps(v)});\n"
    page.add_init_script(script)

with sync_playwright() as p:
    browser = p.chromium.launch()

    # ── TESTE 1: RIOS_Config.html — salvar e recarregar logo/cor ────────────
    page = browser.new_page()
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    seed(page)
    page.goto(f"file://{BASE}/RIOS_Config.html")
    page.wait_for_timeout(200)
    check("Config: página carregou sem erros JS", len(errors) == 0, errors)
    page.fill("#cfg-color", "#123456")
    page.evaluate("""() => {
        _logoB64 = 'data:image/png;base64,AAAA';
        renderPreview();
    }""")
    page.click("text=💾 Salvar")
    page.wait_for_timeout(100)
    cfg = page.evaluate("JSON.parse(localStorage.getItem('rios_config_sistema'))")
    check("Config: salvou logoBase64 e corHex", cfg["logoBase64"] == "data:image/png;base64,AAAA" and cfg["corHex"] == "123456", cfg)

    # reload to confirm persisted values populate the form
    page2 = browser.new_page()
    seed(page2, extra={"rios_config_sistema": json.dumps(cfg)})
    page2.goto(f"file://{BASE}/RIOS_Config.html")
    page2.wait_for_timeout(200)
    color_val = page2.evaluate("document.getElementById('cfg-color').value")
    check("Config: cor recarregada do localStorage", color_val == "#123456", color_val)
    preview_html = page2.evaluate("document.getElementById('logo-preview').innerHTML")
    check("Config: preview mostra logo configurado ao recarregar", "Logo configurado" in preview_html)

    # ── TESTE 2: RIOS_POP.html — novo formulário (Materiais/Indicadores/PCC) ──
    page3 = browser.new_page()
    errors3 = []
    page3.on("pageerror", lambda e: errors3.append(str(e)))
    seed(page3)
    page3.goto(f"file://{BASE}/RIOS_POP.html")
    page3.wait_for_timeout(300)
    check("POP: página carregou sem erros JS", len(errors3) == 0, errors3)
    check("POP: uma linha inicial de Material foi criada", page3.locator("#mat-body tr").count() == 1)
    check("POP: uma linha inicial de Indicador foi criada", page3.locator("#ind-body tr").count() == 1)

    # preenche um material
    page3.fill("#mat-body .mt-grupo", "Ingredientes")
    page3.fill("#mat-body .mt-item", "Farinha — 500 g")
    page3.fill("#mat-body .mt-obs", "Marca X")
    # adiciona indicador
    page3.fill("#ind-body .in-controle", "Textura")
    page3.fill("#ind-body .in-criterio", "Lisa, sem grumos")
    # PCC com responsável
    page3.click("text=+ Adicionar PCC")
    page3.wait_for_timeout(50)
    pcc_row = page3.locator("#pcc-body tr").first
    pcc_row.locator(".pc-etapa").fill("Cocção")
    pcc_row.locator(".pc-risco").fill("Queima")
    pcc_row.locator(".pc-controle").fill("Fogo baixo")
    pcc_row.locator(".pc-resp").fill("Cozinheira")

    form = page3.evaluate("collectForm()")
    check("POP: materiais no novo formato (objeto)", form["materiais"][0] == {"grupo": "Ingredientes", "item": "Farinha — 500 g", "obs": "Marca X"}, form["materiais"])
    check("POP: indicadores no novo formato (objeto)", form["indicadores"][0] == {"controle": "Textura", "criterio": "Lisa, sem grumos"}, form["indicadores"])
    check("POP: pcc inclui responsavel", form["pcc"][0]["responsavel"] == "Cozinheira", form["pcc"])

    # ── TESTE 3: migração de POP salvo em formato antigo (materiais/indicadores como string[]) ──
    pop_antigo = {
        "id": "pop_legacy", "titulo": "POP Legado", "codigo": "POP-0009", "cliente": "Mestre do Pão",
        "clienteId": CLIENTE["id"], "setor": "Cozinha", "status": "ativo", "versao": "1.0",
        "revisoes": [{"numero": 0, "descricao": "Inicial", "revisadoPor": "", "aprovadoPor": "", "data": "2026-01-01"}],
        "objetivo": "x", "aplicacao": "x", "referencias": [], "materiais": ["Balança calibrada", "Luvas"],
        "definicoes": [], "responsaveis": [], "frequencia": "", "etapas": [],
        "pcc": [{"etapa": "E", "risco": "R", "controle": "C"}],
        "registros": [], "indicadores": ["Taxa de ruptura (%)"], "obs": "",
        "atualizadoEm": "2026-01-01T00:00:00.000Z"
    }
    page4 = browser.new_page()
    errors4 = []
    page4.on("pageerror", lambda e: errors4.append(str(e)))
    seed(page4, extra={"rios_pops": json.dumps([pop_antigo])})
    page4.goto(f"file://{BASE}/RIOS_POP.html")
    page4.wait_for_timeout(300)
    page4.click("text=📁 POPs Salvos") if page4.locator("text=📁 POPs Salvos").count() else page4.click("text=POPs Salvos")
    page4.wait_for_timeout(150)
    page4.click("text=POP Legado")
    page4.wait_for_timeout(200)
    check("POP legado: carregou sem quebrar (sem erros JS)", len(errors4) == 0, errors4)
    mat_item_val = page4.evaluate("document.querySelector('#mat-body .mt-item').value")
    check("POP legado: material antigo (string) migrado pro campo Item", mat_item_val == "Balança calibrada", mat_item_val)
    ind_criterio_val = page4.evaluate("document.querySelector('#ind-body .in-criterio').value")
    check("POP legado: indicador antigo (string) migrado pro campo Critério", ind_criterio_val == "Taxa de ruptura (%)", ind_criterio_val)
    pcc_resp_val = page4.evaluate("document.querySelector('#pcc-body .pc-resp').value")
    check("POP legado: PCC sem responsavel não quebra (campo vazio)", pcc_resp_val == "", repr(pcc_resp_val))

    # ── TESTE 4: fallback de branding — sem Papel Timbrado, usa Configurações ──
    page5 = browser.new_page()
    seed(page5, extra={"rios_config_sistema": json.dumps({"logoBase64": "data:image/png;base64,SISTEMA", "corHex": "123456"})})
    page5.goto(f"file://{BASE}/RIOS_POP.html")
    page5.wait_for_timeout(200)
    captured = page5.evaluate("""async () => {
        window.fetch = async (url, opts) => {
            window._capturedBody = JSON.parse(opts.body);
            return { ok: false, status: 599, json: async () => ({error: 'stub'}) };
        };
        await gerarWordDe({titulo: 'Teste Fallback', codigo: 'POP-0001'});
        return window._capturedBody;
    }""")
    check("POP: sem Papel Timbrado, usa logo/cor de Configurações no request", captured["logoBase64"] == "data:image/png;base64,SISTEMA" and captured["cliente"]["corHex"] == "123456", captured)

    # ── TESTE 5: Importar POP via Word — modal, review, aplicar no formulário ──
    page6 = browser.new_page()
    errors6 = []
    page6.on("pageerror", lambda e: errors6.append(str(e)))
    seed(page6)
    page6.goto(f"file://{BASE}/RIOS_POP.html")
    page6.wait_for_timeout(200)
    page6.click(".topbar >> text=Importar POP")
    page6.wait_for_timeout(100)
    check("Importar POP: modal abre", page6.locator("#mo-import-pop.open").count() == 1)
    check("Importar POP: input de arquivo .docx presente", page6.locator("#import-pop-file").count() == 1)

    fake_pop = {
        "titulo": "Recebimento de Peixe", "codigo": "POP-0099", "setor": "Recebimento / Compras", "versao": "2.0",
        "revisoes": [{"descricao": "Revisão inicial", "revisadoPor": "Chef", "aprovadoPor": "Marco", "data": "2026-08-31"}],
        "objetivo": "Padronizar o recebimento.", "aplicacao": "Estoque e cozinha.",
        "referencias": ["Manual BPM"],
        "materiais": [{"grupo": "Utensílios/equipamentos", "item": "Termômetro", "obs": "Calibrado"}],
        "definicoes": [{"termo": "Peixe fresco", "definicao": "Nunca congelado"}],
        "responsaveis": [{"papel": "Estoquista", "tarefas": ["Conferir temperatura", "Registrar lote"]}],
        "frequencia": "Diária",
        "etapas": [{"titulo": "Conferência", "destaque": "critico", "itens": ["Medir temperatura", "Checar odor"]}],
        "pcc": [{"etapa": "Conferência", "risco": "Quebra de cadeia de frio", "controle": "Termômetro a cada lote", "responsavel": "Estoquista"}],
        "registros": ["Temperatura", "Fornecedor"],
        "indicadores": [{"controle": "Temperatura", "criterio": "Até 4°C"}],
        "obs": "", "pendencias": [{"campo": "codigo", "descricao": "Código não encontrado no documento, inferido pelo padrão POP-00XX"}],
    }
    page6.evaluate("(data) => { window._importPopData = data; renderImportPopReview(data); }", fake_pop)
    page6.wait_for_timeout(100)
    check("Importar POP: pendência aparece na revisão", page6.locator("#pop-pend-0").count() == 1)
    check("Importar POP: setor inválido (não listado) não quebra a revisão", "Recebimento / Compras" in page6.locator("#import-pop-review-body").inner_text())

    page6.click("text=✓ Preencher Formulário")
    page6.wait_for_timeout(150)
    check("Importar POP: aplica sem erro JS", len(errors6) == 0, errors6)
    check("Importar POP: modal fecha após aplicar", page6.locator("#mo-import-pop.open").count() == 0)
    check("Importar POP: título aplicado no formulário", page6.evaluate("document.getElementById('f-titulo').value") == "Recebimento de Peixe")
    check("Importar POP: setor aplicado (whitelist bateu)", page6.evaluate("document.getElementById('f-setor').value") == "Recebimento / Compras")
    check("Importar POP: etapa importada com destaque crítico", page6.evaluate("document.querySelector('#etapas-body .etapa-card').className").find("dg-critico") != -1)
    check("Importar POP: material importado", page6.evaluate("document.querySelector('#mat-body .mt-item').value") == "Termômetro")
    check("Importar POP: PCC importado com responsável", page6.evaluate("document.querySelector('#pcc-body .pc-resp').value") == "Estoquista")
    obs_val = page6.evaluate("document.getElementById('f-obs').value")
    check("Importar POP: pendência registrada nas Observações", "A CONFIRMAR" in obs_val and "codigo" in obs_val, obs_val)

    form_after = page6.evaluate("collectForm()")
    check("Importar POP: collectForm() reflete os dados importados", form_after["titulo"] == "Recebimento de Peixe" and len(form_after["pcc"]) == 1 and form_after["pcc"][0]["responsavel"] == "Estoquista", form_after)

    browser.close()

print()
n_fail = sum(1 for _, ok, _ in results if not ok)
print(f"{'='*60}\n{len(results)} checagens, {n_fail} falha(s)")
sys.exit(1 if n_fail else 0)
