"""
tests/stress_test_haiku.py — Teste de estresse para validação do Haiku 4.5.

Envia uma especificação complexa multi-ambiente (6 ambientes, ~15 itens)
e avalia a qualidade da resposta do Haiku em:
  1. Tool use: chamou atualizar_planilha corretamente?
  2. Campos preenchidos: quantos campos ficaram como "(definir)"?
  3. Precisão: marca/modelo/dimensões estão razoáveis?
  4. Organização: itens foram para o ambiente correto?

Uso:
  python tests/stress_test_haiku.py
"""

import sys
import os
import json
import time

# Adicionar raiz do projeto ao path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import anthropic
import toml

from core.system_prompt import SYSTEM_PROMPT, TOOL_DEFINITION, WEB_SEARCH_TOOL
from core.spreadsheet import create_template, load_workbook_from_bytes, get_sheet_data, get_compact_context
from core.claude_client import MODEL, CACHED_SYSTEM, TOOLS, MAX_TOKENS_MAIN


# ─── Config ──────────────────────────────────────────────────────────────────

def load_api_key() -> str:
    """Carrega API key do secrets.toml do Streamlit."""
    secrets_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        ".streamlit", "secrets.toml"
    )
    secrets = toml.load(secrets_path)
    return secrets["ANTHROPIC_API_KEY"]


# ─── Especificação de estresse ───────────────────────────────────────────────

# Cenário realista: apartamento 120m², 6 ambientes, ~15 itens com mix de
# produtos específicos (marca+modelo) e genéricos (só descrição)
STRESS_PROMPT = """
Projeto: Apartamento Vila Madalena 120m² - Cliente Ana Beatriz

Especificação completa dos ambientes:

COZINHA:
- cooktop de indução 5 bocas Electrolux IC80, preto
- coifa de ilha Tramontina Vetro 90cm inox
- cuba dupla Franke Planar PEX 620 inox escovado
- torneira gourmet com filtro Docol Vitalis bica alta, cromada
- geladeira French Door Brastemp BRO85AK inox, 554L

LAVABO:
- cuba de apoio Deca L.1038 marfim, formato oval
- torneira de mesa Deca Polo bica baixa, cromada
- papeleira de parede, acabamento rose gold

BANHEIRO MASTER:
- bacia com caixa acoplada Deca Vogue Plus P.460, branca
- ducha higiênica Docol Riva, cromada
- chuveiro de teto 30x30cm Deca Quadratta, cromado
- nicho embutido 60x30cm para produtos, acabamento em porcelanato

SUÍTE CASAL:
- 2 arandelas de leitura Reka Lumina braço articulado, dourado escovado
- 2 interruptores inteligentes Sonoff M5 3 canais, branco
- persiana motorizada Hunter Douglas Duette, blackout cinza

SALA DE ESTAR:
- pendente central Lumini Bossa 800mm, preto fosco
- trilho eletrificado 3m com 4 spots direcionáveis, preto
- 3 tomadas USB embutidas Tramontina Liz, brancas

VARANDA GOURMET:
- churrasqueira a gás Fischer Debret 4 queimadores, inox
- pia de apoio granito preto 120x55cm com cuba inox 40x34
- pendente externo à prova d'água IP65, 2 unidades
"""

# Campos obrigatórios que devem ser preenchidos em cada item
REQUIRED_FIELDS = ["item", "modelo", "segmento", "acabamento", "marca", "dimensoes", "quantidade"]

# Ambientes esperados
EXPECTED_ENVIRONMENTS = {"COZINHA", "LAVABO", "BANHEIRO MASTER", "SUÍTE CASAL", "SALA DE ESTAR", "VARANDA GOURMET"}

# Contagem mínima de itens esperada por ambiente
EXPECTED_MIN_ITEMS = {
    "COZINHA": 5,
    "LAVABO": 3,
    "BANHEIRO MASTER": 4,
    "SUÍTE CASAL": 3,
    "SALA DE ESTAR": 3,
    "VARANDA GOURMET": 3,
}


# ─── Teste ───────────────────────────────────────────────────────────────────

def run_stress_test():
    print(f"{'='*70}")
    print(f"TESTE DE ESTRESSE — Arqui Specs com {MODEL}")
    print(f"{'='*70}\n")

    # 1. Setup
    api_key = load_api_key()
    client = anthropic.Anthropic(api_key=api_key)
    template_buf = create_template(project_name="Apto Vila Madalena 120m²")
    wb_bytes = template_buf.read()
    wb = load_workbook_from_bytes(wb_bytes)

    context = get_compact_context(wb)
    user_message = f"{context}\n\n---\n{STRESS_PROMPT}"

    print(f"Modelo: {MODEL}")
    print(f"Ambientes na especificação: {len(EXPECTED_ENVIRONMENTS)}")
    print(f"Itens esperados: ~{sum(EXPECTED_MIN_ITEMS.values())}")
    print(f"\nEnviando especificação complexa...\n")

    # 2. Chamada à API (sem streaming para simplificar o teste)
    t0 = time.time()

    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS_MAIN,
        system=CACHED_SYSTEM,
        tools=TOOLS,
        messages=[{"role": "user", "content": user_message}],
    )

    t1 = time.time()
    elapsed = t1 - t0

    # 3. Análise da resposta
    print(f"{'─'*70}")
    print(f"RESPOSTA RECEBIDA ({elapsed:.1f}s)")
    print(f"{'─'*70}\n")

    # Contabilizar usage
    usage = response.usage
    input_cost = (usage.input_tokens / 1_000_000) * 1.00   # Haiku: $1/MTok input
    output_cost = (usage.output_tokens / 1_000_000) * 5.00  # Haiku: $5/MTok output
    cache_read = getattr(usage, 'cache_read_input_tokens', 0) or 0
    cache_create = getattr(usage, 'cache_creation_input_tokens', 0) or 0

    print(f"Tokens — input: {usage.input_tokens}, output: {usage.output_tokens}")
    print(f"Cache — read: {cache_read}, creation: {cache_create}")
    print(f"Custo estimado: ${input_cost + output_cost:.4f}")
    print(f"Stop reason: {response.stop_reason}")
    print()

    # 4. Verificar tool use
    tool_calls = [b for b in response.content if b.type == "tool_use"]
    text_blocks = [b for b in response.content if b.type == "text"]

    print(f"{'─'*70}")
    print(f"ANÁLISE DO TOOL USE")
    print(f"{'─'*70}\n")

    if not tool_calls:
        print("❌ FALHA: Claude NÃO chamou atualizar_planilha!")
        print("   Texto da resposta:")
        for tb in text_blocks:
            print(f"   {tb.text[:500]}")
        return

    # Extrair operações de todas as chamadas tool_use
    all_operations = []
    for tc in tool_calls:
        if tc.name == "atualizar_planilha":
            ops = tc.input.get("operacoes", [])
            all_operations.extend(ops)
            print(f"✅ atualizar_planilha chamado com {len(ops)} operação(ões)")

    total_ops = len(all_operations)
    print(f"\nTotal de operações: {total_ops}")

    if text_blocks:
        print(f"\nTexto do assistente:")
        for tb in text_blocks:
            print(f"  {tb.text[:300]}...")

    # 5. Análise detalhada das operações
    print(f"\n{'─'*70}")
    print(f"ANÁLISE DAS OPERAÇÕES ({total_ops} itens)")
    print(f"{'─'*70}\n")

    # Agrupar por ambiente
    by_env: dict[str, list[dict]] = {}
    for op in all_operations:
        env = (op.get("ambiente") or "SEM_AMBIENTE").upper()
        by_env.setdefault(env, []).append(op)

    # 5a. Verificar ambientes
    found_envs = set(by_env.keys())
    missing_envs = EXPECTED_ENVIRONMENTS - found_envs
    extra_envs = found_envs - EXPECTED_ENVIRONMENTS

    print("Ambientes encontrados:")
    for env in sorted(found_envs):
        count = len(by_env[env])
        expected = EXPECTED_MIN_ITEMS.get(env, "?")
        status = "✅" if count >= (EXPECTED_MIN_ITEMS.get(env, 0)) else "⚠️"
        print(f"  {status} {env}: {count} itens (esperado ≥{expected})")

    if missing_envs:
        print(f"\n❌ Ambientes faltando: {', '.join(missing_envs)}")
    if extra_envs:
        print(f"\n⚠️  Ambientes extras: {', '.join(extra_envs)}")

    # 5b. Verificar preenchimento de campos
    print(f"\n{'─'*70}")
    print(f"QUALIDADE DOS CAMPOS")
    print(f"{'─'*70}\n")

    total_fields = 0
    definir_count = 0
    empty_count = 0
    filled_count = 0

    for op in all_operations:
        for field in REQUIRED_FIELDS:
            total_fields += 1
            value = op.get(field, "")
            if not value or value.strip() == "":
                empty_count += 1
            elif "(definir)" in str(value).lower():
                definir_count += 1
            else:
                filled_count += 1

    fill_rate = (filled_count / total_fields * 100) if total_fields > 0 else 0

    print(f"Total de campos analisados: {total_fields}")
    print(f"  ✅ Preenchidos:     {filled_count} ({filled_count/total_fields*100:.0f}%)")
    print(f"  ⚠️  (definir):       {definir_count} ({definir_count/total_fields*100:.0f}%)")
    print(f"  ❌ Vazios:          {empty_count} ({empty_count/total_fields*100:.0f}%)")
    print(f"\n  Taxa de preenchimento: {fill_rate:.0f}%")

    # 5c. Detalhe de cada item
    print(f"\n{'─'*70}")
    print(f"DETALHE POR ITEM")
    print(f"{'─'*70}\n")

    for env in sorted(by_env.keys()):
        print(f"\n  ── {env} ──")
        for op in by_env[env]:
            item_name = op.get("item", "(sem nome)")
            modelo = op.get("modelo", "")
            marca = op.get("marca", "")
            dims = op.get("dimensoes", "")
            seg = op.get("segmento", "")
            acab = op.get("acabamento", "")
            qtd = op.get("quantidade", "")

            # Flag para campos problemáticos
            issues = []
            if not modelo or "(definir)" in str(modelo).lower():
                issues.append("modelo")
            if not marca or "(definir)" in str(marca).lower():
                issues.append("marca")
            if not dims or "(definir)" in str(dims).lower():
                issues.append("dimensões")

            status = "✅" if not issues else f"⚠️  [{', '.join(issues)}]"
            print(f"    {status} {item_name} | {marca} {modelo} | {seg} | {acab} | {dims} | qty:{qtd}")

    # 6. Resultado final
    print(f"\n{'='*70}")
    print(f"RESUMO DO TESTE DE ESTRESSE")
    print(f"{'='*70}\n")

    score = 0
    max_score = 5

    # Critério 1: Tool use
    if tool_calls:
        score += 1
        print(f"  ✅ [1/5] Tool use: chamou atualizar_planilha")
    else:
        print(f"  ❌ [1/5] Tool use: NÃO chamou a ferramenta")

    # Critério 2: Ambientes corretos
    env_coverage = len(found_envs & EXPECTED_ENVIRONMENTS) / len(EXPECTED_ENVIRONMENTS)
    if env_coverage >= 0.9:
        score += 1
        print(f"  ✅ [2/5] Ambientes: {env_coverage*100:.0f}% cobertos")
    else:
        print(f"  ⚠️  [2/5] Ambientes: {env_coverage*100:.0f}% cobertos (esperado ≥90%)")

    # Critério 3: Contagem de itens
    total_expected = sum(EXPECTED_MIN_ITEMS.values())
    if total_ops >= total_expected * 0.8:
        score += 1
        print(f"  ✅ [3/5] Itens: {total_ops}/{total_expected} ({total_ops/total_expected*100:.0f}%)")
    else:
        print(f"  ⚠️  [3/5] Itens: {total_ops}/{total_expected} ({total_ops/total_expected*100:.0f}%)")

    # Critério 4: Taxa de preenchimento
    if fill_rate >= 70:
        score += 1
        print(f"  ✅ [4/5] Preenchimento: {fill_rate:.0f}% (esperado ≥70%)")
    else:
        print(f"  ⚠️  [4/5] Preenchimento: {fill_rate:.0f}% (esperado ≥70%)")

    # Critério 5: Poucos campos vazios
    if empty_count <= total_fields * 0.05:
        score += 1
        print(f"  ✅ [5/5] Campos vazios: {empty_count} (≤5% do total)")
    else:
        print(f"  ⚠️  [5/5] Campos vazios: {empty_count} (>{5}% do total)")

    print(f"\n  SCORE FINAL: {score}/{max_score}")
    print(f"  Modelo: {MODEL}")
    print(f"  Tempo: {elapsed:.1f}s")
    print(f"  Custo: ${input_cost + output_cost:.4f}")

    if score >= 4:
        print(f"\n  ✅ APROVADO — Haiku 4.5 é viável para este caso de uso.")
    elif score >= 3:
        print(f"\n  ⚠️  PARCIAL — Haiku 4.5 funciona, mas com perdas de qualidade.")
    else:
        print(f"\n  ❌ REPROVADO — Considere manter Sonnet 4.6.")

    print(f"\n{'='*70}")


if __name__ == "__main__":
    run_stress_test()
