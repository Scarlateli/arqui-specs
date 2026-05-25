"""
core/claude_client.py — Integração com a API da Anthropic para o Arqui Specs.

Responsabilidades:
- Inicializar cliente Anthropic com a API key
- Gerenciar o loop de streaming + tool use (SCRUM-49, 51, 60)
- Aplicar prompt caching no system prompt (SCRUM-49)
- Habilitar web search para specs reais (SCRUM-53)
- Formatar mensagens para a API (histórico de conversa)
- Tratar erros de API com mensagens amigáveis (SCRUM-55)

Fluxo de uma mensagem:
  1. Claude recebe mensagem do usuário + estado da planilha
  2. Claude decide: responder com texto OU chamar atualizar_planilha
  3. Se tool use: executamos a ferramenta, devolvemos resultado, Claude confirma
  4. Streamlit exibe texto em tempo real via generator

Sobre prompt caching:
  O system prompt é estável (sem datas, IDs ou dados variáveis).
  Anthropic cacheia o prefixo após ~1024 tokens estáveis.
  Economia esperada: 80-90% dos tokens de input (system prompt é ~1000 tokens).
"""

import anthropic
import openpyxl
from typing import Generator

from .system_prompt import SYSTEM_PROMPT, TOOL_DEFINITION, WEB_SEARCH_TOOL
from .spreadsheet import apply_operations, get_compact_context, save_workbook_to_bytes, load_workbook_from_bytes

# ─── Constantes ──────────────────────────────────────────────────────────────

# claude-haiku-4-5: otimizado para custo no MVP.
# ADR-003: sonnet-4-5 → sonnet-4-6 → haiku-4-5 (redução de custo ~3x, contexto 200K).
MODEL = "claude-haiku-4-5"

MAX_TOKENS_MAIN = 4096    # Para respostas de chat (suficiente para qualquer resposta + tool use)
MAX_TOKENS_FOLLOWUP = 2048  # Para confirmação pós-tool use

# System prompt com cache_control — Anthropic cacheia este prefixo estável
# (economiza 80-90% dos tokens de input após a primeira requisição)
CACHED_SYSTEM = [
    {
        "type": "text",
        "text": SYSTEM_PROMPT,
        "cache_control": {"type": "ephemeral"},  # TTL padrão: 5 minutos
    }
]

# Ferramentas disponíveis para o Claude
TOOLS = [
    TOOL_DEFINITION,  # atualizar_planilha — ferramenta principal
    WEB_SEARCH_TOOL,  # web_search — buscar specs reais de produtos
]


# ─── Inicialização do cliente ────────────────────────────────────────────────

def get_client(api_key: str) -> anthropic.Anthropic:
    """Cria e retorna um cliente Anthropic autenticado."""
    return anthropic.Anthropic(api_key=api_key)


# ─── Formatação de mensagens ─────────────────────────────────────────────────

def build_user_message(user_text: str, workbook: openpyxl.Workbook) -> dict:
    """
    Constrói a mensagem do usuário com contexto compacto da planilha.

    Incluímos o estado da planilha em cada mensagem para que Claude saiba
    o que já foi adicionado, sem precisar buscar via ferramenta.
    """
    context = get_compact_context(workbook)
    content = f"{context}\n\n---\n{user_text}"
    return {"role": "user", "content": content}


# ─── Loop de streaming com tool use ─────────────────────────────────────────

def stream_response(
    api_key: str,
    api_messages: list[dict],
    workbook: openpyxl.Workbook,
) -> Generator[tuple[str, openpyxl.Workbook | None, str | None], None, None]:
    """
    Generator que processa uma mensagem e faz streaming da resposta do Claude.

    Yield: (chunk_text, updated_workbook_or_None, error_or_None)
    - chunk_text: fragmento de texto para exibir em tempo real
    - updated_workbook_or_None: workbook atualizado se houve tool use, None caso contrário
    - error_or_None: mensagem de erro amigável se algo deu errado, None se ok

    Fluxo interno:
    1. Chama Claude com streaming
    2. Se stop_reason == "tool_use" → executa ferramenta → chama Claude novamente
    3. Yield de cada chunk de texto nas duas fases
    """
    client = get_client(api_key)
    updated_wb = None

    try:
        # ── Fase 1: Chamada principal com streaming ──────────────────────────
        phase1_text = ""
        phase1_content = []  # Conteúdo completo para o histórico de API

        with client.messages.stream(
            model=MODEL,
            max_tokens=MAX_TOKENS_MAIN,
            system=CACHED_SYSTEM,
            tools=TOOLS,
            messages=api_messages,
        ) as stream:
            for text in stream.text_stream:
                phase1_text += text
                yield (text, None, None)

            final_msg = stream.get_final_message()
            phase1_content = final_msg.content

        # ── Verificar se houve tool use ──────────────────────────────────────
        if final_msg.stop_reason == "tool_use":
            # Sinaliza que está processando a ferramenta
            yield ("\n\n*Atualizando planilha...*\n\n", None, None)

            tool_result_blocks = []

            for block in final_msg.content:
                if block.type == "tool_use" and block.name == "atualizar_planilha":
                    operacoes = block.input.get("operacoes", [])
                    result = apply_operations(workbook, operacoes)
                    updated_wb = workbook

                    # Feedback para o Claude sobre o resultado da ferramenta
                    result_text = (
                        f"Executado: {result['items_adicionados']} item(s) adicionado(s), "
                        f"{result['items_atualizados']} item(s) atualizado(s). "
                        f"Total na planilha: {result['total_itens']} item(s)."
                    )
                    if result["erros"]:
                        result_text += f" Erros: {'; '.join(result['erros'])}"

                    tool_result_blocks.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": result_text,
                    })

                elif block.type == "tool_use" and block.name == "web_search":
                    # Web search é executado automaticamente pela Anthropic — não precisamos fazer nada
                    # O resultado virá na resposta subsequente
                    pass

            if tool_result_blocks:
                # ── Fase 2: Confirmação do Claude pós-tool use ───────────────
                followup_messages = api_messages + [
                    {"role": "assistant", "content": phase1_content},
                    {"role": "user", "content": tool_result_blocks},
                ]

                with client.messages.stream(
                    model=MODEL,
                    max_tokens=MAX_TOKENS_FOLLOWUP,
                    system=CACHED_SYSTEM,
                    messages=followup_messages,
                ) as stream2:
                    for text in stream2.text_stream:
                        yield (text, None, None)

                    final_msg2 = stream2.get_final_message()

                # Devolver workbook atualizado no último chunk
                yield ("", updated_wb, None)

                # Adicionar mensagens da fase 2 ao histórico
                # (o app.py deve capturar isso para atualizar api_messages)
                # Aqui retornamos via atributo adicional — solução: usar um dict mutável
                # Por simplicidade, o app.py vai reconstruir o histórico após a função

        else:
            # Sem tool use — apenas texto
            yield ("", None, None)

    except anthropic.AuthenticationError:
        yield ("", None, "❌ **Chave de API inválida.** Verifique `ANTHROPIC_API_KEY` em `.streamlit/secrets.toml`.")

    except anthropic.RateLimitError:
        yield ("", None, "⏱️ **Limite de requisições atingido.** Aguarde alguns segundos e tente novamente.")

    except anthropic.APIConnectionError:
        yield ("", None, "🌐 **Erro de conexão.** Verifique sua internet e tente novamente.")

    except anthropic.BadRequestError as e:
        yield ("", None, f"⚠️ **Erro na requisição:** {str(e)[:200]}")

    except Exception as e:
        yield ("", None, f"⚠️ **Erro inesperado:** {str(e)[:200]}")


def stream_response_collecting(
    api_key: str,
    api_messages: list[dict],
    workbook: openpyxl.Workbook,
) -> tuple[str, openpyxl.Workbook | None, list[dict], str | None]:
    """
    Versão não-generator de stream_response para uso em contextos síncronos.
    Retorna (texto_completo, workbook_atualizado, novas_mensagens_api, erro).

    O workbook retornado é o mesmo objeto passado (modificado in-place) ou None.
    """
    full_text = ""
    updated_wb = None
    error = None

    for chunk, wb_update, err in stream_response(api_key, api_messages, workbook):
        full_text += chunk
        if wb_update:
            updated_wb = wb_update
        if err:
            error = err

    return full_text, updated_wb, error


# ─── Helpers de histórico ────────────────────────────────────────────────────

def build_api_messages_from_history(
    display_messages: list[dict],
    workbook: openpyxl.Workbook,
) -> list[dict]:
    """
    Reconstrói a lista de mensagens para a API a partir do histórico de display.

    Nota: Para conversas longas, poderíamos truncar o histórico aqui.
    Para o MVP, mantemos tudo — a janela de contexto do Haiku 4.5 é 200K tokens,
    mais do que suficiente para qualquer projeto de memorial descritivo.
    """
    messages = []
    context = get_compact_context(workbook)

    for i, msg in enumerate(display_messages):
        if msg["role"] == "user":
            # Na primeira mensagem, incluir contexto; nas demais, incluir contexto atual
            if i == 0:
                content = f"{context}\n\n---\n{msg['content']}"
            else:
                content = msg["content"]
            messages.append({"role": "user", "content": content})
        else:
            messages.append({"role": "assistant", "content": msg["content"]})

    return messages
