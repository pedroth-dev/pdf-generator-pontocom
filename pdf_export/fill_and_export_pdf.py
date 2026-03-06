"""
Preenche o modelo Excel (Cobertura Premium) com os dados recebidos
e exporta para PDF usando o Microsoft Excel (xlwings).

Os campos são localizados por BUSCA no Excel (texto exato), não por posição fixa,
para que o modelo possa variar de layout.

Requisitos para rodar este script: Python 3, Microsoft Excel (Windows), pip install xlwings
Para o usuário final: use o .exe gerado por PyInstaller (não precisa instalar Python).

Uso:
  fill_and_export_pdf.exe --template "modelo.xlsx" --data "dados.json" --output "saida.pdf"
  ou: python fill_and_export_pdf.py --template ... --data ... --output ...
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

# Log provisório: arquivo em %TEMP% para inspeção após gerar o PDF
LOG_PATH = Path(os.environ.get("TEMP", os.path.expanduser("~"))) / "cobertura_pdf_export_log.txt"


def _log(msg: str) -> None:
    line = f"[PDF Export] {msg}\n"
    sys.stderr.write(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass

# Constantes para busca no Excel: texto a localizar -> chave no JSON
# Substitui o conteúdo INTEIRO da célula pelo valor formatado (ex.: valor monetário)
# (custoDeslocamento não é mais escrito em célula; permanece apenas no cálculo do Valor Total)
FIELD_SEARCH = {}

# Placeholders: substituir apenas o texto do placeholder DENTRO do texto da célula.
# Ex.: célula "A cobertura é [Tipo de Cobertura]" -> "[Tipo de Cobertura]" vira "ACM", resultado "A cobertura é ACM"
FIELD_PLACEHOLDER_REPLACE = {
    "[Tipo de Cobertura]": "tipoCobertura",
    "[Medida do Pilar]": "medidaPilar",
    "[Telha Térmica]": "telhaTermica",
    "[Forro PVC]": "forroPvc",
    # Cliente (nomeCliente, cpfCnpj, endereco, celularFone, cidade; dataAtual definida na geração)
    "[Nome do Cliente]": "nomeCliente",
    "[CPF/CPNJ]": "cpfCnpj",
    "[Endereço]": "endereco",
    "[Celular/Fone]": "celularFone",
    "[Data Atual]": "dataAtual",
    "[Cidade]": "cidade",
    "[Valor p/ Forma de Pagamento]": "valorFormaPagamento",
}

# Placeholders para a planilha Pergolado (orçamento + mesmos de cliente/valores)
FIELD_PLACEHOLDER_REPLACE_PERGOLADO = {
    "[Medidas do Pergolado]": "medidas",
    "[Dimensão do Tubo Retangular]": "dimensaoTubo",
    "[Tipo do Policarbonato]": "tipoPolicarbonato",
    "[Cor]": "corPolicarbonato",
    "[Nome do Cliente]": "nomeCliente",
    "[CPF/CPNJ]": "cpfCnpj",
    "[Endereço]": "endereco",
    "[Celular/Fone]": "celularFone",
    "[Data Atual]": "dataAtual",
    "[Cidade]": "cidade",
    "[Valor p/ Forma de Pagamento]": "valorFormaPagamento",
}

# Placeholders para a planilha Cobertura Retrátil (apenas dados do cliente e valor; medidas vão na descrição D43)
FIELD_PLACEHOLDER_REPLACE_COBERTURA_RETRATIL = {
    "[Nome do Cliente]": "nomeCliente",
    "[CPF/CPNJ]": "cpfCnpj",
    "[Endereço]": "endereco",
    "[Celular/Fone]": "celularFone",
    "[Data Atual]": "dataAtual",
    "[Cidade]": "cidade",
    "[Valor p/ Forma de Pagamento]": "valorFormaPagamento",
}

# Placeholders para a planilha Porta (dados do cliente e valor; descrição montada em código em D43)
FIELD_PLACEHOLDER_REPLACE_PORTA = {
    "[Nome do Cliente]": "nomeCliente",
    "[CPF/CPNJ]": "cpfCnpj",
    "[Endereço]": "endereco",
    "[Celular/Fone]": "celularFone",
    "[Data Atual]": "dataAtual",
    "[Cidade]": "cidade",
    "[Valor p/ Forma de Pagamento]": "valorFormaPagamento",
}

# Campos que recebem o valor parcelado em 10x (total + 10%); todas as células com esse texto são preenchidas
FIELD_TOTAL_LABEL = "[Valor Total]"
# Cobertura Retrátil: valor total geral (cobertura + automatização) à vista
FIELD_TOTAL_GERAL_LABEL = "[Valor Total Geral]"

# Célula fixa que recebe o texto de especificação (sempre D43, primeira planilha)
D43_CELL = "D43"
# Células fixas Cobertura Retrátil modo automatizado (D44 = descrição automatizador; M44 e N44 = valor da abertura automatizada)
D44_CELL = "D44"
M44_CELL = "M44"
N44_CELL = "N44"

# Template do texto de especificação. Placeholders: [Medidas da Cobertura], [Tipo de Cobertura], [Pintura da Cobertura], [Espessura da Telha Térmica], [Tipo do Forro].
# Regra Item 3 (situacional): se temPilar != "Sim", remove o bloco "Item 3: Pilar metálico..." e renumera 4->3, 5->4, 6->5, 7->6.
TEXTO_ESPECIFICACAO_TEMPLATE = """Cobertura Premium medidas [Medidas da Cobertura]

Item 1: Treliça metálica com 40 cm de altura. 
Detalhamento de fabricação: Banzos superior e inferior em perfil U simples 75x40 #14, montantes e diagonais em perfil U simples 68x30 #14. A treliça contorna toda estrutura sendo o objeto principal de estruturação da cobertura.

Item 2: Revestimento das treliças em [Tipo de Cobertura] [Pintura da Cobertura].

Item 3: Pilar metálico de 100x100 #14.

Item 4: Vigas metálicas e terças metálicas em metalon 50 x 50 #18. Esse item está locado na parte interna da cobertura para receber telhas térmicas e calha.

Item 5: Telha térmica EPS de [Espessura da Telha Térmica] com acabamento em filme para dar resistência na instalação e não ocorrer o desplacamento do EPS.

Item 6: Calhas e rufos galvanizados afim de garantir a vedação por completo do telhado e escoamento da água.

Item 7: Forro PVC [Tipo do Forro] amadeirado nivelado na parte de baixo da cobertura."""

# Mapeamento placeholder -> chave no JSON para o texto de D43
D43_PLACEHOLDERS = {
    "[Medidas da Cobertura]": "medidas",
    "[Tipo de Cobertura]": "tipoCobertura",
    "[Pintura da Cobertura]": "corOuPintura",
    "[Espessura da Telha Térmica]": "telhaTermica",
    "[Tipo do Forro]": "forroPvc",
}

# Bloco do Item 3 (pilar); removido quando temPilar != "Sim". \n explícitos para garantir match.
ITEM_3_BLOCO = "\n\nItem 3: Pilar metálico de 100x100 #14.\n\n"


def build_texto_especificacao_d43(data: dict) -> str:
    """
    Monta o texto de especificação com placeholders substituídos.
    Aplica a regra do Item 3: se não tiver pilar, remove o bloco do Item 3 e renumera 4->3, 5->4, 6->5, 7->6.
    """
    text = TEXTO_ESPECIFICACAO_TEMPLATE
    for placeholder, json_key in D43_PLACEHOLDERS.items():
        value = data.get(json_key) or ""
        text = text.replace(placeholder, str(value))

    tem_pilar = (data.get("temPilar") or "").strip() == "Sim"
    if not tem_pilar:
        # Remove só a linha do pilar, mantendo \n\n entre Item 2 e o que vira Item 3 (evita "Preto.Item 4:" grudado)
        bloco_completo = "\n\nItem 3: Pilar metálico de 100x100 #14.\n\n"
        substitui_por = "\n\n"  # preserva parágrafo entre Item 2 e Item 3
        if bloco_completo in text:
            text = text.replace(bloco_completo, substitui_por)
        else:
            text = text.replace("\n\nItem 3: Pilar metálico de 100x100 #14.\n\n", substitui_por)
        # Renumera em uma única passada: Item 4->3, 5->4, 6->5, 7->6
        def _renum(match):
            n = int(match.group(1))
            return "\nItem {}:".format(n - 1)
        text = re.sub(r"\nItem ([4-7]):", _renum, text)

    return text


def build_texto_especificacao_d43_retratil(data: dict) -> str:
    """
    Monta o texto de especificação D43 para Cobertura Retrátil.
    - Telha Térmica: template com cor parte superior/inferior; linha do modo de abertura só se NÃO Automatizada.
    - Policarbonato: template com material; linha do modo de abertura só se NÃO Automatizada.
    - Evita duplicar a palavra "medidas" quando data["medidas"] já vem com prefixo "medidas ".
    """
    tipo = (data.get("tipoCobertura") or "").strip()
    medidas_raw = (data.get("medidas") or "").strip()
    # Evitar "medidas: medidas ..." na linha 1: usar só o restante se já vier com prefixo "medidas "
    if medidas_raw.lower().startswith("medidas "):
        medidas_display = medidas_raw[8:].strip()
    else:
        medidas_display = medidas_raw
    modo_abertura = (data.get("modoAbertura") or "").strip()
    is_automatizada = modo_abertura == "Automatizada"
    # Frase do modo de abertura só quando NÃO for Automatizada (retirar quando for Automatizada).
    sufixo_modo = "\n\nCobertura com modo de abertura [Modo de Abertura]\n" if not is_automatizada else ""

    if tipo == "Telha Térmica":
        template = (
            "Cobertura Metálica Retrátil, medidas: [Medidas]\n\n"
            "Cobertura metálica retrátil sendo uma folha de abrir e outra fixa com telha isotérmica sendo aço/aço 50mm, "
            "acabamento [Cor da Parte Superior] na parte superior da telha e acabamento em aço [Cor da Parte Inferior] na parte inferior da telha, "
            "tendo calha e rufo. Acabamento na parte metálica sendo pintura automotiva cor preto fosco."
        )
        cor_superior = (data.get("corParteSuperior") or "").strip().lower()
        cor_inferior = (data.get("corParteInferior") or "").strip().lower()
        texto = (
            template.replace("[Medidas]", medidas_display)
            .replace("[Cor da Parte Superior]", cor_superior)
            .replace("[Cor da Parte Inferior]", cor_inferior)
        )
        if not is_automatizada:
            texto += sufixo_modo.replace("[Modo de Abertura]", modo_abertura.strip().lower())
        return texto

    # Policarbonato Compacto 3mm ou Alveolar 6mm
    material = (
        "policarbonato alveolar 6mm"
        if tipo == "Policarbonato Alveolar 6mm"
        else "policarbonato compacto 3mm"
    )
    template = (
        "Cobertura Metálica Retrátil, medidas: [Medidas]\n\n"
        "Cobertura metálica retrátil sendo uma folha de abrir e outra fixa com [material], "
        "tendo calha e rufo. Acabamento na parte metálica sendo pintura automotiva cor preto fosco."
    )
    texto = template.replace("[Medidas]", medidas_display).replace("[material]", material)
    if not is_automatizada:
        texto += sufixo_modo.replace("[Modo de Abertura]", modo_abertura.strip().lower())
    return texto


def build_texto_especificacao_d43_porta(data: dict) -> str:
    """
    Monta o texto de especificação D43 para Porta.
    Todos os valores na descrição em minúsculas.
    Linhas condicionais: Bandeirola e Alizar só se selecionados.
    Para Ferro Forjado e Aço Corten, o trecho [Acondicionamento] é omitido da linha da porta.
    """
    def _(s: str) -> str:
        return (s or "").strip().lower()

    modelo = _(data.get("modeloPorta") or "")
    modo_puxador = _(data.get("modoPuxador") or "")
    medidas_geral = (data.get("medidasPortaGeral") or "").strip()
    medidas_porta = (data.get("medidasPorta") or "").strip()
    sistema = _(data.get("sistemaAbertura") or "")
    estilo = _(data.get("estiloFolha") or "")
    acond = (data.get("acondicionamentoEfetivo") or "").strip().lower()
    espessura = (data.get("espessuraChapa") or "").strip().lower()
    pintura = (data.get("corPintura") or "").strip().lower()
    modo_entrega = _(data.get("modoEntrega") or "")

    # Linha 1: Porta modelo [Modelo]. Medida total: [Medidas geral].
    linha1 = f"Porta modelo {modelo}. Medida total: {medidas_geral}."

    # Linha 2: Porta: [Modo Puxador], [Sistema], [Estilo] [, Acondicionamento]. Fabricada em chapa [Espessura]. [Medidas porta].
    # Para Ferro Forjado e Aço Corten: sem [Acondicionamento].
    modelo_raw = (data.get("modeloPorta") or "").strip()
    sem_acondicionamento_na_linha = modelo_raw in ("Ferro Forjado", "Aço Corten")
    if sem_acondicionamento_na_linha:
        linha2 = f"Porta: {modo_puxador}, {sistema}, {estilo}. Fabricada em chapa {espessura}. {medidas_porta}."
    else:
        parte_acond = f", {acond}" if acond else ""
        linha2 = f"Porta: {modo_puxador}, {sistema}, {estilo}{parte_acond}. Fabricada em chapa {espessura}. {medidas_porta}."

    linhas = [linha1, "", linha2]

    if data.get("bandeirola"):
        med_band = (data.get("medidasBandeirola") or "").strip()
        linhas.append("")
        linhas.append(f"Bandeirola: {med_band}.")

    if data.get("alizar"):
        med_alizar = (data.get("medidaAlizar") or "").strip()
        linhas.append("")
        linhas.append(f"Alizar: {med_alizar}, em dobra especial.")

    linhas.append("")
    linhas.append(f"Acabamento: {pintura}, {modo_entrega}.")
    linhas.append("")
    linhas.append("Incluso: fechadura rolete ou maçaneta simples.")
    linhas.append("Não incluso: vidro, puxadores especiais e fechadura eletrônica.")

    return "\n".join(linhas)


def format_currency(raw: str) -> str:
    """Converte dígitos (ex: '150000') em 'R$ 1.500,00'."""
    digits = "".join(c for c in (raw or "") if c.isdigit())
    if not digits:
        return "R$ 0,00"
    cents = digits[-2:].rjust(2, "0")
    int_part = digits[:-2] or "0"
    if len(int_part) > 3:
        parts = []
        while int_part:
            parts.append(int_part[-3:])
            int_part = int_part[:-3]
        int_part = ".".join(reversed(parts))
    return f"R$ {int_part},{cents}"


def raw_to_reais(raw: str) -> float:
    """Converte string de dígitos (centavos) em valor em reais. Ex: '150000' -> 1500.00"""
    digits = "".join(c for c in (raw or "") if c.isdigit())
    if not digits:
        return 0.0
    return int(digits) / 100.0


def parse_medidas_m2(medidas: str) -> float:
    """Extrai as duas dimensões de '5,00m x 2,00m' e retorna m² (ex: 5 * 2 = 10.0)."""
    s = (medidas or "").strip()
    # Aceita "5,00m x 2,00m" ou "5.00 x 2.00" (opcional: m ou m² entre número e x)
    m = re.search(r"(\d+[,.]?\d*)\s*m?\s*[xX×]\s*(\d+[,.]?\d*)\s*m?", s, re.IGNORECASE)
    if not m:
        return 0.0
    a = float(m.group(1).replace(",", "."))
    b = float(m.group(2).replace(",", "."))
    return round(a * b, 2)


def parse_m2_direto(value: str) -> float:
    """Converte string de m² direto (ex: '25,50' ou '25.50') em float."""
    s = (value or "").strip().replace(",", ".")
    if not s:
        return 0.0
    try:
        return round(float(s), 2)
    except ValueError:
        return 0.0


def get_total_m2(data: dict) -> float:
    """
    Retorna a área total em m² a partir do payload.
    - duas_areas: m²₁ + m²₂.
    - tres_areas: m²₁ + m²₂ + m²₃.
    - m2_direto: valor informado diretamente (ex: "25,50").
    - Caso contrário (área única ou payload antigo): parse_medidas_m2(medidas).
    """
    if data.get("tipoMedidas") == "duas_areas":
        m1 = data.get("medidas1") or ""
        m2 = data.get("medidas2") or ""
        if m1 or m2:
            return round(parse_medidas_m2(m1) + parse_medidas_m2(m2), 2)
    if data.get("tipoMedidas") == "tres_areas":
        m1 = data.get("medidas1") or ""
        m2 = data.get("medidas2") or ""
        m3 = data.get("medidas3") or ""
        if m1 or m2 or m3:
            return round(
                parse_medidas_m2(m1) + parse_medidas_m2(m2) + parse_medidas_m2(m3), 2
            )
    if data.get("tipoMedidas") == "m2_direto":
        raw = (data.get("m2Direto") or "").strip()
        if raw:
            return parse_m2_direto(raw)
    return parse_medidas_m2(data.get("medidas") or "")


# Valor por m² do forro vinílico (R$), somado ao total quando Forro PVC = Vinílico
FORRO_VINILICO_VALOR_M2 = 120.0

# Acréscimos cartão: 5x = +6%, 10x = +10%
CARTAO_5X_ACRECIMO = 0.06
CARTAO_10X_ACRECIMO = 0.10


def get_valor_total_reais(data: dict) -> float:
    """
    Retorna o valor total em reais (float) com a mesma lógica de compute_valor_total.
    Usado para calcular o texto da forma de pagamento.
    """
    valor_m2_raw = data.get("valorM2") or ""
    valor_pilar_raw = (data.get("valorPilar") or "") if data.get("temPilar") == "Sim" else ""
    custo_raw = data.get("custoDeslocamento") or ""
    forro_pvc = (data.get("forroPvc") or "").strip()

    m2 = get_total_m2(data)
    valor_m2_reais = raw_to_reais(valor_m2_raw)
    valor_pilar_reais = raw_to_reais(valor_pilar_raw)
    custo_reais = raw_to_reais(custo_raw)

    parte_area = m2 * valor_m2_reais
    total_reais = parte_area + valor_pilar_reais + custo_reais

    if forro_pvc == "Vinílico":
        total_reais += FORRO_VINILICO_VALOR_M2 * m2

    return round(total_reais, 2)


# Tabela fixa valor/m² para Pergolado (tipo policarbonato x dimensão tubo)
PERGOLADO_VALOR_M2 = {
    ("Compacto 3mm", "150 x 50"): 1300.0,
    ("Compacto 3mm", "100 x 50"): 1200.0,
    ("Alveolar 6mm", "150 x 50"): 800.0,
    ("Alveolar 6mm", "100 x 50"): 700.0,
}


def get_valor_total_reais_pergolado(data: dict) -> float:
    """
    Valor total para Pergolado: m² × valor por m² + custo de deslocamento.
    Se valorM2 veio no payload (dimensão manual), usa esse valor; senão usa a tabela fixa.
    """
    m2 = get_total_m2(data)
    valor_m2_raw = data.get("valorM2")
    if valor_m2_raw and str(valor_m2_raw).strip():
        valor_m2_reais = raw_to_reais(str(valor_m2_raw))
    else:
        tipo = (data.get("tipoPolicarbonato") or "").strip()
        dimensao = (data.get("dimensaoTubo") or "").strip()
        valor_m2_reais = PERGOLADO_VALOR_M2.get((tipo, dimensao), 0.0)
    custo_desloc_reais = raw_to_reais(data.get("custoDeslocamento") or "")
    total_reais = m2 * valor_m2_reais + custo_desloc_reais
    return round(total_reais, 2)


def get_valor_cobertura_retratil_reais(data: dict) -> float:
    """
    Valor da cobertura retrátil SEM o custo da abertura automatizada.
    Usado como base para cálculo de juros (5x/10x) e para [Valor Total].
    Fórmula: (m² × valor por m²) + custo deslocamento.
    """
    m2 = get_total_m2(data)
    valor_m2_reais = raw_to_reais(data.get("valorM2") or "")
    custo_desloc_reais = raw_to_reais(data.get("custoDeslocamento") or "")
    return round(m2 * valor_m2_reais + custo_desloc_reais, 2)


def get_valor_total_reais_cobertura_retratil(data: dict) -> float:
    """
    Valor total para Cobertura Retrátil (cobertura + automatização):
    (m² × valor por m²) + custo deslocamento + custo da abertura automatizada.
    Em modo Manual, custo da abertura é 0. Valores no JSON em centavos.
    """
    valor_cobertura = get_valor_cobertura_retratil_reais(data)
    custo_abertura_reais = raw_to_reais(data.get("custoAberturaAutomatizada") or "")
    return round(valor_cobertura + custo_abertura_reais, 2)


def get_m2_porta(data: dict) -> float:
    """
    m² para Porta: (alturaPorta + alturaBandeirola) × (larguraPorta + larguraBandeirola).
    Valores vêm em metros (número). Se não houver bandeirola, usa só porta.
    """
    alt_porta = float(data.get("alturaPorta") or 0)
    larg_porta = float(data.get("larguraPorta") or 0)
    if data.get("bandeirola"):
        alt_band = float(data.get("alturaBandeirola") or 0)
        larg_band = float(data.get("larguraBandeirola") or 0)
        return round((alt_porta + alt_band) * (larg_porta + larg_band), 2)
    return round(alt_porta * larg_porta, 2)


def get_valor_total_reais_porta(data: dict) -> float:
    """
    Valor total para Porta: m² × valor por m² + custo de deslocamento.
    Valores no JSON: valorM2 e custoDeslocamento em centavos.
    """
    m2 = get_m2_porta(data)
    valor_m2_reais = raw_to_reais(data.get("valorM2") or "")
    custo_reais = raw_to_reais(data.get("custoDeslocamento") or "")
    return round(m2 * valor_m2_reais + custo_reais, 2)


def build_texto_forma_pagamento(
    total_a_vista_reais: float,
    total_a_vista_geral: float | None = None,
) -> str:
    """
    Monta o texto para [Valor p/ Forma de Pagamento]:
    "5x de R$ X,XX, 10x de R$ Y,YY ou R$ Z,ZZ A Vista"
    - 5x: total_a_vista_reais * 1.06 / 5
    - 10x: total_a_vista_reais * 1.10 / 10
    - À vista: total_a_vista_geral se informado, senão total_a_vista_reais (Cobertura Retrátil: juros só na cobertura; à vista = total geral).
    """
    total_5x = total_a_vista_reais * (1 + CARTAO_5X_ACRECIMO)
    total_10x = total_a_vista_reais * (1 + CARTAO_10X_ACRECIMO)
    parcela_5x = total_5x / 5
    parcela_10x = total_10x / 10

    avista_reais = total_a_vista_geral if total_a_vista_geral is not None else total_a_vista_reais
    s_avista = format_currency(str(int(round(avista_reais * 100))))
    s_5x = format_currency(str(int(round(parcela_5x * 100))))
    s_10x = format_currency(str(int(round(parcela_10x * 100))))

    return f"5x de {s_5x}, 10x de {s_10x} ou {s_avista} A Vista"


def compute_valor_total(data: dict) -> str:
    """
    Fórmula: Valor Total = (m² × valor por m²) + valor do pilar + custo de deslocamento.
    Se Forro PVC = Vinílico: soma também (120 * m²) ao total.
    Valores no JSON vêm em centavos (ex: "150000" = R$ 1.500,00).
    Retorna string formatada (ex: 'R$ 16.500,00').
    """
    valor_m2_raw = data.get("valorM2") or ""
    valor_pilar_raw = (data.get("valorPilar") or "") if data.get("temPilar") == "Sim" else ""
    custo_raw = data.get("custoDeslocamento") or ""
    forro_pvc = (data.get("forroPvc") or "").strip()

    m2 = get_total_m2(data)
    valor_m2_reais = raw_to_reais(valor_m2_raw)
    valor_pilar_reais = raw_to_reais(valor_pilar_raw)
    custo_reais = raw_to_reais(custo_raw)

    # Cálculo explícito: (m² * valor/m²) + pilar + deslocamento
    parte_area = m2 * valor_m2_reais
    total_reais = parte_area + valor_pilar_reais + custo_reais

    # Forro Vinílico: soma 120 * m² ao total
    if forro_pvc == "Vinílico":
        valor_forro_vinilico = FORRO_VINILICO_VALOR_M2 * m2
        total_reais += valor_forro_vinilico
        _log(f"Forro Vinílico: 120 * m² = {valor_forro_vinilico} | total_reais após forro = {total_reais}")

    total_cents = int(round(total_reais * 100))
    resultado_formatado = format_currency(str(total_cents))

    _log(
        f"compute_valor_total: m2={m2} | valor_m2_reais={valor_m2_reais} | "
        f"pilar_reais={valor_pilar_reais} | custo_reais={custo_reais} | "
        f"parte_area(m2*valor_m2)={parte_area} | total_reais={total_reais} | "
        f"total_cents={total_cents} -> '{resultado_formatado}'"
    )
    return resultado_formatado


def find_cell_by_text(wb, text: str):
    """
    Localiza a primeira célula que CONTÉM o texto indicado (busca em todas as planilhas).
    Retorna (sheet, endereço) ou (None, None) se não encontrar.
    """
    xlValues = -4163
    xlFormulas = -4123
    xlPart = 2

    for sheet in wb.sheets:
        for look_in in (xlValues, xlFormulas):
            try:
                found = sheet.api.Cells.Find(
                    What=text,
                    After=sheet.api.Cells(1, 1),
                    LookAt=xlPart,
                    LookIn=look_in,
                    SearchOrder=1,
                    SearchDirection=1,
                    MatchCase=False,
                )
                if found is not None:
                    addr = found.Address
                    if addr:
                        return (sheet, addr)
            except Exception:
                continue
    return (None, None)


def find_all_cells_by_text(wb, text: str):
    """
    Localiza TODAS as células que contêm o texto (em todas as planilhas).
    Usa Find + FindNext. Retorna lista de (sheet, endereço).
    """
    xlValues = -4163
    xlFormulas = -4123
    xlPart = 2
    results = []
    seen = set()  # (sheet_name, address) para não duplicar

    for sheet in wb.sheets:
        for look_in in (xlValues, xlFormulas):
            try:
                found = sheet.api.Cells.Find(
                    What=text,
                    After=sheet.api.Cells(1, 1),
                    LookAt=xlPart,
                    LookIn=look_in,
                    SearchOrder=1,
                    SearchDirection=1,
                    MatchCase=False,
                )
                if found is None:
                    continue
                first_addr = found.Address
                while found is not None:
                    addr = found.Address
                    key = (sheet.name, addr)
                    if key not in seen:
                        seen.add(key)
                        results.append((sheet, addr))
                    found = sheet.api.Cells.FindNext(found)
                    if found is None:
                        break
                    if found.Address == first_addr:
                        break
            except Exception:
                continue
    return results


def main() -> int:
    parser = argparse.ArgumentParser(description="Preenche Excel e exporta para PDF.")
    parser.add_argument("--template", required=True, help="Caminho do arquivo .xlsx modelo")
    parser.add_argument("--data", required=True, help="Caminho do arquivo .json com os dados")
    parser.add_argument("--output", required=True, help="Caminho do arquivo .pdf de saída")
    args = parser.parse_args()

    template_path = Path(args.template)
    data_path = Path(args.data)
    output_path = Path(args.output)

    if not template_path.exists():
        print(f"Erro: modelo não encontrado: {template_path}", file=sys.stderr)
        return 1
    if not data_path.exists():
        print(f"Erro: arquivo de dados não encontrado: {data_path}", file=sys.stderr)
        return 1

    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Data atual = momento da geração do PDF; formato brasileiro dd/mm/yyyy (dia/mês/ano)
    _hoje = datetime.now()
    data["dataAtual"] = f"{_hoje.day:02d}/{_hoje.month:02d}/{_hoje.year}"

    # Garantir "medidas" para placeholders/Excel conforme tipo de medição.
    # Quando são dimensões (1, 2 ou 3 áreas), prefixar com "medidas "; m² direto fica "X,XX metros quadrados".
    if data.get("tipoMedidas") == "duas_areas":
        m1, m2 = (data.get("medidas1") or "").strip(), (data.get("medidas2") or "").strip()
        if m1 and m2 and not (data.get("medidas") or "").strip():
            data["medidas"] = f"medidas {m1} e {m2}"
    elif data.get("tipoMedidas") == "tres_areas":
        m1 = (data.get("medidas1") or "").strip()
        m2 = (data.get("medidas2") or "").strip()
        m3 = (data.get("medidas3") or "").strip()
        if (m1 or m2 or m3) and not (data.get("medidas") or "").strip():
            data["medidas"] = "medidas " + " e ".join(x for x in (m1, m2, m3) if x)
    elif data.get("tipoMedidas") == "m2_direto":
        raw = (data.get("m2Direto") or "").strip()
        if raw:
            valor_m2 = parse_m2_direto(raw)
            # Cobertura Premium: "25,50 m²"; demais propostas: "25,50 metros quadrados"
            if data.get("tipoProposta") == "cobertura":
                data["medidas"] = f"{valor_m2:.2f}".replace(".", ",") + " m²"
            else:
                data["medidas"] = f"{valor_m2:.2f}".replace(".", ",") + " metros quadrados"
    else:
        # area_unica: prefixar com "medidas " o valor já vindo no payload
        med = (data.get("medidas") or "").strip()
        if med and not med.lower().startswith("medidas "):
            data["medidas"] = f"medidas {med}"

    is_pergolado = data.get("tipoProposta") == "pergolado"
    is_cobertura_retratil = data.get("tipoProposta") == "cobertura_retratil"
    is_porta = data.get("tipoProposta") == "porta"
    valor_cobertura_retratil_reais: float | None = None  # só usado para Cobertura Retrátil (base para juros e [Valor Total])
    if is_pergolado:
        total_a_vista_reais = get_valor_total_reais_pergolado(data)
        data["valorFormaPagamento"] = build_texto_forma_pagamento(total_a_vista_reais)
    elif is_cobertura_retratil:
        valor_cobertura_retratil_reais = get_valor_cobertura_retratil_reais(data)
        total_a_vista_reais = get_valor_total_reais_cobertura_retratil(data)  # total geral (cobertura + automatização)
        data["valorFormaPagamento"] = build_texto_forma_pagamento(
            valor_cobertura_retratil_reais, total_a_vista_reais
        )
    elif is_porta:
        total_a_vista_reais = get_valor_total_reais_porta(data)
        data["valorFormaPagamento"] = build_texto_forma_pagamento(total_a_vista_reais)
    else:
        total_a_vista_reais = get_valor_total_reais(data)
        data["valorFormaPagamento"] = build_texto_forma_pagamento(total_a_vista_reais)

    # Limpar log anterior para esta execução
    try:
        if LOG_PATH.exists():
            LOG_PATH.write_text("", encoding="utf-8")
    except Exception:
        pass

    _log(f"Dados recebidos (JSON): {json.dumps(data, ensure_ascii=False)}")
    _log(f"Campo 'custoDeslocamento' (bruto): {repr(data.get('custoDeslocamento'))}")

    try:
        import xlwings as xw
    except ImportError:
        print("Erro: xlwings não instalado. Execute: pip install xlwings", file=sys.stderr)
        return 1

    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(str(template_path.resolve()))

        for excel_text, json_key in FIELD_SEARCH.items():
            _log(f"Procurando no Excel (todas as planilhas) texto contendo: '{excel_text}' (campo JSON: '{json_key}')")
            sheet, address = find_cell_by_text(wb, excel_text)
            if sheet is None or address is None:
                _log(f"AVISO: campo '{excel_text}' NÃO encontrado em nenhuma planilha.")
                print(f"Aviso: campo '{excel_text}' não encontrado no Excel.", file=sys.stderr)
                continue
            _log(f"Célula encontrada: planilha '{sheet.name}', endereço {address}")
            raw = data.get(json_key, "")
            value = format_currency(raw) if raw else "R$ 0,00"
            _log(f"Valor a preencher: bruto={repr(raw)} -> formatado='{value}'")
            cell = sheet.range(address)
            cell.value = value
            _log(f"Valor escrito na célula {address}.")
            try:
                cell.number_format = "R$ #.##0,00"
                _log("Formato de número aplicado.")
            except Exception as fmt_err:
                _log(f"Formato de número não aplicado: {fmt_err}")

        # Placeholders: substituir apenas o placeholder dentro do texto da célula (resto do texto permanece)
        if is_pergolado:
            placeholder_map = FIELD_PLACEHOLDER_REPLACE_PERGOLADO
        elif is_cobertura_retratil:
            placeholder_map = FIELD_PLACEHOLDER_REPLACE_COBERTURA_RETRATIL
        elif is_porta:
            placeholder_map = FIELD_PLACEHOLDER_REPLACE_PORTA
        else:
            placeholder_map = FIELD_PLACEHOLDER_REPLACE
        for placeholder_text, json_key in placeholder_map.items():
            _log(f"Procurando placeholder no texto da célula: '{placeholder_text}' (campo JSON: '{json_key}')")
            sheet, address = find_cell_by_text(wb, placeholder_text)
            if sheet is None or address is None:
                _log(f"AVISO: placeholder '{placeholder_text}' NÃO encontrado em nenhuma planilha.")
                print(f"Aviso: placeholder '{placeholder_text}' não encontrado no Excel.", file=sys.stderr)
                continue
            _log(f"Célula encontrada: planilha '{sheet.name}', endereço {address}")
            cell = sheet.range(address)
            current = cell.value
            if current is None:
                current = ""
            current_str = str(current)
            value = data.get(json_key, "")
            if value is None:
                value = ""
            value_str = str(value)
            if is_pergolado and json_key == "corPolicarbonato":
                value_str = value_str.lower()
            new_text = current_str.replace(placeholder_text, value_str)
            # Célula de data: forçar formato Texto para o Excel não reinterpretar dd/mm/yyyy como mm/dd/yyyy
            if json_key == "dataAtual":
                try:
                    cell.number_format = "@"
                except Exception:
                    pass
            cell.value = new_text
            _log(f"Placeholder substituído: '{placeholder_text}' -> '{value_str}'; célula agora: '{new_text}'")

        # Valor nas células "[Valor Total]": valor parcelado em 10x (base + 10%). Cobertura Retrátil: juros só na cobertura.
        if is_cobertura_retratil and valor_cobertura_retratil_reais is not None:
            valor_10x_reais = valor_cobertura_retratil_reais * (1 + CARTAO_10X_ACRECIMO)
            _log(f"[Valor Total] Cobertura Retrátil: base cobertura={valor_cobertura_retratil_reais} -> 10x (só cobertura)")
        else:
            valor_10x_reais = total_a_vista_reais * (1 + CARTAO_10X_ACRECIMO)
        valor_10x_cents = int(round(valor_10x_reais * 100))
        valor_10x_str = format_currency(str(valor_10x_cents))
        _log(f"Valor em 10x (para células [Valor Total]): 10x = '{valor_10x_str}'")
        total_cells = find_all_cells_by_text(wb, FIELD_TOTAL_LABEL)
        _log(f"Células com '{FIELD_TOTAL_LABEL}': {len(total_cells)} encontrada(s)")
        for sheet, address in total_cells:
            cell = sheet.range(address)
            cell.value = valor_10x_str
            try:
                cell.number_format = "R$ #.##0,00"
            except Exception:
                pass
            _log(f"  Preenchido: planilha '{sheet.name}', {address}")

        # Cobertura Retrátil: preencher [Valor Total Geral] = cobertura com juros 10% + valor da automatização
        if is_cobertura_retratil:
            cobertura_10x = valor_cobertura_retratil_reais * (1 + CARTAO_10X_ACRECIMO)
            custo_abertura_reais = raw_to_reais(data.get("custoAberturaAutomatizada") or "")
            total_geral_reais = round(cobertura_10x + custo_abertura_reais, 2)
            total_geral_str = format_currency(str(int(round(total_geral_reais * 100))))
            total_geral_cells = find_all_cells_by_text(wb, FIELD_TOTAL_GERAL_LABEL)
            _log(f"Células com '{FIELD_TOTAL_GERAL_LABEL}': {len(total_geral_cells)} encontrada(s)")
            for sheet, address in total_geral_cells:
                cell = sheet.range(address)
                cell.value = total_geral_str
                try:
                    cell.number_format = "R$ #.##0,00"
                except Exception:
                    pass
                _log(f"  Preenchido: planilha '{sheet.name}', {address}")

        # Texto de especificação na célula D43 (Cobertura Premium, Cobertura Retrátil ou Porta; Pergolado não usa D43)
        if is_cobertura_retratil:
            texto_d43 = build_texto_especificacao_d43_retratil(data)
            texto_d43_excel = texto_d43.replace("\n", "\r\n")
            sheet_d43 = wb.sheets[0]
            cell_d43 = sheet_d43.range(D43_CELL)
            cell_d43.value = texto_d43_excel
            try:
                cell_d43.api.WrapText = True
            except Exception:
                pass
            _log(f"Célula {D43_CELL} preenchida com texto de especificação Cobertura Retrátil (planilha '{sheet_d43.name}').")
            # D44, M44, N44 só quando modo de abertura for Automatizada
            modo_abertura = (data.get("modoAbertura") or "").strip()
            if modo_abertura == "Automatizada":
                sheet_ret = wb.sheets[0]
                qtd_motores = (data.get("quantidadeMotores") or "").strip() or "[Quantidade de Motores]"
                texto_d44 = f"Automatizador para cobertura retrátil marca PPA Jetflex {qtd_motores}"
                sheet_ret.range(D44_CELL).value = texto_d44
                custo_abertura_raw = data.get("custoAberturaAutomatizada") or ""
                valor_abertura_fmt = format_currency(custo_abertura_raw)
                sheet_ret.range(M44_CELL).value = valor_abertura_fmt
                sheet_ret.range(N44_CELL).value = valor_abertura_fmt
                try:
                    sheet_ret.range(M44_CELL).number_format = "R$ #.##0,00"
                    sheet_ret.range(N44_CELL).number_format = "R$ #.##0,00"
                except Exception:
                    pass
                _log(f"Células {D44_CELL}, {M44_CELL}, {N44_CELL} preenchidas (modo Automatizada).")
        elif is_porta:
            texto_d43 = build_texto_especificacao_d43_porta(data)
            texto_d43_excel = texto_d43.replace("\n", "\r\n")
            sheet_d43 = wb.sheets[0]
            cell_d43 = sheet_d43.range(D43_CELL)
            cell_d43.value = texto_d43_excel
            try:
                cell_d43.api.WrapText = True
            except Exception:
                pass
            # Negrito nos cabeçalhos da descrição Porta
            try:
                # Primeira linha inteira em negrito
                primeira_linha = texto_d43_excel.split("\r\n")[0]
                if primeira_linha:
                    cell_d43.characters[0 : len(primeira_linha)].font.bold = True
                for cabecalho in ("Porta:", "Bandeirola:", "Alizar:", "Acabamento:", "Incluso:", "Não incluso:"):
                    pos = texto_d43_excel.find(cabecalho)
                    if pos >= 0:
                        cell_d43.characters[pos : pos + len(cabecalho)].font.bold = True
                _log("Negrito aplicado aos cabeçalhos da descrição Porta na célula D43.")
            except Exception as fmt_err:
                _log(f"Negrito D43 Porta não aplicado (ignorado): {fmt_err}")
            _log(f"Célula {D43_CELL} preenchida com texto de especificação Porta (planilha '{sheet_d43.name}').")
        elif not is_pergolado:
            texto_d43 = build_texto_especificacao_d43(data)
            texto_d43_excel = texto_d43.replace("\n", "\r\n")
            sheet_d43 = wb.sheets[0]
            cell_d43 = sheet_d43.range(D43_CELL)
            cell_d43.value = texto_d43_excel
            try:
                cell_d43.api.WrapText = True
            except Exception:
                pass
            try:
                if texto_d43_excel.startswith("Cobertura Premium"):
                    cell_d43.characters[0:18].font.bold = True
                for m in re.finditer(r"Item \d+:", texto_d43_excel):
                    cell_d43.characters[m.start() : m.end()].font.bold = True
                _log("Negrito aplicado a 'Cobertura Premium' e aos 'Item N:' na célula D43.")
            except Exception as fmt_err:
                _log(f"Negrito D43 não aplicado (ignorado): {fmt_err}")
            _log(f"Célula {D43_CELL} preenchida com texto de especificação (planilha '{sheet_d43.name}').")

        # Descrição adicional opcional na célula D44 (Pergolado e Cobertura Premium; Cobertura Retrátil usa D44 para automatizador)
        if not is_cobertura_retratil:
            desc_adicional = (data.get("descricaoAdicional") or "").strip()
            if desc_adicional:
                sheet_d44 = wb.sheets[0]
                sheet_d44.range(D44_CELL).value = desc_adicional.replace("\n", "\r\n")
                try:
                    sheet_d44.range(D44_CELL).api.WrapText = True
                except Exception:
                    pass
                _log(f"Célula {D44_CELL} preenchida com descrição adicional (planilha '{sheet_d44.name}').")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf_path = os.path.abspath(str(output_path.resolve()))
        _log(f"Exportando para PDF: {pdf_path}")
        wb.api.ExportAsFixedFormat(0, pdf_path)  # 0 = xlTypePDF
        wb.close()
        _log(f"PDF gerado com sucesso. Log completo em: {LOG_PATH}")
        return 0
    except Exception as e:
        _log(f"Erro: {e}")
        print(f"Erro ao gerar PDF: {e}", file=sys.stderr)
        return 1
    finally:
        if app:
            app.quit()


if __name__ == "__main__":
    sys.exit(main())
