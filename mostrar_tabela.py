#!/usr/bin/env python3
"""
mostrar_tabela.py - Exibe qualquer tabela de texto de forma organizada.

Aceita tabelas coladas do Word/Excel/Packet Tracer (separadas por TAB),
separadas por 2+ espaços, ou por | ; , e as desenha com bordas alinhadas.

Uso:
  python mostrar_tabela.py                    # modo interativo (cole e dê Enter em linha vazia)
  python mostrar_tabela.py arquivo.txt        # lê de um arquivo
  cat arquivo.txt | python mostrar_tabela.py  # lê da entrada padrão
  python mostrar_tabela.py --exemplo          # mostra a tabela de rede do exemplo

Opções:
  -f, --formato {grade,simples,markdown}   estilo de saída (padrão: grade)
  -d, --delimitador DELIM                  força o separador (ex: ';' ou '\\t')
  --sem-cabecalho                          não trata a 1ª linha como cabeçalho
  --preencher                              repete o valor da célula de cima quando vazia
  --separar-linhas                         desenha linha entre todas as linhas
"""

import argparse
import re
import sys
import unicodedata

EXEMPLO = """Dispositivo\tInterface / VLAN\tEndereço IP\tMáscara de Sub-rede\tGateway Padrão
R1\tg0/0.10 (VLAN 10)\t192.168.10.1\t255.255.255.128\t-
\tg0/0.20 (VLAN 20)\t192.168.10.129\t255.255.255.192\t-
\tg0/1.30 (VLAN 30)\t192.168.20.1\t255.255.255.128\t-
\tg0/1.40 (VLAN 40)\t192.168.20.129\t255.255.255.192\t-
\tg0/2 (Link R2)\t10.0.0.1\t255.255.255.252\t-
R2\tg0/0 (Link R1)\t10.0.0.2\t255.255.255.252\t-
\tg0/1 (Link R3)\t10.0.0.5\t255.255.255.252\t-
\tg0/2.50 (VLAN 50)\t192.168.30.1\t255.255.255.128\t-
\tg0/2.60 (VLAN 60)\t192.168.30.129\t255.255.255.192\t-
R3\tg0/1 (Link R2)\t10.0.0.6\t255.255.255.252\t-
\tg0/0.70 (VLAN 70)\t192.168.40.1\t255.255.255.128\t-
\tg0/0.80 (VLAN 80)\t192.168.40.129\t255.255.255.192\t-
PCs\tDHCP / Estático\tIPs .2 e .130\t/25 e /26\tIPs .1 e .129
"""


# ----------------------------------------------------------------------
# Leitura / interpretação
# ----------------------------------------------------------------------
def detectar_delimitador(linhas):
    """Escolhe o separador mais provável olhando as linhas não vazias."""
    amostra = [l for l in linhas if l.strip()]
    for delim in ("\t", "|", ";"):
        if sum(delim in l for l in amostra) >= max(1, len(amostra) // 2):
            return delim
    if sum(bool(re.search(r"\S {2,}\S", l)) for l in amostra) >= max(1, len(amostra) // 2):
        return "  "  # 2+ espaços
    if sum("," in l for l in amostra) >= max(1, len(amostra) // 2):
        return ","
    return None  # cada linha vira uma única coluna


def dividir_linha(linha, delim):
    if delim is None:
        return [linha.strip()]
    if delim == "  ":
        # preserva célula vazia inicial (linha começando com muitos espaços)
        partes = re.split(r" {2,}", linha.rstrip())
        if linha.startswith("  "):
            partes = [""] + [p for p in partes if p != ""]
        return [p.strip() for p in partes]
    partes = linha.split(delim)
    if delim == "|":  # tabelas markdown: |a|b|
        if partes and partes[0].strip() == "":
            partes = partes[1:]
        if partes and partes[-1].strip() == "":
            partes = partes[:-1]
    return [p.strip() for p in partes]


def parsear(texto, delimitador=None):
    linhas = [l.rstrip("\r\n") for l in texto.splitlines()]
    linhas = [l for l in linhas if l.strip()]
    # remove linhas de separação markdown (|---|---|) e bordas ASCII (+---+)
    linhas = [l for l in linhas if not re.fullmatch(r"[\s|+:\-=─│┼┌┐└┘├┤┬┴╔╗╚╝═║]+", l)]
    if not linhas:
        return []
    if delimitador is None:
        delimitador = detectar_delimitador(linhas)
    else:
        delimitador = delimitador.replace("\\t", "\t")
    linhas_dividas = [dividir_linha(l, delimitador) for l in linhas]
    n = max(len(r) for r in linhas_dividas)
    return [r + [""] * (n - len(r)) for r in linhas_dividas]  # completa linhas curtas


# ----------------------------------------------------------------------
# Desenho
# ----------------------------------------------------------------------
def largura(texto):
    """Largura visual (considera acentos combinados e caracteres largos)."""
    total = 0
    for ch in texto:
        if unicodedata.combining(ch):
            continue
        total += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return total


def preencher(texto, tam, alinhamento="esq"):
    falta = tam - largura(texto)
    if alinhamento == "centro":
        e = falta // 2
        return " " * e + texto + " " * (falta - e)
    return texto + " " * falta


def calcular_larguras(linhas):
    return [max(largura(l[i]) for l in linhas) for i in range(len(linhas[0]))]


def renderizar_grade(linhas, cabecalho, separar_todas):
    w = calcular_larguras(linhas)

    def borda(esq, meio, dir_):
        return esq + meio.join("─" * (x + 2) for x in w) + dir_

    def linha(celulas, alin="esq"):
        return "│" + "│".join(" " + preencher(c, w[i], alin) + " " for i, c in enumerate(celulas)) + "│"

    # agrupa quando a 1ª coluna tem células vazias (estilo "células mescladas")
    agrupado = any(l[0] == "" for l in linhas[1 if cabecalho else 0:])

    saida = [borda("┌", "┬", "┐")]
    inicio = 0
    if cabecalho:
        saida.append(linha(linhas[0], "centro"))
        saida.append(borda("├", "┼", "┤"))
        inicio = 1
    for i, l in enumerate(linhas[inicio:], start=inicio):
        if i > inicio and (separar_todas or (agrupado and l[0] != "")):
            saida.append(borda("├", "┼", "┤"))
        saida.append(linha(l))
    saida.append(borda("└", "┴", "┘"))
    return "\n".join(saida)


def renderizar_simples(linhas, cabecalho):
    w = calcular_larguras(linhas)
    fmt = lambda l: "  ".join(preencher(c, w[i]) for i, c in enumerate(l)).rstrip()
    saida = []
    for i, l in enumerate(linhas):
        saida.append(fmt(l))
        if cabecalho and i == 0:
            saida.append("  ".join("─" * x for x in w))
    return "\n".join(saida)


def renderizar_markdown(linhas, cabecalho):
    w = calcular_larguras(linhas)
    fmt = lambda l: "| " + " | ".join(preencher(c, w[i]) for i, c in enumerate(l)) + " |"
    if not cabecalho:
        linhas = [[f"Col {i + 1}" for i in range(len(w))]] + linhas
        w = calcular_larguras(linhas)
    saida = [fmt(linhas[0]), "|" + "|".join("-" * (x + 2) for x in w) + "|"]
    saida += [fmt(l) for l in linhas[1:]]
    return "\n".join(saida)


def preencher_vazias(linhas, inicio):
    """Repete o valor de cima nas células vazias (útil p/ células mescladas)."""
    for i in range(inicio + 1, len(linhas)):
        for j, c in enumerate(linhas[i]):
            if c == "" and j == 0:
                linhas[i][j] = linhas[i - 1][j]
    return linhas


# ----------------------------------------------------------------------
# Entrada
# ----------------------------------------------------------------------
def ler_entrada(args):
    if args.exemplo:
        return EXEMPLO
    if args.arquivo:
        with open(args.arquivo, encoding="utf-8-sig") as f:
            return f.read()
    if sys.stdin.isatty():
        print("Cole sua tabela abaixo e pressione Enter em uma linha vazia para terminar:\n")
        linhas = []
        while True:
            try:
                l = input()
            except EOFError:
                break
            if l.strip() == "" and linhas:
                break
            linhas.append(l)
        return "\n".join(linhas)
    return sys.stdin.read()


def main():
    p = argparse.ArgumentParser(description="Exibe tabelas de texto de forma organizada.",
                                formatter_class=argparse.RawDescriptionHelpFormatter,
                                epilog=__doc__)
    p.add_argument("arquivo", nargs="?", help="arquivo com a tabela (opcional)")
    p.add_argument("-f", "--formato", choices=["grade", "simples", "markdown"], default="grade")
    p.add_argument("-d", "--delimitador", help="força o delimitador (ex: ';' ou '\\t')")
    p.add_argument("--sem-cabecalho", action="store_true")
    p.add_argument("--preencher", action="store_true")
    p.add_argument("--separar-linhas", action="store_true")
    p.add_argument("--exemplo", action="store_true", help="usa a tabela de exemplo")
    args = p.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    linhas = parsear(ler_entrada(args), args.delimitador)
    if not linhas:
        print("Nenhum dado encontrado.")
        return

    cabecalho = not args.sem_cabecalho
    if args.preencher:
        linhas = preencher_vazias(linhas, 1 if cabecalho else 0)

    if args.formato == "grade":
        print(renderizar_grade(linhas, cabecalho, args.separar_linhas))
    elif args.formato == "simples":
        print(renderizar_simples(linhas, cabecalho))
    else:
        print(renderizar_markdown(linhas, cabecalho))


if __name__ == "__main__":
    main()
