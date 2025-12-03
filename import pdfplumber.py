#!/usr/bin/env python3
import pandas as pd
from pathlib import Path
import re
import sys

# =========================================================
# CONFIGURAÇÕES
# =========================================================
ROOT = Path("/home/usuario/usuario/pedro /smrede/")
INPUT_XLSX = ROOT / "10221 (1).xlsx"
OUTDIR = ROOT / "turmas_separadas"
MASTER_OUT = ROOT / "ALUNOS_TODOS.xlsx"
RELATORIO = ROOT / "relatorio_erros.txt"

# Garante que pastas existam
OUTDIR.mkdir(exist_ok=True, parents=True)

# =========================================================
# TURMAS VÁLIDAS
# =========================================================
TURMAS_VALIDAS = {
    "10221","10321","10421","11121","11221","11321","11421",
    "11511","11521","11611","11711","11811","11911",
    "22111","22211","27311"
}

# =========================================================
# REGEX
# =========================================================
RE_MATRICULA = re.compile(r"\b\d{6,10}\b")
RE_TEL = re.compile(r"(\(?\d{2}\)?\s*\d{4,5}-\d{4})")
RE_NUM5 = re.compile(r"\b\d{5}\b")

PALAVRAS_PROIBIDAS = [
    "RUA","AVENIDA","ESTRADA","TRAVESSA","ALAMEDA",
    "PRAIA","CAMINHO","RODOVIA","LADEIRA"
]

def clean(x):
    return "" if pd.isna(x) else str(x).strip()

def parece_nome(texto):
    if not texto:
        return False
    t = texto.strip().upper()
    for p in PALAVRAS_PROIBIDAS:
        if t.startswith(p + " "):
            return False
    if len(t.split()) < 2:
        return False
    if t.isdigit():
        return False
    return True


# =========================================================
# 1. CARREGA EXCEL
# =========================================================
try:
    sheets = pd.read_excel(INPUT_XLSX, sheet_name=None, header=None)
except Exception as e:
    print("\n❌ ERRO ao ler o Excel!")
    print(f"Motivo: {e}")
    sys.exit(1)

df = pd.concat(sheets.values(), ignore_index=True).dropna(how="all").reset_index(drop=True)

print(f"✔ Linhas carregadas: {len(df)}")

# Abre relatório
rel = open(RELATORIO, "w", encoding="utf-8")
rel.write("==== RELATÓRIO DE ERROS ====\n\n")

# =========================================================
# 2. PARSER
# =========================================================
alunos = []
i = 0

while i < len(df):

    L1_raw = clean(df.iloc[i, 0])
    L1_full = " ".join([clean(x) for x in df.iloc[i] if clean(x)])

    # Verifica nome
    if not parece_nome(L1_raw):
        i += 1
        continue

    nome = L1_raw

    # Matrícula
    matricula = ""
    m = RE_MATRICULA.search(L1_full)
    if m:
        matricula = m.group(0)
    else:
        rel.write(f"ALERTA: '{nome}' sem matrícula.\n")

    # Turma
    turma = ""
    for c in RE_NUM5.findall(L1_full):
        if c in TURMAS_VALIDAS:
            turma = c
            break

    if not turma:
        rel.write(f"ALERTA: '{nome}' sem turma válida.\n")

    # Responsável
    responsavel = L1_full
    for item in [nome, matricula, turma]:
        responsavel = responsavel.replace(item, "").strip()

    # ----------------------------------------------
    # L2 → ENDEREÇO
    # ----------------------------------------------
    endereco = ""
    telefone = ""

    if i + 1 < len(df):
        L2_raw = clean(df.iloc[i + 1, 0])
        texto_L2 = " ".join(clean(x) for x in df.iloc[i + 1] if clean(x))

        endereco = L2_raw

        m = RE_TEL.search(texto_L2)
        if m:
            telefone = m.group(1)

    # ----------------------------------------------
    # L3 → Telefone extra
    # ----------------------------------------------
    if not telefone and i + 2 < len(df):
        texto_L3 = " ".join(clean(x) for x in df.iloc[i + 2] if clean(x))
        m = RE_TEL.search(texto_L3)
        if m:
            telefone = m.group(1)

    # SALVA
    alunos.append({
        "NOME": nome,
        "MATRÍCULA": matricula,
        "TURMA": turma,
        "ENDEREÇO": endereco,
        "RESPONSÁVEL": responsavel,
        "TELEFONE/CELULAR": telefone
    })

    # Avança linhas
    if telefone and i + 2 < len(df):
        i += 3
    else:
        i += 2


# =========================================================
# 3. SALVAR ARQUIVO GERAL
# =========================================================
df_final = pd.DataFrame(alunos)
df_final["TURMA"] = df_final["TURMA"].apply(lambda t: t if t in TURMAS_VALIDAS else "")

df_final.to_excel(MASTER_OUT, index=False)

print(f"✔ Arquivo geral salvo em: {MASTER_OUT}")

# =========================================================
# 4. SALVAR POR TURMAS
# =========================================================
print("\n📚 Salvando turmas separadas...")

for turma in sorted(TURMAS_VALIDAS, key=int):
    grupo = df_final[df_final["TURMA"] == turma]
    grupo.to_excel(OUTDIR / f"{turma}.xlsx", index=False)
    print(f"  ✔ {turma}.xlsx criado ({len(grupo)} alunos)")

grupo_inv = df_final[df_final["TURMA"] == ""]
grupo_inv.to_excel(OUTDIR / "TURMA_INVALIDA.xlsx", index=False)
print(f"\n⚠ Alunos sem turma válida: {len(grupo_inv)}")

rel.close()
print("\n🎉 Processo concluído com sucesso!")
