import os
import sys
from datetime import datetime

import pandas as pd
import yfinance as yf

# ============================================================
# CONFIG
# ============================================================
OUT_XLSX = "top10_acoes_br.xlsx"

# Lista base + novas ações do print (sem duplicar)
TICKERS_BR_BASE = [
    "PETR4", "VALE3", "ITUB4", "BBDC4", "BBAS3",
    "ABEV3", "WEGE3", "B3SA3", "RENT3", "ITSA4"
]

TICKERS_PRINT = ["BBSE3", "ODPV3", "KLBN4", "TAEE11", "SAPR4", "CMIG4"]

# Junta e remove duplicados preservando ordem
def unique_preserve_order(seq):
    seen = set()
    out = []
    for x in seq:
        x = x.strip().upper()
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out

TICKERS_BR = unique_preserve_order(TICKERS_BR_BASE + TICKERS_PRINT)


# ============================================================
# HELPERS
# ============================================================
def get_desktop_path() -> str:
    home = os.path.expanduser("~")
    desktop = os.path.join(home, "Desktop")
    return desktop if os.path.isdir(desktop) else home


def to_yahoo_symbol(br_ticker: str) -> str:
    return f"{br_ticker}.SA"


def safe_get(d: dict, *keys):
    for k in keys:
        if isinstance(d, dict) and k in d and d[k] is not None:
            return d[k]
    return None


def calc_opportunity(pe, dy_pct, pb):
    """
    Heurística (demo) - NÃO é recomendação:
    - "Oportunidade (heur.)" se:
        P/L <= 12  e DY >= 6% e P/VP <= 1.5
    - "Neutro" se tem dados mas não bate
    - "Sem dados" se faltar info
    """
    if pe is None or dy_pct is None or pb is None:
        return "Sem dados"
    try:
        if pe <= 12 and dy_pct >= 6 and pb <= 1.5:
            return "Oportunidade (heur.)"
        return "Neutro"
    except Exception:
        return "Sem dados"


def fmt_money(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


# ============================================================
# MAIN LOGIC
# ============================================================
def fetch_infos(tickers_br: list[str]) -> dict:
    yahoo = [to_yahoo_symbol(t) for t in tickers_br]
    pack = yf.Tickers(" ".join(yahoo))

    out = {}
    for t_br, t_y in zip(tickers_br, yahoo):
        try:
            tk = pack.tickers.get(t_y)
            info = tk.info if tk else {}
            out[t_br] = info or {}
        except Exception:
            out[t_br] = {}
    return out


def build_table(tickers_br: list[str], infos: dict) -> pd.DataFrame:
    # >>> AQUI: DataColeta só com o DIA <<<
    now_day = datetime.now().strftime("%d/%m/%Y")

    rows = []
    for t in tickers_br:
        info = infos.get(t, {})

        name = safe_get(info, "longName", "shortName") or t
        price = safe_get(info, "currentPrice", "regularMarketPrice")
        prev_close = safe_get(info, "previousClose")

        change_pct = None
        if price is not None and prev_close not in (None, 0):
            try:
                change_pct = (float(price) / float(prev_close) - 1.0) * 100.0
            except Exception:
                change_pct = None

        mcap = safe_get(info, "marketCap")
        pe = safe_get(info, "trailingPE", "forwardPE")
        pb = safe_get(info, "priceToBook")
        sector = safe_get(info, "sector") or "-"
        high52 = safe_get(info, "fiftyTwoWeekHigh")
        low52 = safe_get(info, "fiftyTwoWeekLow")
        vol = safe_get(info, "volume", "averageVolume")

        dy = safe_get(info, "dividendYield")
        dy_pct = None
        if dy is not None:
            try:
                dy_pct = float(dy) * 100.0
            except Exception:
                dy_pct = None

        opportunity = calc_opportunity(pe, dy_pct, pb)

        rows.append({
            "DataColeta": now_day,
            "Ticker": t,
            "Nome": name,
            "PrecoAtual": price,
            "VariacaoDiaPct": change_pct,
            "MarketCap": mcap,
            "PL": pe,
            "DYpct": dy_pct,
            "PVP": pb,
            "52wHigh": high52,
            "52wLow": low52,
            "Volume": vol,
            "Setor": sector,
            "Oportunidade": opportunity
        })

    return pd.DataFrame(rows)


def print_terminal(df: pd.DataFrame):
    print("\n==============================")
    print("ACOES BR (Yahoo Finance)")
    print("==============================\n")

    show = df[["Ticker", "PrecoAtual", "VariacaoDiaPct", "PL", "DYpct", "PVP", "MarketCap", "Oportunidade"]].copy()
    show["PrecoAtual"] = show["PrecoAtual"].apply(fmt_money)
    show["VariacaoDiaPct"] = show["VariacaoDiaPct"].apply(lambda x: "-" if pd.isna(x) or x is None else f"{x:.2f}%".replace(".", ","))
    show["DYpct"] = show["DYpct"].apply(lambda x: "-" if pd.isna(x) or x is None else f"{x:.2f}%".replace(".", ","))
    show["PL"] = show["PL"].apply(lambda x: "-" if pd.isna(x) or x is None else f"{x:.2f}".replace(".", ","))
    show["PVP"] = show["PVP"].apply(lambda x: "-" if pd.isna(x) or x is None else f"{x:.2f}".replace(".", ","))
    show["MarketCap"] = show["MarketCap"].apply(lambda x: "-" if pd.isna(x) or x is None else f"{int(x):,}".replace(",", "."))

    print(show.to_string(index=False))


def save_xlsx_overwrite(df: pd.DataFrame, out_path: str):
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except PermissionError:
            raise PermissionError(
                f"O arquivo está aberto no Excel e não pode ser substituído:\n{out_path}\n"
                "Feche o Excel e rode novamente."
            )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Acoes", index=False)


def main():
    try:
        print("Tickers usados:", ", ".join(TICKERS_BR))
        infos = fetch_infos(TICKERS_BR)
        df = build_table(TICKERS_BR, infos)

        print_terminal(df)

        desktop = get_desktop_path()
        out_file = os.path.join(desktop, OUT_XLSX)

        save_xlsx_overwrite(df, out_file)
        print(f"\nArquivo Excel atualizado em:\n{out_file}\n")
        print("Obs.: 'Oportunidade' é heurística (demo), não é recomendação.")

    except Exception as e:
        print(f"\nERRO: {e}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
