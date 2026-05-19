
import pandas as pd
import matplotlib.pyplot as plt

# ★これを追加
plt.rcParams['font.family'] = 'MS Gothic'

file_path = "input.xlsm"

# =========================
# ① 正しいヘッダで読み込み（ここが確定）
# =========================
df = pd.read_excel(file_path, header=91, engine="openpyxl")

print("✅ 列一覧")
print(df.columns)

# =========================
# ② 必要列取得
# =========================
col_order   = "OrderNo"
col_machine = "機種"
col_month   = "生産月"
col_serial  = "号機"
col_busho   = "BushoCd"

# =========================
# ③ 第2工場抽出（BushoCd=11）
# =========================
df2 = df[df[col_busho] == 11]

print("第2工場件数:", len(df2))

# =========================
# ④ 出荷のみ抽出
# =========================
df2 = df2[df2.astype(str).apply(lambda row: row.str.contains("出荷").any(), axis=1)]

print("出荷件数:", len(df2))

# =========================
# ⑤ 重複除外（最重要）
# =========================
df2 = df2.drop_duplicates(subset=[col_order, col_serial])

print("ユニーク台数:", len(df2))

# =========================
# ⑥ 月整形（YYYYMM → YYYY-MM）
# =========================
df2[col_month] = df2[col_month].astype(str)
df2["月"] = df2[col_month].str[:4] + "-" + df2[col_month].str[4:]

# =========================
# ⑦ 機種別集計
# =========================
machine_total = df2.groupby(col_machine).size().sort_values(ascending=False)

print("\n=== 機種別台数 ===")
print(machine_total)

# =========================
# ⑧ 月×機種ピボット
# =========================
pivot = pd.pivot_table(
    df2,
    index="月",
    columns=col_machine,
    aggfunc="size",
    fill_value=0
)

print("\n=== 月別テーブル ===")
print(pivot)

# =========================
# ⑨ CSV出力
# =========================
pivot.to_csv("monthly_machine_summary.csv", encoding="utf-8-sig")
machine_total.to_csv("machine_total_summary.csv", encoding="utf-8-sig")

print("\n✅ CSV出力完了")

# =========================
# ⑩ L2だけグラフ
# =========================
if "L2" in pivot.columns:
    pivot["L2"].plot(kind="bar")
    plt.title("第2工場 L2 月別台数")
    plt.xlabel("月")
    plt.ylabel("台数")
    plt.tight_layout()
    plt.savefig("L2_monthly.png")
    plt.show()

# =========================
# ⑪ ピーク月
# =========================
print("\n✅ ピーク月:", pivot.sum(axis=1).idxmax())