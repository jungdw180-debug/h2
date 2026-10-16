"""충주바이오그린수소충전소 핵심설비 에너지분석 - 연차보고서용 그림 생성.

데이터 (../data):
  weekly_summary.csv                 주차별 종합 데이터 (kWh x kg x SEC), 2023-08 ~ 2025-11
  hourly_equipment_kwh_2025H2.csv    설비별 시간당 kWh, 2025-07-01 ~ 2025-11-30
  daily_h2_production_2025H2.csv     일별 수소 생산량(kg), 2025-07-01 ~ 2025-11-30

분석 구간(2차년도): 2024-12-30 ~ 2025-10-26 (43주).
타 충전소 비교(A/B), 베이스라인(22개소 통계), 대기전력 시뮬레이션 값은 원시 데이터가 없어 보고서 기재값을 사용.
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from matplotlib import font_manager as fm
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Patch

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data")
OUT = HERE

fm.fontManager.addfont(os.path.join(HERE, "..", "fonts", "NanumGothic-Regular.ttf"))
plt.rcParams.update({
    "font.family": "NanumGothic", "axes.unicode_minus": False,
    "axes.spines.top": False, "axes.spines.right": False,
    "font.size": 11, "savefig.dpi": 300, "savefig.bbox": "tight",
    "savefig.facecolor": "white",
})
NAVY, BLUE, LIGHT, ORANGE, GRAY, DARK = "#1f4e79", "#4c7fb0", "#a9c4de", "#d9822b", "#9aa5b1", "#2b2f36"
PROC_COLOR = {"생산": NAVY, "압축": LIGHT, "충전": ORANGE}

# ---------------------------------------------------------------- 데이터
PERIOD = ("2024-12-30", "2025-10-20")  # 43주 (주 시작일 기준)

# 핵심 설비 그룹 (보고서 명칭 -> 계측 설비)
GROUPS = {
    "저압압축기": ["압축기_LP01", "압축기_LP02"],
    "압축기용 냉각기": ["칠러_LP01", "칠러_LP02_HP"],
    "메인 칠러": ["칠러_PCHE"],
    "중압압축기": ["압축기_MP01", "압축기_MP02"],
    "서브 칠러": ["칠러_공용_MP_PCHE"],
    "고압압축기": ["압축기_HP01"],
}
PROCESS = {
    "생산": ["수소추출기_01", "수소추출기_02", "칠러_수소추출기01"],
    "압축": ["압축기_HP01", "압축기_LP01", "압축기_LP02", "압축기_MP01", "압축기_MP02", "칠러_LP01", "칠러_LP02_HP"],
    "충전": ["칠러_PCHE", "칠러_공용_MP_PCHE"],
}


def load_weekly():
    d = pd.read_csv(os.path.join(DATA, "weekly_summary.csv"))
    d["start"] = pd.to_datetime(d["Week"].str[:10])
    w = d[(d["start"] >= PERIOD[0]) & (d["start"] <= PERIOD[1])].copy().reset_index(drop=True)
    assert len(w) == 43, len(w)
    # 주간 kWh 는 MP1+MP2 합산 컬럼만 존재
    w["압축기_MP01_kWh"] = w["압축기(MP1+MP2)_kWh"]
    w["압축기_MP02_kWh"] = 0.0
    return w


def sec_sum(w, eqs):
    """설비 SEC 합 (주간). 가동하지 않은 설비(0)는 제외, 모두 미가동이면 NaN."""
    return w[[f"{e}_SEC" for e in eqs]].replace(0, np.nan).sum(axis=1, min_count=1)


def kwh_sum(w, eqs):
    return w[[f"{e}_kWh" for e in eqs]].sum(axis=1)


W = load_weekly()
HOURLY = pd.read_csv(os.path.join(DATA, "hourly_equipment_kwh_2025H2.csv"), parse_dates=["DateTime"]).set_index("DateTime")
HOURLY = HOURLY.drop(columns=["전력계측기_메인", "액화 질소 탱크로리 & 컨테이너 전원"])
H2_DAILY = pd.read_csv(os.path.join(DATA, "daily_h2_production_2025H2.csv"), parse_dates=["Date"]).set_index("Date")["Hydrogen_Production_kg"]

STATS = {}  # 보고서 문구 검증용 수치


def save(fig, name):
    fig.savefig(os.path.join(OUT, name))
    plt.close(fig)


def title(ax, text):
    ax.set_title(text, loc="left", fontsize=13, color=NAVY, pad=12)


def box(ax, x, y, w, h, text, fc="white", ec=NAVY, tc=DARK, fs=10.5, lw=1.4):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08", fc=fc, ec=ec, lw=lw))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=tc, linespacing=1.5)


def arrow(ax, p1, p2, color=GRAY, lw=1.8):
    ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=16, color=color, lw=lw))


# ================================================================ 1차년도
def fig01_monitoring_system():
    fig, ax = plt.subplots(figsize=(9.5, 5.0)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 6)
    ax.text(0, 5.95, "충주바이오그린수소충전소 실시간 전력 모니터링 시스템 구성", fontsize=13, color=NAVY, va="top")
    ax.text(0.1, 5.05, "4대 핵심 설비", fontsize=10.5, color=GRAY)
    for i, t in enumerate(["수소 압축기", "압축기용 냉각기", "충전기", "충전기용 냉각기(칠러)"]):
        box(ax, 0.1, 3.95 - i * 1.0, 2.2, 0.75, t, fs=9.5)
        arrow(ax, (2.32, 4.32 - i * 1.0), (2.95, 2.7 + (1.5 - i) * 0.3))
    box(ax, 3.0, 1.5, 1.8, 2.4, "메인 분전반\n+\n설비별\n적산전력계", fc="#eef3f9")
    arrow(ax, (4.82, 2.7), (5.35, 2.7))
    box(ax, 5.4, 1.5, 1.8, 2.4, "실시간 전력\n모니터링\n시스템", fc="#eef3f9")
    ax.text(6.3, 1.2, "수집: 전압·전류·전력량\n역률·온도·가동상태 (시간 단위)", fontsize=8.8, color=GRAY, ha="center", va="top", linespacing=1.5)
    arrow(ax, (7.22, 2.7), (7.6, 2.7))
    ax.text(7.65, 5.35, "분석·활용", fontsize=10.5, color=GRAY)
    for i, t in enumerate(["에너지 피크 관리", "부하 구간별 가동률\n(경·중·대부하)", "설비별 전력 사용량",
                           "기간별 소모 트렌드\n(일·주·월·연)", "SEC(kWh/kgH$_2$)\n월별 추적"]):
        box(ax, 7.65, 4.3 - i * 0.98, 2.3, 0.72, t, fs=8.8, ec=BLUE)
    save(fig, "fig01_monitoring_system.png")


# ================================================================ 2차년도: 43주 분석
def fig06_weekly_process_kwh():
    x = np.arange(len(W))
    fig, ax = plt.subplots(figsize=(10, 4.6))
    bottom = np.zeros(len(W))
    for p in ["생산", "압축", "충전"]:
        v = kwh_sum(W, PROCESS[p]).values
        ax.bar(x, v, bottom=bottom, color=PROC_COLOR[p], width=0.8, label=f"{p} 공정")
        bottom += v
    ax.set_ylabel("주간 에너지 사용량 (kWh)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax2 = ax.twinx()
    ax2.plot(x, W["Hydrogen_Production"], color=DARK, lw=1.6, marker="o", ms=3, label="수소 생산량")
    ax2.set_ylabel("주간 수소 생산량 (kg)")
    ax2.spines["right"].set_visible(True); ax2.spines["top"].set_visible(False)
    ax2.set_ylim(0, W["Hydrogen_Production"].max() * 1.25)
    ax.set_ylim(0, bottom.max() * 1.25)
    ticks = list(range(0, len(W), 4))
    ax.set_xticks(ticks); ax.set_xticklabels([W["start"][i].strftime("%y.%m.%d") for i in ticks], rotation=45, ha="right", fontsize=9)
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left", ncol=4, frameon=False, fontsize=9.5)
    title(ax, "주차별 공정별 에너지 사용량 및 수소 생산량 (2024.12.30 ~ 2025.10.26, 43주)")
    save(fig, "fig06_weekly_process_kwh.png")


def fig07_process_share():
    kwh = {p: kwh_sum(W, PROCESS[p]).sum() for p in PROCESS}
    sec = {p: sec_sum(W, PROCESS[p]).mean() for p in PROCESS}
    kt, st = sum(kwh.values()), sum(sec.values())
    STATS["process_kwh_share"] = {p: round(v / kt * 100, 1) for p, v in kwh.items()}
    STATS["process_sec"] = {p: round(v, 1) for p, v in sec.items()}
    STATS["process_sec_share"] = {p: round(v / st * 100, 1) for p, v in sec.items()}

    fig, ax = plt.subplots(figsize=(9, 4.0))
    rows = [("에너지 소비량(kWh) 비중", {p: v / kt * 100 for p, v in kwh.items()}, None),
            ("단위 수소당 에너지(SEC) 비중", {p: v / st * 100 for p, v in sec.items()}, sec)]
    for yi, (_, share, raw) in zip([1, 0], rows):
        left = 0
        for p in ["생산", "압축", "충전"]:
            v = share[p]
            ax.barh(yi, v, left=left, color=PROC_COLOR[p], height=0.55, edgecolor="white")
            t = f"{p} {v:.1f}%" + (f"\n({raw[p]:.1f} kWh/kg)" if raw else "")
            ax.text(left + v / 2, yi, t, ha="center", va="center", color=DARK if p == "압축" else "white", fontsize=10)
            left += v
    ax.set_yticks([1, 0]); ax.set_yticklabels([r[0] for r in rows], fontsize=11)
    ax.set_xlim(0, 100); ax.set_xlabel("비중 (%)"); ax.set_ylim(-1.3, 1.6)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    k, s = STATS["process_kwh_share"]["충전"], STATS["process_sec_share"]["충전"]
    ax.annotate(f"충전 공정: 소비량 비중 {k}% → SEC 비중 {s}% (약 {s / k:.1f}배)\n→ 시스템 내 핵심 비효율 구간",
                xy=(80, -0.3), xytext=(28, -0.95), fontsize=9.5, color=ORANGE,
                arrowprops=dict(arrowstyle="->", color=ORANGE, lw=1.5), va="center")
    title(ax, "공정별 에너지 소비 비중 vs SEC 비중 (43주)")
    fig.text(0.01, -0.04, "※ 공정 SEC = 공정 내 설비별 SEC의 합(주간 평균). 생산: 개질기·개질기 칠러 / 압축: 압축기·압축기용 냉각기 / 충전: 메인·서브 칠러",
             fontsize=8, color=GRAY)
    save(fig, "fig07_process_share.png")


def fig08_process_sec_vs_throughput():
    cases = [("생산", W["Hydrogen_Production"], "주간 수소 생산량 (kg)"),
             ("압축", W["Hydrogen_Production"], "주간 수소 생산량 (kg)"),
             ("충전", W["Car_Charging_kg"], "주간 차량 충전량 (kg)")]
    fig, axes = plt.subplots(1, 3, figsize=(11, 3.8))
    STATS["corr"] = {}
    for ax, (p, xv, xl) in zip(axes, cases):
        y = sec_sum(W, PROCESS[p])
        m = y.notna() & (xv > 0)
        r = np.corrcoef(xv[m], y[m])[0, 1]
        STATS["corr"][p] = round(r, 2)
        ax.scatter(xv[m], y[m], s=22, color=PROC_COLOR[p], edgecolor=NAVY if p == "압축" else "none", alpha=0.85)
        b = np.polyfit(xv[m], y[m], 1); xs = np.linspace(xv[m].min(), xv[m].max(), 50)
        ax.plot(xs, np.polyval(b, xs), color=DARK, lw=1.2, ls="--")
        ax.set_title(f"{p} 공정  (r = {r:.2f})", fontsize=11.5, color=NAVY)
        ax.set_xlabel(xl, fontsize=9.5); ax.set_ylim(bottom=0)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{int(v):,}"))
    axes[0].set_ylabel("공정 SEC (kWh/kgH$_2$)")
    fig.suptitle("수소 처리량에 따른 공정별 SEC 변화 (43주, 주간 단위)", x=0.01, ha="left", fontsize=13, color=NAVY)
    fig.tight_layout()
    save(fig, "fig08_process_sec_vs_throughput.png")


def fig09_equipment_kwh():
    vals = {g: kwh_sum(W, eqs).sum() for g, eqs in GROUPS.items()}
    total = sum(vals.values())
    STATS["equipment_kwh"] = {g: (round(v, 1), round(v / total * 100, 1)) for g, v in vals.items()}
    lp_weekly = kwh_sum(W, GROUPS["저압압축기"])
    STATS["lp_weekly_IQR"] = (round(lp_weekly.quantile(0.25)), round(lp_weekly.quantile(0.75)))
    items = sorted(vals.items(), key=lambda kv: kv[1])
    fig, ax = plt.subplots(figsize=(9, 4.3))
    names = [n for n, _ in items]; v = [x for _, x in items]
    cols = [ORANGE if n == "메인 칠러" else (BLUE if ("칠러" in n or "냉각" in n) else NAVY) for n in names]
    ax.barh(names, v, color=cols, height=0.62)
    for i, x in enumerate(v):
        ax.text(x + 3000, i, f"{x:,.1f} kWh ({x / total * 100:.1f}%)", va="center", fontsize=9.5)
    ax.set_xlim(0, max(v) * 1.5); ax.set_xlabel("누적 에너지 사용량 (kWh)")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
    ax.legend(handles=[Patch(color=NAVY, label="압축기"), Patch(color=BLUE, label="냉각기·서브 칠러"),
                       Patch(color=ORANGE, label="메인 칠러(충전)")], loc="lower right", frameon=False, fontsize=9.5)
    title(ax, "핵심 설비별 누적 에너지 사용량 (43주)")
    save(fig, "fig09_equipment_kwh.png")


def fig10_equipment_sec():
    order = ["고압압축기", "저압압축기", "중압압축기", "압축기용 냉각기", "서브 칠러", "메인 칠러"]
    data = [sec_sum(W, GROUPS[g]).dropna().values for g in order]
    means = [d.mean() for d in data]
    STATS["equipment_sec_mean"] = {g: round(m, 2) for g, m in zip(order, means)}
    STATS["main_chiller_sec_share"] = round(means[-1] / sum(means) * 100, 1)
    STATS["main_chiller_sec_max"] = round(data[-1].max(), 2)
    comp = np.concatenate(data[:4])
    STATS["compression_sec_P5_P95"] = (round(np.percentile(comp, 5), 2), round(np.percentile(comp, 95), 2))

    fig, ax = plt.subplots(figsize=(9, 4.6))
    bp = ax.boxplot(data, widths=0.5, patch_artist=True, showfliers=True,
                    medianprops=dict(color=DARK, lw=2), flierprops=dict(marker="o", ms=3.5, mfc=GRAY, mec="none"))
    for i, b in enumerate(bp["boxes"]):
        c = ORANGE if order[i] == "메인 칠러" else (BLUE if i in (3, 4) else NAVY)
        b.set(facecolor=c, alpha=0.75, edgecolor=c)
    for i, m in enumerate(means, 1):
        ax.plot(i, m, marker="D", color="white", mec=DARK, ms=6, zorder=5)
        ax.text(i + 0.3, m, f"{m:.1f}", va="center", fontsize=10.5, color=DARK)
    ax.set_xticks(range(1, len(order) + 1)); ax.set_xticklabels(order)
    ax.set_ylabel("주간 SEC (kWh/kgH$_2$)"); ax.set_ylim(0, 17)
    ax.annotate(f"메인 칠러: 평균 {means[-1]:.1f} kWh/kg\n핵심 설비 SEC 합의 {STATS['main_chiller_sec_share']}%\n최대 {data[-1].max():.1f} kWh/kg",
                xy=(5.75, 10.5), xytext=(2.6, 11.8), fontsize=10, color=ORANGE,
                arrowprops=dict(arrowstyle="->", color=ORANGE, lw=1.4))
    ax.plot([], [], marker="D", color="white", mec=DARK, ls="none", label="평균")
    ax.legend(loc="upper left", frameon=False, fontsize=9.5)
    title(ax, "핵심 설비별 단위 수소 처리량당 에너지 사용량(SEC) 분포 (43주)")
    save(fig, "fig10_equipment_sec.png")


def fig11_main_chiller_standby():
    c = HOURLY["칠러_PCHE"]
    night = c[(c.index.hour >= 20) | (c.index.hour < 8)]
    day = c[(c.index.hour >= 8) & (c.index.hour < 19)]
    STATS["chiller_hourly"] = dict(period=f"{c.index.min():%Y-%m-%d} ~ {c.index.max():%Y-%m-%d}", hours=len(c),
                                   night_mean=round(night.mean(), 1), day_mean=round(day.mean(), 1), max=round(c.max(), 1),
                                   night_share=round(night.sum() / c.sum() * 100, 1),
                                   night_kwh_per_day=round(night.sum() / (len(c) / 24), 1))

    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(11.5, 4.3), gridspec_kw=dict(width_ratios=[2.2, 1]))
    seg = c["2025-09-01":"2025-09-14"]
    t = np.arange(len(seg))
    ax.fill_between(t, seg.values, color=ORANGE, alpha=0.25, lw=0)
    ax.plot(t, seg.values, color=ORANGE, lw=1.1)
    for k, ts in enumerate(seg.index):
        if ts.hour >= 20 or ts.hour < 8:
            ax.axvspan(k - 0.5, k + 0.5, color=GRAY, alpha=0.12, lw=0)
    ax.axhline(night.mean(), color=NAVY, ls="--", lw=1.4)
    ax.text(len(seg) - 1, 3.0, f"야간 대기전력 평균 {night.mean():.1f} kWh/h (점선)", ha="right", color=NAVY, fontsize=10,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.9))
    ax.set_xticks(range(0, len(seg), 48)); ax.set_xticklabels([seg.index[i].strftime("%m/%d") for i in range(0, len(seg), 48)])
    ax.set_ylabel("메인 칠러 시간당 전력량 (kWh/h)"); ax.set_ylim(0, 38); ax.set_xlim(0, len(seg) - 1)
    ax.legend(handles=[Patch(color=GRAY, alpha=0.3, label="야간(20~08시)")], loc="upper left", frameon=False, fontsize=9.5)
    title(ax, "메인 칠러 시간별 전력 (예시 2주: 2025.9.1~9.14)")

    prof = c.groupby(c.index.hour).mean()
    cols = [GRAY if (h >= 20 or h < 8) else ORANGE for h in prof.index]
    ax2.bar(prof.index, prof.values, color=cols, width=0.8)
    ax2.axhline(night.mean(), color=NAVY, ls="--", lw=1.2)
    ax2.set_xticks([0, 6, 12, 18, 23]); ax2.set_xlabel("시각 (시)"); ax2.set_ylim(0, 38)
    ax2.set_title(f"시간대별 평균 ({STATS['chiller_hourly']['hours']:,}시간)", loc="left", fontsize=11.5, color=NAVY)
    ax2.text(0, 31, f"야간 소비 비중 {STATS['chiller_hourly']['night_share']}%\n(약 {STATS['chiller_hourly']['night_kwh_per_day']:.0f} kWh/일)",
             fontsize=10, color=NAVY)
    fig.text(0.01, -0.03, f"※ 데이터: {STATS['chiller_hourly']['period']} 시간 단위 계측. 주간(08~19시) 평균 {day.mean():.1f} kWh/h, 최대 {c.max():.1f} kWh/h",
             fontsize=8.5, color=GRAY)
    fig.tight_layout()
    save(fig, "fig11_main_chiller_standby.png")


def fig13_boxplot_models():
    # 타 충전소(A) 원시 데이터 미보유 -> 2차년도 보고서 기재 통계값 사용
    groups = [("(A) 도심형", "생산 공정", 5.5, 5.9, 6.6, 1.1, NAVY),
              ("(A) 도심형", "차량 충전", 6.8, 7.5, 8.7, 1.9, BLUE),
              ("(B) 마더 스테이션", "생산 공정", 4.5, 5.2, 6.7, 2.2, NAVY),
              ("(B) 마더 스테이션", "TT 충전(85%)", 7.6, 8.4, 10.2, 2.7, GRAY),
              ("(B) 마더 스테이션", "차량 충전(15%)", 9.3, 10.4, 12.8, 3.4, ORANGE)]
    xs = [1, 2, 3.6, 4.6, 5.6]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    for x, (_, proc, q1, med, q3, iqr, c) in zip(xs, groups):
        ax.add_patch(plt.Rectangle((x - 0.3, q1), 0.6, q3 - q1, fc=c, alpha=0.6, ec=c, lw=1.8))
        ax.hlines(med, x - 0.3, x + 0.3, color=DARK, lw=2.6)
        ax.text(x + 0.34, med, f"{med}", va="center", fontsize=10)
        ax.text(x, q3 + 0.35, f"IQR {iqr}", ha="center", fontsize=10, color=ORANGE if proc.startswith("차량 충전(") else DARK)
    ax.set_xticks(xs); ax.set_xticklabels([g[1] for g in groups], fontsize=10)
    ax.axvline(2.8, color=LIGHT, lw=1.2)
    ax.text(1.5, 15.3, "(A) 도심형 전용 충전소\n차량 충전 100% (연속)", ha="center", fontsize=10.5, color=NAVY)
    ax.text(4.6, 15.3, "(B) 광역 마더 스테이션\nTT 출하 85% · 차량 충전 15% (간헐)", ha="center", fontsize=10.5, color=NAVY)
    ax.set_xlim(0.4, 6.2); ax.set_ylim(3, 17.2)
    ax.set_ylabel("공정별 SEC (kWh/kgH$_2$)")
    ax.annotate("간헐 충전 + 상시 대기전력\n→ 중앙값·산포 최대", xy=(5.27, 12.3), xytext=(3.2, 13.4), fontsize=9.5, color=ORANGE,
                arrowprops=dict(arrowstyle="->", color=ORANGE, lw=1.3))
    title(ax, "운영 모델별 공정 SEC 분포 비교 (Box Plot, IQR)")
    fig.text(0.01, -0.03, "※ 박스: Q1~Q3, 굵은선·숫자: 중앙값 (2차년도 분석 통계값)", fontsize=8.5, color=GRAY)
    save(fig, "fig13_boxplot_models.png")


def fig14_baseline_target():
    fig, ax = plt.subplots(figsize=(6.8, 4.3))
    ax.bar(["베이스라인\n(상용 22개소 통계)", "최종 목표\n(20% 절감)"], [9.5, 7.6], color=[GRAY, NAVY], width=0.5)
    ax.text(0, 9.65, "9.5", ha="center", fontsize=14)
    ax.text(1, 7.75, "7.6", ha="center", fontsize=14)
    arrow(ax, (0.3, 9.3), (0.7, 7.85), color=ORANGE, lw=2)
    ax.text(0.5, 8.9, "-20%\n(-1.9)", ha="center", fontsize=11, color=ORANGE)
    ax.set_ylim(0, 11.2); ax.set_ylabel("에너지 효율 (kWh/kgH$_2$)")
    title(ax, "사업대상지 에너지 효율 베이스라인 및 목표")
    save(fig, "fig14_baseline_target.png")


# ================================================================ 다음 단계 계획
def fig18_load_heatmap():
    tot = HOURLY.sum(axis=1)
    piv = tot.groupby([tot.index.dayofweek, tot.index.hour]).mean().unstack()
    prof = tot.groupby(tot.index.hour).mean()
    top = prof.sort_values(ascending=False).index[:4].sort_values()
    STATS["load_profile"] = dict(night_kW=round(prof[[0, 1, 2, 3, 4, 5, 6, 7, 20, 21, 22, 23]].mean()),
                                 peak_hours=f"{top.min()}~{top.max() + 1}시", peak_kW=round(prof.max()))
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(10, 5.8), sharex=True, gridspec_kw=dict(height_ratios=[2.2, 1]))
    im = ax.imshow(piv.values, aspect="auto", cmap="YlOrRd", extent=(-0.5, 23.5, 6.5, -0.5))
    ax.set_yticks(range(7)); ax.set_yticklabels(["월", "화", "수", "목", "금", "토", "일"])
    cb = fig.colorbar(im, ax=[ax, ax2], fraction=0.025, pad=0.01); cb.set_label("충전소 전체 평균 전력 (kW)")
    title(ax, f"요일·시간대별 충전소 전체 전력 패턴 ({HOURLY.index.min():%Y.%m.%d} ~ {HOURLY.index.max():%Y.%m.%d})")
    cols = [ORANGE if h in top else (GRAY if (h >= 20 or h < 8) else BLUE) for h in prof.index]
    ax2.bar(prof.index, prof.values, color=cols, width=0.8)
    ax2.set_ylim(prof.min() * 0.85, prof.max() * 1.05)
    ax2.set_ylabel("평균 (kW)"); ax2.set_xlabel("시각 (시)"); ax2.set_xticks(range(0, 24, 2))
    ax2.legend(handles=[Patch(color=ORANGE, label="피크"), Patch(color=BLUE, label="주간"), Patch(color=GRAY, label="오프피크(야간)")],
               loc="upper left", ncol=3, frameon=False, fontsize=9)
    save(fig, "fig18_load_heatmap.png")


def fig16_three_stage_plan():
    fig, ax = plt.subplots(figsize=(10, 4.6)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 5)
    ax.text(0, 4.85, "다음 단계 연구개발 내용: 3단계 추진 체계", fontsize=13, color=NAVY, va="top")
    heads = ["1단계\n운영 패턴 정밀 분석", "2단계\n차등 운영 시나리오 수립", "3단계\n현장 적용 및 효과 검증"]
    prof = HOURLY.sum(axis=1).groupby(HOURLY.index.hour).mean()
    top = prof.sort_values(ascending=False).index[:4]
    lp = f"{top.min()}~{top.max() + 1}시"
    bodies = [f"· 원시 데이터 전처리\n  (노이즈 제거·결측 보정)\n· 월·일·요일·시간대별\n  충전 수요 패턴(히트맵)\n· 피크/오프피크 식별\n  (충전 12~14시,\n   전력 {lp})\n· 유휴시간-대기전력\n  상관관계 규명",
              "· 피크/오프피크별\n  압축기 가동률·부하 분산\n· 설비별 절전 모드 진입\n  로직 설계\n  (Operation→Standby\n   →Deep Sleep)\n· 소비량 시뮬레이션,\n  절감률·절감량 확정",
              "· 고효율 부품·설비 적용\n· Deep Sleep 로직 이식\n· Before/After 비교\n  (메인 칠러 야간\n   162→24 kWh/일)\n· 외부전문가·KTC 입회\n· ROI 평가·표준 운영\n  매뉴얼 작성"]
    xs = [0.1, 3.5, 6.9]
    for i in range(3):
        box(ax, xs[i], 3.55, 3.0, 0.95, heads[i], fc=NAVY, ec=NAVY, tc="white", fs=11)
        box(ax, xs[i], 0.2, 3.0, 3.2, bodies[i], fc="#f4f7fb", ec=LIGHT, fs=9.5)
        if i < 2:
            arrow(ax, (xs[i] + 3.02, 4.02), (xs[i + 1] - 0.02, 4.02), color=ORANGE, lw=2.4)
    save(fig, "fig16_three_stage_plan.png")


def fig19_state_transition():
    fig, ax = plt.subplots(figsize=(10, 4.6)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 5.0)
    ax.text(0, 4.95, "설비 유휴시간 기반 단계별 절전 모드 상태 전이(State Transition)", fontsize=13, color=NAVY, va="top")
    st = [("정상 운전\nOperation", "100% 부하 정상 가동", "절감 0%", NAVY),
          ("대기 모드\nStandby", "보조설비(냉각기 등) 저속 운전\n즉시 복귀 가능", "에너지 약 30% 절감", BLUE),
          ("절전 모드\nDeep Sleep", "필수 모니터링 전원 외\n설비 가동 중지", "에너지 약 90% 절감", ORANGE)]
    xs = [0.2, 3.7, 7.2]
    for (t, d, s, fc), x in zip(st, xs):
        box(ax, x, 1.85, 2.6, 1.3, t, fc=fc, ec=fc, tc="white", fs=12)
        ax.text(x + 1.3, 1.45, d, ha="center", va="top", fontsize=9.5, color=DARK, linespacing=1.5)
        ax.text(x + 1.3, 0.35, s, ha="center", va="center", fontsize=10.5, color=NAVY if fc == BLUE else fc)
    for i, lab in enumerate(["미가동\n10분 지속", "미가동\n60분 지속"]):
        arrow(ax, (xs[i] + 2.62, 2.7), (xs[i + 1] - 0.02, 2.7), color=DARK, lw=2)
        ax.text(xs[i] + 3.15, 2.95, lab, ha="center", fontsize=8.5, color=DARK)
    ax.annotate("", xy=(1.5, 3.2), xytext=(8.5, 3.2),
                arrowprops=dict(arrowstyle="-|>", color=GRAY, lw=1.4, connectionstyle="arc3,rad=0.15"))
    ax.text(5.0, 4.15, "충전 수요 발생 시 즉시 복귀", fontsize=9.5, color=GRAY, ha="center")
    save(fig, "fig19_state_transition.png")


def fig21_two_track_strategy():
    fig, ax = plt.subplots(figsize=(9.5, 4.6)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 5)
    ax.text(0, 4.9, "H/W·S/W 투트랙(Two-Track) 융합 연구개발 추진전략", fontsize=13, color=NAVY, va="top")
    box(ax, 0.1, 2.9, 4.4, 1.2, "Track 01  하드웨어 부문\n(참여기관 중심)", fc=BLUE, ec=BLUE, tc="white", fs=11.5)
    box(ax, 0.1, 1.55, 4.4, 1.15, "노후 설비 교체\n고효율 핵심 부품·설비 현장 도입", fc="#f4f7fb", ec=LIGHT, fs=10)
    box(ax, 5.5, 2.9, 4.4, 1.2, "Track 02  소프트웨어·운영 부문\n(주관: 고등기술연구원)", fc=NAVY, ec=NAVY, tc="white", fs=11.5)
    box(ax, 5.5, 1.55, 4.4, 1.15, "피크/오프피크 식별\n부하 반응형 차등 제어(Standby·Deep Sleep)", fc="#f4f7fb", ec=LIGHT, fs=10)
    arrow(ax, (2.3, 1.55), (4.0, 0.95), color=ORANGE, lw=2); arrow(ax, (7.7, 1.55), (6.0, 0.95), color=ORANGE, lw=2)
    box(ax, 2.0, 0.05, 6.0, 0.9, "융합 → 에너지 소비량 20% 절감 (9.5 → 7.6 kWh/kgH$_2$)\nKTC 등 공인기관 입회평가로 검증", fc=ORANGE, ec=ORANGE, tc="white", fs=10.5)
    save(fig, "fig21_two_track_strategy.png")


# ================================================================ 단계평가 추가 그림
def load_weekly_all():
    d = pd.read_csv(os.path.join(DATA, "weekly_summary.csv"))
    d["start"] = pd.to_datetime(d["Week"].str[:10])
    d["압축기_MP01_kWh"] = d["압축기(MP1+MP2)_kWh"]
    d["압축기_MP02_kWh"] = 0.0
    eq = [e for l in PROCESS.values() for e in l]
    d["kwh"] = kwh_sum(d, eq)
    return d[(d["start"] >= "2024-08-19") & (d["kwh"] > 0)].reset_index(drop=True)


def fig02_data_pipeline():
    fig, ax = plt.subplots(figsize=(11, 5.0)); ax.axis("off"); ax.set_xlim(0, 11); ax.set_ylim(0, 5.6)
    ax.text(0, 5.55, "수소충전소 에너지 데이터 수집·전처리 및 SEC 산출 체계", fontsize=13, color=NAVY, va="top")
    steps = [
        ("1. 데이터 수집", "설비별 적산전력계 15채널\n(2분 간격 V·I·PF·W)\n수소 생산 유량계(MFM)\n디스펜서 충전 기록"),
        ("2. 전처리", "전압 ≥400 V, 전류 ≥500 A\n이상치 제거\n설비별 전류·전력 상한 필터\n1분 단위 리샘플링"),
        ("3. 전력량 산정", "사다리꼴 적분 → kWh\n24 h 이상 결측 구간 제외\n시·일·주·월 단위 집계"),
        ("4. 수소량 산정", "생산량 = 유량 적분\n차량 충전량 = 디스펜서\nTT 출하량 = 생산 - 차량"),
        ("5. SEC 산정", "설비 kWh ÷ 설비별 분모\nSEC 40 kWh/kg 이상\n이상치 제외\n주간·월간 평균"),
    ]
    w, gap, x0 = 1.95, 0.22, 0.05
    for i, (h, b) in enumerate(steps):
        x = x0 + i * (w + gap)
        box(ax, x, 3.85, w, 0.75, h, fc=NAVY, ec=NAVY, tc="white", fs=11)
        box(ax, x, 1.95, w, 1.75, b, fc="#f4f7fb", ec=LIGHT, fs=8.8)
        if i < len(steps) - 1:
            arrow(ax, (x + w + 0.01, 4.22), (x + w + gap - 0.01, 4.22), color=ORANGE, lw=2)
    ax.text(0.05, 1.55, "SEC 분모 매핑", fontsize=10.5, color=NAVY)
    maps = [("총 생산량 (TT + 차량)", "개질기·개질기 칠러, 저압·중압 압축기, 압축기용 냉각기(LP01)", NAVY),
            ("차량 충전량", "고압압축기, 메인 칠러(PCHE)", ORANGE),
            ("TT + 2 × 차량 충전량", "서브 칠러(공용 MP·PCHE), 압축기용 냉각기(LP02·HP)", BLUE)]
    for i, (dn, eqs, c) in enumerate(maps):
        y = 1.05 - i * 0.45
        box(ax, 0.05, y, 2.6, 0.36, dn, fc=c, ec=c, tc="white", fs=9)
        ax.text(2.8, y + 0.18, eqs, va="center", fontsize=9, color=DARK)
    save(fig, "fig02_data_pipeline.png")


def fig04_full_period_sec():
    d = load_weekly_all()
    ok = d["Hydrogen_Production"] >= 1000
    q = d.groupby(d["start"].dt.to_period("Q")).apply(lambda g: pd.Series(dict(sec=g["kwh"].sum() / g["Hydrogen_Production"].sum(), kg=g["Hydrogen_Production"].sum())))
    STATS["quarterly_station_sec"] = {str(k): dict(sec=round(v["sec"], 2), kg=round(v["kg"])) for k, v in q.iterrows()}
    STATS["full_period"] = dict(start=f"{d['start'].min():%Y-%m-%d}", end=f"{(d['start'].max() + pd.Timedelta(days=6)):%Y-%m-%d}",
                                weeks=len(d), kwh=round(d["kwh"].sum()), kg=round(d["Hydrogen_Production"].sum()),
                                sec=round(d["kwh"].sum() / d["Hydrogen_Production"].sum(), 2), low_weeks=int((~ok).sum()))

    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(11, 6.2), sharex=True, gridspec_kw=dict(height_ratios=[2.3, 1]))
    x = d["start"]; bottom = np.zeros(len(d))
    for p in ["생산", "압축", "충전"]:
        v = np.where(ok, kwh_sum(d, PROCESS[p]) / d["Hydrogen_Production"].where(ok, np.nan), 0)
        ax.bar(x, v, bottom=bottom, width=6, color=PROC_COLOR[p], label=f"{p} 공정")
        bottom += v
    for k, r in q.iterrows():
        mid = k.start_time + (k.end_time - k.start_time) / 2
        ax.hlines(r["sec"], max(k.start_time, x.min()), min(k.end_time, x.max() + pd.Timedelta(days=6)), color=DARK, lw=2)
        ax.text(mid, r["sec"] + 1.0, f"{r['sec']:.1f}", ha="center", fontsize=10, color=DARK,
                bbox=dict(fc="white", ec="none", pad=1, alpha=0.8))
    ax.axhline(9.5, color=GRAY, ls="--", lw=1.2); ax.text(pd.Timestamp("2024-10-31"), 9.9, "베이스라인 9.5", fontsize=9, color=GRAY)
    ax.axhline(7.6, color=ORANGE, ls="--", lw=1.2); ax.text(pd.Timestamp("2024-10-31"), 6.8, "목표 7.6", fontsize=9, color=ORANGE)
    ax.set_ylim(0, 30); ax.set_ylabel("충전소 SEC (kWh/kgH$_2$)")
    ax.plot([], [], color=DARK, lw=2, label="분기 평균 SEC")
    ax.legend(loc="upper right", ncol=4, frameon=False, fontsize=9.5)
    title(ax, f"모니터링 전 기간 주간 공정별 SEC 및 분기 평균 추이 ({STATS['full_period']['start'][:7].replace('-', '.')} ~ {STATS['full_period']['end'][:7].replace('-', '.')})")
    ax2.bar(x, d["Hydrogen_Production"], width=6, color=np.where(ok, BLUE, GRAY))
    ax2.set_ylabel("수소 생산량\n(kg/주)"); ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax2.legend(handles=[Patch(color=GRAY, label="저가동 주(생산 1,000 kg 미만, SEC 미표시)")], loc="upper left", frameon=False, fontsize=9)
    fig.text(0.01, -0.02, "※ 공정 SEC = 공정 kWh ÷ 주간 총 수소 생산량 (적층 합 = 충전소 SEC). 분기 평균 = 분기 총 kWh ÷ 분기 총 생산량", fontsize=8.5, color=GRAY)
    fig.tight_layout()
    save(fig, "fig04_full_period_sec.png")


def fig05_monthly_2025_heatmap():
    m = pd.read_csv(os.path.join(DATA, "monthly_2025_equipment_kwh.csv"), index_col="Month")
    STATS["annual_2025_kwh"] = round(m["Total_Sum"].sum())
    order = [("수소추출기_01", "개질기 #1"), ("칠러_수소추출기01", "개질기 #1 칠러"), ("압축기_LP01", "저압압축기 #1"), ("칠러_LP01", "압축기용 냉각기 #1"),
             ("수소추출기_02", "개질기 #2"), ("압축기_LP02", "저압압축기 #2"), ("칠러_LP02_HP", "압축기용 냉각기 #2·HP"),
             ("압축기_MP01", "중압압축기 #1"), ("압축기_MP02", "중압압축기 #2"), ("압축기_HP01", "고압압축기"),
             ("칠러_PCHE", "메인 칠러"), ("칠러_공용_MP_PCHE", "서브 칠러")]
    hm = m[[c for c, _ in order]].T / 1000
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(10, 7.2), gridspec_kw=dict(height_ratios=[3.2, 1]), sharex=True)
    im = ax.imshow(hm.values, aspect="auto", cmap="YlOrRd", extent=(0.5, 12.5, len(order) - 0.5, -0.5))
    for i in range(hm.shape[0]):
        for j in range(12):
            v = hm.values[i, j]
            ax.text(j + 1, i, f"{v:.1f}", ha="center", va="center", fontsize=7.5, color="white" if v > 30 else DARK)
    ax.set_yticks(range(len(order))); ax.set_yticklabels([n for _, n in order], fontsize=9.5)
    ax.axhline(3.5, color=NAVY, lw=1.5); ax.axhline(6.5, color=NAVY, lw=1.5)
    ax.text(12.7, 1.5, "1계열", va="center", fontsize=9.5, color=NAVY); ax.text(12.7, 5.0, "2계열\n(5~6월 증설)", va="center", fontsize=9.5, color=NAVY)
    cb = fig.colorbar(im, ax=[ax, ax2], fraction=0.025, pad=0.15); cb.set_label("월간 전력량 (MWh)")
    title(ax, f"2025년 월별·설비별 에너지 사용량 (연간 {STATS['annual_2025_kwh'] / 1000:,.1f} MWh)")
    ax2.bar(m.index, m["Total_Sum"] / 1000, color=NAVY, width=0.7)
    for k, v in (m["Total_Sum"] / 1000).items():
        ax2.text(k, v + 3, f"{v:.0f}", ha="center", fontsize=8.5)
    ax2.set_ylim(0, 190); ax2.set_ylabel("합계 (MWh)"); ax2.set_xticks(range(1, 13)); ax2.set_xticklabels([f"{i}월" for i in range(1, 13)])
    save(fig, "fig05_monthly_2025_heatmap.png")


def fig12_day_night_equipment():
    h = HOURLY.copy(); hr = h.index.hour
    night = (hr >= 20) | (hr < 8); day = (hr >= 8) & (hr < 19)
    h["중압압축기"] = h["압축기_MP01"] + h["압축기_MP02"]
    items = [("칠러_PCHE", "메인 칠러"), ("칠러_공용_MP_PCHE", "서브 칠러"), ("압축기_HP01", "고압압축기"),
             ("칠러_LP02_HP", "압축기용 냉각기 #2·HP"), ("압축기_LP02", "저압압축기 #2"), ("중압압축기", "중압압축기"), ("수소추출기_02", "개질기 #2")]
    dn = {n: (h.loc[day, c].mean(), h.loc[night, c].mean()) for c, n in items}
    STATS["day_night_kW"] = {n: (round(a, 1), round(b, 1)) for n, (a, b) in dn.items()}
    fig, ax = plt.subplots(figsize=(10, 4.6))
    y = np.arange(len(items))[::-1]
    ax.barh(y + 0.19, [dn[n][0] for _, n in items], height=0.36, color=BLUE, label="주간(08~19시)")
    ax.barh(y - 0.19, [dn[n][1] for _, n in items], height=0.36, color=GRAY, label="야간(20~08시)")
    for yy, (_, n) in zip(y, items):
        dv, nv = dn[n]
        ax.text(dv + 0.8, yy + 0.19, f"{dv:.1f}", va="center", fontsize=9)
        ax.text(nv + 0.8, yy - 0.19, f"{nv:.1f}  ({nv / dv * 100:.0f}%)" if dv > 0 else f"{nv:.1f}", va="center", fontsize=9,
                color=ORANGE if n == "메인 칠러" else DARK)
    ax.set_yticks(y); ax.set_yticklabels([n for _, n in items])
    ax.get_yticklabels()[0].set_color(ORANGE)
    ax.set_xlabel("평균 전력 (kW = kWh/h)"); ax.set_xlim(0, 72)
    ax.legend(loc="center right", frameon=False, fontsize=9.5)
    r = dn["메인 칠러"][1] / dn["메인 칠러"][0] * 100
    ax.text(30, y[1] + 0.35, f"차량 충전 설비(고압압축기·서브 칠러)는 야간 정지\n메인 칠러만 주간 대비 {r:.0f}% 수준으로 상시 운전",
            fontsize=9.5, color=ORANGE, va="center")
    title(ax, f"설비별 주간·야간 평균 전력 비교 ({HOURLY.index.min():%Y.%m} ~ {HOURLY.index.max():%Y.%m}, 괄호: 주간 대비 야간 비율)")
    save(fig, "fig12_day_night_equipment.png")


def fig17_daytime_2026():
    d = pd.read_csv(os.path.join(DATA, "daytime_daily_avg_2026_monthly.csv"), index_col="설비명")
    names = {"수소추출기_02": "개질기 #2", "수소추출기_01": "개질기 #1", "칠러_수소추출기01": "개질기 #1 칠러", "압축기_LP02": "저압압축기 #2",
             "압축기_LP01": "저압압축기 #1", "압축기_MP01": "중압압축기 #1", "압축기_MP02": "중압압축기 #2", "압축기_HP01": "고압압축기",
             "칠러_LP02_HP": "압축기용 냉각기 #2·HP", "칠러_LP01": "압축기용 냉각기 #1", "칠러_PCHE": "메인 칠러", "칠러_공용_MP_PCHE": "서브 칠러"}
    d = d.loc[list(names)].rename(index=names)
    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    im = ax.imshow(d.values, aspect="auto", cmap="YlOrRd")
    for i in range(d.shape[0]):
        for j in range(d.shape[1]):
            v = d.values[i, j]
            ax.text(j, i, f"{v:.0f}", ha="center", va="center", fontsize=9, color="white" if v > 380 else DARK)
    ax.set_yticks(range(len(d))); ax.set_yticklabels(d.index)
    ax.get_yticklabels()[list(d.index).index("메인 칠러")].set_color(ORANGE)
    ax.set_xticks(range(5)); ax.set_xticklabels([f"{m}월" for m in d.columns])
    cb = fig.colorbar(im, ax=ax, fraction=0.04, pad=0.02); cb.set_label("주간(08~18시) 일평균 전력량 (kWh/일)")
    title(ax, "2026년 월별 설비별 주간 일평균 에너지 사용량 (3차년도 지속 모니터링)")
    save(fig, "fig17_daytime_2026.png")


def fig20_standby_simulation():
    c = HOURLY["칠러_PCHE"]; hr = c.index.hour; days = len(c) / 24
    day_e = c[(hr >= 8) & (hr < 20)].sum() / days
    night_e = c[(hr >= 20) | (hr < 8)].sum() / days
    save_e = ((c[hr == 20] * 0.3).sum() + (c[(hr >= 21) | (hr < 8)] * 0.9).sum()) / days
    total = HOURLY.sum().sum() / days
    kg = H2_DAILY.sum() / days
    STATS["standby_sim"] = dict(day_kwh=round(day_e, 1), night_kwh=round(night_e, 1), night_after=round(night_e - save_e, 1),
                                saving_kwh_day=round(save_e, 1), night_cut_pct=round(save_e / night_e * 100, 1),
                                chiller_cut_pct=round(save_e / (day_e + night_e) * 100, 1), station_cut_pct=round(save_e / total * 100, 2),
                                sec_before=round(total / kg, 2), sec_after=round((total - save_e) / kg, 2),
                                annual_mwh=round(save_e * 365 / 1000, 1))
    s = STATS["standby_sim"]
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(11, 4.4), gridspec_kw=dict(width_ratios=[1.3, 1]))
    labels = ["현행 운전", "Standby·Deep Sleep\n적용(시뮬레이션)"]
    ax.bar(labels, [day_e, day_e], color=BLUE, width=0.5, label="주간(08~20시)")
    ax.bar(labels, [night_e, night_e - save_e], bottom=[day_e, day_e], color=[GRAY, ORANGE], width=0.5, label="야간(20~08시)")
    ax.text(0, day_e + night_e / 2, f"야간\n{night_e:.0f}", ha="center", va="center", color="white", fontsize=10)
    ax.text(1.27, day_e + (night_e - save_e) / 2, f"야간 {night_e - save_e:.0f}", va="center", fontsize=10, color=ORANGE)
    ax.text(0, day_e / 2, f"주간\n{day_e:.0f}", ha="center", va="center", color="white", fontsize=10)
    ax.text(1, day_e / 2, f"주간\n{day_e:.0f}", ha="center", va="center", color="white", fontsize=10)
    ax.annotate(f"야간 대기전력 -{s['night_cut_pct']:.0f}%\n메인 칠러 전체 -{s['chiller_cut_pct']:.0f}%",
                xy=(0.85, day_e + night_e - save_e + 4), xytext=(0.28, day_e + night_e + 25), fontsize=10, color=ORANGE,
                arrowprops=dict(arrowstyle="->", color=ORANGE, lw=1.3))
    ax.set_ylabel("메인 칠러 일평균 전력량 (kWh/일)"); ax.set_ylim(0, (day_e + night_e) * 1.35)
    ax.legend(loc="upper right", frameon=False, fontsize=9)
    title(ax, "메인 칠러 대기전력 절감 시뮬레이션 (실측 3,664 h 적용)")
    ax2.axis("off")
    rows = [("메인 칠러 절감량", f"{s['saving_kwh_day']:.0f} kWh/일  (연 {s['annual_mwh']:.0f} MWh)"),
            ("충전소 전체 전력 절감률", f"{s['station_cut_pct']:.1f} %"),
            ("충전소 SEC", f"{s['sec_before']:.2f} → {s['sec_after']:.2f} kWh/kg"),
            ("가정", "20시 Standby(-30%)\n21~08시 Deep Sleep(-90%)\n야간 차량 충전 없음\n(고압압축기 야간 0.1 kW)")]
    for i, (k, v) in enumerate(rows):
        yy = 0.92 - i * 0.2
        ax2.text(0.0, yy, k, fontsize=10, color=GRAY, va="top", transform=ax2.transAxes)
        ax2.text(0.0, yy - 0.065, v, fontsize=11.5 if i < 3 else 9.5, color=NAVY, va="top", transform=ax2.transAxes, linespacing=1.4)
    fig.tight_layout()
    save(fig, "fig20_standby_simulation.png")


def fig03_station_layout():
    fig, ax = plt.subplots(figsize=(12, 5.4)); ax.axis("off"); ax.set_xlim(0, 12.2); ax.set_ylim(0, 6)
    ax.text(0, 5.95, "충주바이오그린수소충전소 수소 공급·압축·저장·충전 계통 및 전력 계측 설비", fontsize=13, color=NAVY, va="top")
    # 공급부
    box(ax, 0.05, 4.15, 1.75, 0.8, "튜브트레일러\n(Off-site, 200 bar)", fs=9)
    box(ax, 0.05, 2.95, 1.75, 0.8, "개질기 #1·#2\n(On-site)", fc=NAVY, ec=NAVY, tc="white", fs=9.5)
    box(ax, 2.2, 3.55, 1.25, 0.8, "버퍼 탱크\n5,500 L", fs=9)
    arrow(ax, (1.82, 4.55), (2.18, 4.05)); arrow(ax, (1.82, 3.35), (2.18, 3.85))
    # 압축·저장 라인
    line = [("1. 저압압축기\nLP", "7 → 90 bar", "LP Bank\n720 L"),
            ("2. 중압압축기\nMP", "90 → 200/450 bar", "MP Bank\n720 L"),
            ("3. 고압압축기\nHP", "450 → 850 bar", "HP Bank\n340 L")]
    x = 3.8
    arrow(ax, (3.47, 3.95), (x - 0.02, 3.95))
    for comp, pr, bank in line:
        box(ax, x, 3.55, 1.15, 0.8, comp, fc=NAVY, ec=NAVY, tc="white", fs=9)
        ax.text(x + 0.575, 4.45, pr, ha="center", fontsize=8.5, color=DARK)
        arrow(ax, (x + 1.17, 3.95), (x + 1.33, 3.95))
        box(ax, x + 1.35, 3.62, 0.7, 0.66, bank, fs=8)
        arrow(ax, (x + 2.07, 3.95), (x + 2.23, 3.95))
        x += 2.25
    box(ax, x, 3.55, 1.0, 0.8, "충전기\n(Dispenser)", fc=ORANGE, ec=ORANGE, tc="white", fs=9)
    arrow(ax, (x + 1.02, 3.95), (x + 1.3, 3.95)); ax.text(x + 1.35, 3.95, "수소차\n버스", va="center", fontsize=9, color=DARK)
    # 냉각 계통
    box(ax, 0.05, 1.45, 1.75, 0.8, "개질기 칠러", fc=LIGHT, ec=BLUE, fs=9.5)
    box(ax, 5.1, 1.45, 2.3, 0.8, "4. 압축기용 칠러\n(LP·MP·HP 열교환)", fc=LIGHT, ec=BLUE, fs=9.5)
    box(ax, 8.4, 1.45, 1.35, 0.8, "5. 서브 칠러\n(충전 보조)", fc=LIGHT, ec=BLUE, fs=9)
    box(ax, 9.95, 1.45, 1.45, 0.8, "6. 메인 칠러\n(충전기 직결)", fc=ORANGE, ec=ORANGE, tc="white", fs=9)
    kw = dict(color=BLUE, lw=1.3, ls="--")
    ax.plot([0.92, 0.92], [2.25, 2.95], **kw)
    for cx in (4.375, 6.625, 8.875):
        ax.plot([6.25, cx], [2.25, 3.55], **kw)
    ax.plot([9.07, x + 0.25], [2.25, 3.55], **kw)
    ax.plot([10.67, x + 0.75], [2.25, 3.55], color=ORANGE, lw=1.6, ls="--")
    ax.text(0.05, 0.85, "※ 개질기·압축기·칠러 등 설비별 적산전력계(15채널) 설치, 점선: 냉각 계통", fontsize=8.8, color=GRAY)
    ax.text(0.05, 0.45, "※ 2025년 5~6월 제2계열(개질기 #2, 저압압축기 #2, 압축기용 칠러 #2·HP) 증설", fontsize=8.8, color=GRAY)
    save(fig, "fig03_station_layout.png")


def fig15_ktc_verification():
    # KTC 시험성적서 GT2025-00687 (측정 2025.02.13~08.04, 4,144 h) vs 모니터링 주간 데이터(2025.02.10~08.03)
    d = pd.read_csv(os.path.join(DATA, "weekly_summary.csv")); d["start"] = pd.to_datetime(d["Week"].str[:10])
    w = d[(d["start"] >= "2025-02-10") & (d["start"] <= "2025-07-28")]
    items = [("개질기 #1", 130551, ["수소추출기_01"]), ("개질기 칠러", 161646, ["칠러_수소추출기01"]),
             ("압축기 LP #1", 78836, ["압축기_LP01"]), ("압축기 MP1+MP2", 50115 + 14308, ["압축기(MP1+MP2)"]),
             ("압축기 HP", 12662, ["압축기_HP01"]), ("메인 칠러(PCHE)", 69205, ["칠러_PCHE"]),
             ("서브 칠러(MP_PCHE)", 22404, ["칠러_공용_MP_PCHE"])]
    rows = [(n, k, w[[c + "_kWh" for c in cols]].sum().sum()) for n, k, cols in items]
    h2 = w["Hydrogen_Production"].sum()
    errs = [(m - k) / k * 100 for _, k, m in rows] + [(h2 - 76101) / 76101 * 100]
    STATS["ktc"] = dict(rows={n: (k, round(m)) for n, k, m in rows}, h2=(76101, round(h2)),
                        err_pct={n: round(e, 1) for (n, _, _), e in zip(rows, errs)}, mean_abs_err=round(np.mean(np.abs(errs)), 1))
    fig, ax = plt.subplots(figsize=(10, 4.8))
    names = [r[0] for r in rows] + ["수소 생산량 (kg)"]
    ktc = [r[1] for r in rows] + [76101]; mon = [r[2] for r in rows] + [h2]
    y = np.arange(len(names))[::-1]
    ax.barh(y + 0.19, ktc, height=0.36, color=GRAY, label="KTC 공인 측정값")
    ax.barh(y - 0.19, mon, height=0.36, color=NAVY, label="모니터링 데이터 (본 과제)")
    for yy, k, m, e in zip(y, ktc, mon, errs):
        ax.text(max(k, m) + 2500, yy, f"{e:+.1f}%", va="center", fontsize=10, color=DARK)
    ax.set_yticks(y); ax.set_yticklabels(names)
    ax.set_xlabel("누적 전력량 (kWh) / 수소 생산량 (kg)")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax.set_xlim(0, max(ktc + mon) * 1.18)
    ax.legend(loc="lower right", frameon=False, fontsize=9.5)
    title(ax, f"KTC 공인시험성적서 대비 모니터링 데이터 교차 검증 (평균 오차 {STATS['ktc']['mean_abs_err']}%)")
    fig.text(0.01, -0.03, "※ KTC: GT2025-00687, 2025.02.13~08.04(4,144 h). 모니터링: 주 단위 집계 2025.02.10~08.03. "
             "제2계열 설비와 칠러(LP_HP)는 측정 범위·채널 정의 차이로 비교에서 제외", fontsize=8.3, color=GRAY)
    save(fig, "fig15_ktc_verification.png")


if __name__ == "__main__":
    for fn in (fig01_monitoring_system, fig02_data_pipeline, fig03_station_layout, fig04_full_period_sec,
               fig05_monthly_2025_heatmap, fig06_weekly_process_kwh, fig07_process_share, fig08_process_sec_vs_throughput,
               fig09_equipment_kwh, fig10_equipment_sec, fig11_main_chiller_standby, fig12_day_night_equipment,
               fig13_boxplot_models, fig14_baseline_target, fig15_ktc_verification, fig16_three_stage_plan,
               fig17_daytime_2026, fig18_load_heatmap, fig19_state_transition, fig20_standby_simulation,
               fig21_two_track_strategy):
        fn()
    import json
    with open(os.path.join(OUT, "stats.json"), "w", encoding="utf-8") as f:
        json.dump(STATS, f, ensure_ascii=False, indent=1, default=str)
    print(json.dumps(STATS, ensure_ascii=False, indent=1, default=str))
