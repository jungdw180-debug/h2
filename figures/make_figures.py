"""충주바이오그린수소충전소 핵심설비 에너지분석 - 연차보고서용 그림 생성.
모든 수치는 연차보고서 본문 기재값만 사용 (원시 시계열 데이터 미사용)."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

for f in fm.findSystemFonts():
    if "wqy-zenhei" in f.lower():
        fm.fontManager.addfont(f)
plt.rcParams.update({
    "font.family": "WenQuanYi Zen Hei", "axes.unicode_minus": False,
    "axes.spines.top": False, "axes.spines.right": False,
    "font.size": 11, "savefig.dpi": 300, "savefig.bbox": "tight",
    "savefig.facecolor": "white",
})
OUT = os.path.dirname(os.path.abspath(__file__))
NAVY, BLUE, LIGHT, ORANGE, GRAY, DARK = "#1f4e79", "#4c7fb0", "#a9c4de", "#d9822b", "#9aa5b1", "#2b2f36"
PROC = {"생산": NAVY, "압축": LIGHT, "충전": ORANGE}


def save(fig, name):
    fig.savefig(os.path.join(OUT, name))
    plt.close(fig)


def box(ax, x, y, w, h, text, fc="white", ec=NAVY, tc=DARK, fs=10.5, lw=1.4, bold=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                fc=fc, ec=ec, lw=lw))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=tc,
            fontweight="bold" if bold else "normal", linespacing=1.5)


def arrow(ax, p1, p2, color=GRAY, lw=1.8, style="-|>"):
    ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle=style, mutation_scale=16, color=color, lw=lw))


# ---------- 그림 1. 모니터링 시스템 구성 ----------
def fig1():
    fig, ax = plt.subplots(figsize=(9.5, 5.0)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 6)
    ax.text(0, 5.95, "충주바이오그린수소충전소 실시간 전력 모니터링 시스템 구성", fontsize=13, fontweight="bold", color=NAVY, va="top")
    ax.text(0.1, 5.05, "4대 핵심 설비", fontsize=10.5, color=GRAY)
    eq = ["수소 압축기", "압축기용 냉각기", "충전기", "충전기용 냉각기(칠러)"]
    for i, t in enumerate(eq):
        box(ax, 0.1, 3.95 - i * 1.0, 2.2, 0.75, t, fs=9.5)
        arrow(ax, (2.32, 4.32 - i * 1.0), (2.95, 2.7 + (1.5 - i) * 0.3))
    box(ax, 3.0, 1.5, 1.8, 2.4, "메인 분전반\n+\n설비별\n적산전력계", fc="#eef3f9", fs=10.5, bold=True)
    arrow(ax, (4.82, 2.7), (5.35, 2.7))
    box(ax, 5.4, 1.5, 1.8, 2.4, "실시간 전력\n모니터링\n시스템", fc="#eef3f9", fs=10.5, bold=True)
    ax.text(6.3, 1.2, "수집: 전압·전류·전력량\n역률·온도·가동상태 (시간 단위)", fontsize=8.8, color=GRAY, ha="center", va="top", linespacing=1.5)
    arrow(ax, (7.22, 2.7), (7.6, 2.7))
    ax.text(7.65, 5.35, "분석·활용", fontsize=10.5, color=GRAY)
    outs = ["에너지 피크 관리", "부하 구간별 가동률\n(경·중·대부하)", "설비별 전력 사용량", "기간별 소모 트렌드\n(일·주·월·연)", "SEC(kWh/kgH₂)\n월별 추적"]
    for i, t in enumerate(outs):
        box(ax, 7.65, 4.3 - i * 0.98, 2.3, 0.72, t, fs=8.8, ec=BLUE)
    save(fig, "fig1_monitoring_system.png")


# ---------- 그림 2. 공정별 kWh 비중 vs SEC 비중 ----------
def fig2():
    fig, ax = plt.subplots(figsize=(8.5, 3.9))
    rows = [("에너지 소비량(kWh) 비중", [("생산", 51.0), ("압축", 36.3), ("충전", 12.7)]),
            ("단위 수소당 에너지(SEC) 비중", [("생산", 35.9), ("압축", 28.9), ("충전", 35.2)])]
    sec_avg = {"생산": 7.2, "압축": 5.8, "충전": 9.1}
    for yi, (label, vals) in zip([1, 0], rows):
        left = 0
        for name, v in vals:
            ax.barh(yi, v, left=left, color=PROC[name], height=0.55, edgecolor="white")
            tc = "white" if name != "압축" else DARK
            t = f"{name} {v:.1f}%" + (f"\n(평균 {sec_avg[name]} kWh/kg)" if yi == 0 else "")
            ax.text(left + v / 2, yi, t, ha="center", va="center", color=tc, fontsize=10, fontweight="bold")
            left += v
    ax.set_yticks([1, 0]); ax.set_yticklabels([r[0] for r in rows], fontsize=11)
    ax.set_xlim(0, 100); ax.set_xlabel("비중 (%)"); ax.set_ylim(-1.3, 1.9)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.text(0, 1.75, "공정별 에너지 소비 비중 vs SEC 비중 (43주 유효 데이터)", fontsize=13, fontweight="bold", color=NAVY, va="center")
    ax.annotate("충전 공정: 소비량 비중 12.7% → SEC 비중 35.2%\n(약 2.8배 증가, 핵심 비효율 구간)",
                xy=(82, -0.3), xytext=(30, -0.9), fontsize=9.5, color=ORANGE, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=ORANGE, lw=1.5), va="center")
    save(fig, "fig2_process_share.png")


# ---------- 그림 3. 설비별 누적 에너지 ----------
def fig3():
    items = [("저압압축기", 218216.5), ("압축기용 냉각기", 141545.9), ("메인 칠러", 118729.0),
             ("중압압축기", 111251.8), ("서브 칠러", 38640.6), ("고압압축기", 22804.3)]
    total = sum(v for _, v in items)
    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    names = [n for n, _ in items][::-1]; vals = [v for _, v in items][::-1]
    cols = [ORANGE if "메인" in n else (BLUE if "칠러" in n or "냉각" in n else NAVY) for n in names]
    ax.barh(names, vals, color=cols, height=0.62)
    for i, v in enumerate(vals):
        ax.text(v + 3000, i, f"{v:,.1f} kWh ({v / total * 100:.1f}%)", va="center", fontsize=9.5)
    ax.set_xlim(0, 330000); ax.set_xlabel("누적 에너지 사용량 (kWh)")
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
    ax.set_title("핵심 설비별 누적 에너지 사용량 (43주)", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=NAVY, label="압축기"), Patch(color=BLUE, label="냉각기·서브 칠러"),
                       Patch(color=ORANGE, label="메인 칠러(충전)")], loc="lower right", frameon=False, fontsize=9.5)
    save(fig, "fig3_equipment_kwh.png")


# ---------- 그림 4. 설비별 SEC ----------
def fig4():
    items = [("고압압축기", 1.4), ("저압압축기", 1.6), ("중압압축기", 0.9), ("압축기용\n냉각기", 1.1), ("메인 칠러\n(충전)", 7.9)]
    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    cols = [NAVY] * 3 + [BLUE, ORANGE]
    ax.bar([n for n, _ in items], [v for _, v in items], color=cols, width=0.58)
    for i, (_, v) in enumerate(items):
        ax.text(i, v + 0.2, f"{v}", ha="center", fontsize=11, fontweight="bold")
    ax.plot([4], [15.05], marker="v", color=ORANGE, ms=9)
    ax.vlines(4, 8.9, 15.05, color=ORANGE, lw=1.6, ls="--")
    ax.text(4.12, 11.5, "최대\n14.7~15.4", fontsize=9.5, color=ORANGE, va="center")
    ax.annotate("설비 전체 SEC의 60.1% 차지\n→ 핵심 병목 설비", xy=(3.62, 7.2), xytext=(1.0, 10.8),
                fontsize=10.5, color=ORANGE, fontweight="bold", arrowprops=dict(arrowstyle="->", color=ORANGE, lw=1.5))
    ax.text(-0.4, 3.0, "압축 공정 설비 분포 범위 0.6~2.5 kWh/kg", fontsize=9, color=GRAY)
    ax.set_ylim(0, 16.5); ax.set_ylabel("평균 SEC (kWh/kgH₂)")
    ax.set_title("핵심 설비별 단위 수소 처리량당 에너지 사용량(SEC)", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    save(fig, "fig4_equipment_sec.png")


# ---------- 그림 5. 메인 칠러 이중 전력 구조 ----------
def fig5():
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    ax.bar(["대기 모드\n(Standby, 충전 차량 없음)", "충전 모드\n(Active, 차량 충전 시)"], [14.9, 33.6],
           color=[GRAY, ORANGE], width=0.5)
    ax.text(0, 14.9 + 0.8, "평균 14.9 kWh", ha="center", fontsize=12, fontweight="bold")
    ax.text(1, 33.6 + 0.8, "최대 33.6 kWh", ha="center", fontsize=12, fontweight="bold")
    ax.set_ylim(0, 40); ax.set_ylabel("메인 칠러 전력 (kWh)")
    ax.text(0, 7.5, "사전 냉각(-40°C) 상시 유지\nSAE J2601 (T40)", ha="center", va="center", color="white", fontsize=10)
    ax.set_title("충전 공정 메인 칠러의 이중 전력 소비 구조 (650시간 분석)", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    save(fig, "fig5_chiller_dual_mode.png")


# ---------- 그림 6. 운영 모델 비교 (Box plot) ----------
def fig6():
    stats = [dict(label="(A) 도심형 전용 충전소\n(연속 충전)", q1=6.8, med=7.5, q3=8.7, iqr=1.9),
             dict(label="(B) 광역 마더 스테이션\n(간헐 충전)", q1=9.3, med=10.4, q3=12.8, iqr=3.4)]
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(9, 4.4), gridspec_kw=dict(width_ratios=[1.7, 1]))
    for i, s in enumerate(stats, 1):
        c = BLUE if i == 1 else ORANGE
        ax.add_patch(plt.Rectangle((i - 0.25, s["q1"]), 0.5, s["q3"] - s["q1"], fc=c, alpha=0.55, ec=c, lw=1.8))
        ax.hlines(s["med"], i - 0.25, i + 0.25, color=DARK, lw=2.6)
        ax.text(i + 0.32, s["med"], f"중앙값 {s['med']}", va="center", fontsize=10, fontweight="bold")
        ax.text(i + 0.32, s["q3"], f"Q3 {s['q3']}", va="center", fontsize=9, color=GRAY)
        ax.text(i + 0.32, s["q1"], f"Q1 {s['q1']}", va="center", fontsize=9, color=GRAY)
        ax.text(i, s["q3"] + 0.9, f"IQR {s['iqr']}", ha="center", fontsize=10.5, fontweight="bold", color=c)
    ax.set_xlim(0.4, 2.9); ax.set_ylim(4, 16)
    ax.set_xticks([1, 2]); ax.set_xticklabels([s["label"] for s in stats])
    ax.set_ylabel("차량 충전 공정 SEC (kWh/kgH₂)")
    ax.set_title("운영 모델별 충전 공정 SEC 분포 (IQR)", loc="left", fontsize=12.5, fontweight="bold", color=NAVY, pad=12)
    ax2.pie([85, 15], colors=[GRAY, ORANGE], startangle=90, counterclock=False, wedgeprops=dict(width=0.38, edgecolor="white"),
            labels=["튜브트레일러\n출하 85.0%", "차량 충전\n15.0%"], textprops=dict(fontsize=9.5))
    ax2.set_title("(B) 출하 구성", fontsize=11, color=NAVY, fontweight="bold")
    fig.text(0.01, -0.03, "※ 박스: Q1~Q3, 굵은선: 중앙값 (보고서 기재 통계값 기준, 수염·이상치 미표기)", fontsize=8.5, color=GRAY)
    fig.tight_layout()
    save(fig, "fig6_boxplot_models.png")


# ---------- 그림 7. 베이스라인 및 목표 ----------
def fig7():
    fig, ax = plt.subplots(figsize=(6.8, 4.3))
    ax.bar(["베이스라인\n(상용 22개소 통계)", "최종 목표\n(20% 절감)"], [9.5, 7.6], color=[GRAY, NAVY], width=0.5)
    ax.text(0, 9.5 + 0.15, "9.5", ha="center", fontsize=14, fontweight="bold")
    ax.text(1, 7.6 + 0.15, "7.6", ha="center", fontsize=14, fontweight="bold")
    arrow(ax, (0.3, 9.3), (0.7, 7.85), color=ORANGE, lw=2)
    ax.text(0.5, 8.9, "-20%\n(-1.9)", ha="center", fontsize=11, color=ORANGE, fontweight="bold")
    ax.set_ylim(0, 11.2); ax.set_ylabel("에너지 효율 (kWh/kgH₂)")
    ax.set_title("사업대상지 에너지 효율 베이스라인 및 목표", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    save(fig, "fig7_baseline_target.png")


# ---------- 그림 8. 3단계 연구개발 흐름 ----------
def fig8():
    fig, ax = plt.subplots(figsize=(10, 4.6)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 5)
    ax.text(0, 4.85, "다음 단계 연구개발 내용: 3단계 추진 체계", fontsize=13, fontweight="bold", color=NAVY, va="top")
    heads = ["1단계\n운영 패턴 정밀 분석", "2단계\n차등 운영 시나리오 수립", "3단계\n현장 적용 및 효과 검증"]
    bodies = ["· 원시 데이터 전처리\n  (노이즈 제거·결측 보정)\n· 월·일·요일·시간대별\n  충전 수요 패턴(히트맵)\n· 피크/오프피크 식별\n  (예: 11:00~14:00)\n· 유휴시간-대기전력\n  상관관계 규명",
              "· 피크/오프피크별\n  압축기 가동률·부하 분산\n· 설비별 절전 모드 진입\n  로직 설계\n  (Operation→Standby\n   →Deep Sleep)\n· 소비량 시뮬레이션,\n  절감률·절감량 확정",
              "· 고효율 부품·설비 적용\n· Deep Sleep 로직 이식\n· Before/After 비교\n  (대기전력 1,800→450 kWh)\n· 외부전문가·KTC 입회\n· ROI 평가·표준 운영\n  매뉴얼 작성"]
    xs = [0.1, 3.5, 6.9]
    for i in range(3):
        box(ax, xs[i], 3.55, 3.0, 0.95, heads[i], fc=NAVY, ec=NAVY, tc="white", fs=11, bold=True)
        box(ax, xs[i], 0.2, 3.0, 3.2, bodies[i], fc="#f4f7fb", ec=LIGHT, fs=9.5)
        if i < 2:
            arrow(ax, (xs[i] + 3.02, 4.02), (xs[i + 1] - 0.02, 4.02), color=ORANGE, lw=2.4)
    save(fig, "fig8_three_stage_plan.png")


# ---------- 그림 9. 절전 모드 상태 전이 ----------
def fig9():
    fig, ax = plt.subplots(figsize=(10, 4.6)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 5.0)
    ax.text(0, 4.95, "설비 유휴시간 기반 단계별 절전 모드 상태 전이(State Transition)", fontsize=13, fontweight="bold", color=NAVY, va="top")
    st = [("정상 운전\nOperation", "100% 부하 정상 가동", "절감 0%", NAVY, "white"),
          ("대기 모드\nStandby", "보조설비(냉각기 등) 저속 운전\n즉시 복귀 가능", "에너지 약 30% 절감", BLUE, "white"),
          ("절전 모드\nDeep Sleep", "필수 모니터링 전원 외\n설비 가동 중지", "에너지 약 90% 절감", ORANGE, "white")]
    xs = [0.2, 3.7, 7.2]
    for (t, d, s, fc, tc), x in zip(st, xs):
        box(ax, x, 1.85, 2.6, 1.3, t, fc=fc, ec=fc, tc=tc, fs=12, bold=True)
        ax.text(x + 1.3, 1.45, d, ha="center", va="top", fontsize=9.5, color=DARK, linespacing=1.5)
        ax.text(x + 1.3, 0.35, s, ha="center", va="center", fontsize=10.5, color=fc if fc != BLUE else NAVY, fontweight="bold")
    for i, lab in enumerate(["미가동\n10분 지속", "미가동\n60분 지속"]):
        arrow(ax, (xs[i] + 2.62, 2.7), (xs[i + 1] - 0.02, 2.7), color=DARK, lw=2)
        ax.text(xs[i] + 3.15, 2.95, lab, ha="center", fontsize=8.5, color=DARK)
    ax.annotate("", xy=(1.5, 3.2), xytext=(8.5, 3.2),
                arrowprops=dict(arrowstyle="-|>", color=GRAY, lw=1.4, connectionstyle="arc3,rad=0.15"))
    ax.text(5.0, 4.15, "충전 수요 발생 시 즉시 복귀", fontsize=9.5, color=GRAY, ha="center")
    save(fig, "fig9_state_transition.png")


# ---------- 그림 10. 대기전력 저감 목표 ----------
def fig10():
    fig, ax = plt.subplots(figsize=(7.2, 4.3))
    ax.bar(["현행 운영\n(시뮬레이션)", "최적화 적용 후\n(목표)"], [1800, 450], color=[GRAY, NAVY], width=0.5)
    ax.text(0, 1860, "1,800 kWh", ha="center", fontsize=13, fontweight="bold")
    ax.text(1, 510, "450 kWh", ha="center", fontsize=13, fontweight="bold")
    arrow(ax, (0.3, 1650), (0.7, 600), color=ORANGE, lw=2)
    ax.text(0.68, 1500, "대기전력\n-75%", ha="center", fontsize=12, color=ORANGE, fontweight="bold")
    ax.text(0.5, 2250, "전체 에너지 소비량 약 18.5% 이상 절감 예상", ha="center", fontsize=10.5, color=NAVY,
            bbox=dict(boxstyle="round,pad=0.4", fc="#eef3f9", ec=LIGHT))
    ax.set_ylim(0, 2500); ax.set_ylabel("대기 전력 낭비량 (kWh)")
    ax.set_title("부하 반응형 차등 제어에 의한 대기전력 저감 목표", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    save(fig, "fig10_standby_reduction.png")


# ---------- 그림 11. 투트랙 전략 ----------
def fig11():
    fig, ax = plt.subplots(figsize=(9.5, 4.6)); ax.axis("off"); ax.set_xlim(0, 10); ax.set_ylim(0, 5)
    ax.text(0, 4.9, "H/W·S/W 투트랙(Two-Track) 융합 연구개발 추진전략", fontsize=13, fontweight="bold", color=NAVY, va="top")
    box(ax, 0.1, 2.9, 4.4, 1.2, "Track 01  하드웨어 부문\n(참여기관 중심)", fc=BLUE, ec=BLUE, tc="white", fs=11.5, bold=True)
    box(ax, 0.1, 1.55, 4.4, 1.15, "노후 설비 교체\n고효율 핵심 부품·설비 현장 도입", fc="#f4f7fb", ec=LIGHT, fs=10)
    box(ax, 5.5, 2.9, 4.4, 1.2, "Track 02  소프트웨어·운영 부문\n(주관: 고등기술연구원)", fc=NAVY, ec=NAVY, tc="white", fs=11.5, bold=True)
    box(ax, 5.5, 1.55, 4.4, 1.15, "피크/오프피크 식별\n부하 반응형 차등 제어(Standby·Deep Sleep)", fc="#f4f7fb", ec=LIGHT, fs=10)
    arrow(ax, (2.3, 1.55), (4.0, 0.95), color=ORANGE, lw=2); arrow(ax, (7.7, 1.55), (6.0, 0.95), color=ORANGE, lw=2)
    box(ax, 2.0, 0.05, 6.0, 0.9, "융합 → 에너지 소비량 20% 절감 (9.5 → 7.6 kWh/kgH₂)\nKTC 등 공인기관 입회평가로 검증", fc=ORANGE, ec=ORANGE, tc="white", fs=10.5, bold=True)
    save(fig, "fig11_two_track_strategy.png")


for fn in (fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, fig9, fig10, fig11):
    fn()
print("done")
