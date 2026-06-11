"""
ai_comment.py — 공장 랭킹 데이터 기반 코멘트 생성
ANTHROPIC_API_KEY가 st.secrets 또는 환경변수에 있으면 Claude API 사용,
없으면 규칙 기반 fallback 코멘트 생성.
"""
import json


def _trend_mark(trend: str) -> str:
    """추이 표시: ↑(악화) ↓(개선) →(보합) (데이터 부족)"""
    return {
        'up':   '↑ 악화',
        'down': '↓ 개선',
        'flat': '→ 보합',
        'new':  '(데이터 부족)',
    }.get(trend, '→ 보합')


def fallback_comment(data: dict) -> str:
    """Claude API 없을 때 기본 코멘트 생성"""
    lines = [f"📊 분석 기간: {data['period']} | 전체 {data['total_factories']}개 공장\n"]

    if data['top5_best']:
        lines.append("✅ 추천 공장 (불량률 우수)")
        for r in data['top5_best'][:3]:
            lines.append(f"  · {r['factory']} ({r['region']}) — {r['rate']}% {_trend_mark(r['trend'])}")

    if data['top5_worst']:
        lines.append("\n⚠️ 주의 공장 (불량률 높음)")
        for r in data['top5_worst'][:3]:
            lines.append(f"  · {r['factory']} ({r['region']}) — {r['rate']}% {_trend_mark(r['trend'])}")

    if data['notable_changes']:
        lines.append("\n🚨 급등 공장 (즉시 확인 필요)")
        for r in data['notable_changes']:
            lines.append(f"  · {r['factory']} ({r['region']}) — {r['rate']}% 급등")

    return "\n".join(lines)


def get_comment(ai_data: dict, api_key: str | None = None) -> str:
    """ANTHROPIC_API_KEY가 있으면 Claude로 코멘트 생성, 없으면 fallback"""
    if not api_key:
        return fallback_comment(ai_data)

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        prompt = f"""당신은 섬유·의류 품질관리 전문가입니다.
아래 공장별 불량률 데이터를 분석하여 한국어로 실무적인 코멘트를 작성해주세요.

분석 기간: {ai_data['period']}
전체 공장 수: {ai_data['total_factories']}개

[불량률 우수 공장 TOP5]
{json.dumps(ai_data['top5_best'], ensure_ascii=False, indent=2)}

[불량률 주의 공장 TOP5]
{json.dumps(ai_data['top5_worst'], ensure_ascii=False, indent=2)}

[급격한 변화 공장]
{json.dumps(ai_data['notable_changes'], ensure_ascii=False, indent=2)}

다음 형식으로 작성해주세요:
1. ✅ 추천 공장 (발주 우선 고려): 공장명, 지역, 불량률, 추천 이유를 2~3문장으로
2. ⚠️ 주의 공장 (개선 요청 또는 재검토): 공장명, 지역, 불량률, 주의 사항을 2~3문장으로
3. 📊 전체 요약: 전체적인 품질 수준 평가 1~2문장

간결하고 실무적으로 작성해주세요."""

        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text
    except Exception:
        return fallback_comment(ai_data)
