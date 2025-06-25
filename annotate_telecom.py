import json
import re
from pathlib import Path

# Define file paths
INPUT_PATH = Path('data') / '抽样1000.json'
OUTPUT_PATH = Path('data') / '抽样1000_已标注.json'

# Heuristic keyword lists
HIGH_RELEVANCE_KEYWORDS = [
    # China Telecom core service keywords (Chinese & English)
    '电信', '通信', '网络', '基站', '5g', '4g', '光纤', '光缆', '宽带',
    '云', '云计算', '数据中心', 'IDC', '物联网', 'IoT', 'ICT', '互联网',
    '数字化', '信息化', '信息系统', '大数据', 'CDN', '服务器', '交换机', '路由器','智慧','AI','人工智能','图像识别','ai',
]

MEDIUM_RELEVANCE_KEYWORDS = [
    '智能', '监控', '软件', '系统', '平台', '数据', '电子', '自动化', '集成','运维','网络'
]

HIGH_POTENTIAL_KEYWORDS = [
    # New construction for school, hospital, govt building
    '新建教学楼', '新建医院', '新建政府大楼',
    # 单独出现关键词也可能代表潜力大
    '教学楼', '医院', '政府大楼', '政务服务中心', '办公大楼', '校区', '大学', '中学', '小学', '幼儿园','工程','施工', '扩建'
]


def score_relevance(summary: str) -> int:
    """Return 3/2/1 indicating relevance to China Telecom business scope."""
    text = summary.lower()
    # High relevance if any high keyword present
    if any(k.lower() in text for k in HIGH_RELEVANCE_KEYWORDS):
        return 3
    # Medium relevance if any medium keyword present
    if any(k.lower() in text for k in MEDIUM_RELEVANCE_KEYWORDS):
        return 2
    return 1


def score_potential(summary: str) -> int:
    """Return 3/2/1 indicating commercial potential."""
    text = summary
    if any(k in text for k in HIGH_POTENTIAL_KEYWORDS):
        return 3
    # 设备采购、系统集成等一般认为中等潜力
    if re.search(r'(采购|改造|升级|扩容)', text):
        return 2
    return 1


def main():
    # Load original data
    with INPUT_PATH.open('r', encoding='utf-8') as f:
        records = json.load(f)

    # Annotate each record
    for item in records:
        summary = item.get('公告摘要', '')
        rel_score = score_relevance(summary)
        pot_score = score_potential(summary)
        item['相关性标注1低2中3高'] = rel_score
        item['商业潜力标注1低2中3高'] = pot_score

    # Save annotated data
    with OUTPUT_PATH.open('w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    print(f'Annotated {len(records)} records. Saved to {OUTPUT_PATH}')


if __name__ == '__main__':
    main() 