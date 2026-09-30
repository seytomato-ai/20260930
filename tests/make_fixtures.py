"""테스트용 '가상' 시간표·기존 기록 파일을 만든다 (실제 학생 정보 없음).  python3 tests/make_fixtures.py"""
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

NAMES = ['강가온','고나래','권다솜','김라온','김마루','남보라','노새봄','류아름','문여울','박은솔',
         '배하람','서나린','손다온','송미르','신바다','안소리','양하늘','오누리','유가람','윤별하',
         '이도담','이로운','임새론','장한결','전해솔','정다인','조은별','최푸름']
ELECT = [('역학과 에너지', '정우진'), ('식품과 영양', '서예린'), ('인공지능 기초', '김도현'), ('세계사', '오하린')]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '학생시간표'
head = ['번호', '이름'] + [d + str(p) for d in '월화수목금' for p in range(1, 8)]
ws.append(head)
common = [''] * len(head)
common[1] = '공통'
common[head.index('수1')] = '영어Ⅱ(한서준)'
common[head.index('수2')] = '문학(윤지아)'
common[head.index('수3')] = '기하(문지후)'
common[head.index('수4')] = '스포츠 생활2A(최민재)'
ws.append(common)
for i, n in enumerate(NAMES, 1):
    row = [i, n] + [''] * (len(head) - 2)
    for p in (5, 6, 7):
        s, t = ELECT[(i + p) % len(ELECT)]
        row[head.index('수%d' % p)] = f'{s}({t})' if p != 7 else f'{s}\n{t}'
    ws.append(row)
wb.save('tests/fixtures/fake-timetable.xlsx')

# 기존 누적 기록 파일 (사용자 양식과 같은 머리글)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '2학년 2학기'
ws.append(['학번/학급', '이름', '날짜', '과목', '교시', '출결 종류', '담당 교사', '요청중/완료', '전달 사항'])
ws.append(['2-3', '권다솜', '2026.09.16', '세계사', '5교시', None, '오하린', '완료', '교시 마감 부탁드립니다.'])
ws.append(['2-3', '권다솜', '2026.09.23', '영어Ⅱ', '1교시', None, '한서준', '요청중', '교시 마감 부탁드립니다.'])
dv = DataValidation(type='list', formula1='"요청중,완료"', allow_blank=True)
ws.add_data_validation(dv)
dv.add('H2:H934')
ws.auto_filter.ref = 'A1:I934'
wb.save('tests/fixtures/fake-existing.xlsx')
print('ok')
