"""
구글 플레이 리뷰 분석기 — Streamlit 웹 버전 v2
VIC GAME STUDIOS | 일본사업실 박경원
"""

import streamlit as st
import pandas as pd
import time, re, io
from datetime import datetime, date, timedelta

try:
    from google_play_scraper import reviews, Sort, app as gp_app
    HAS_SCRAPER = True
except ImportError:
    HAS_SCRAPER = False

try:
    import requests as _requests
    HAS_APP_STORE = True
except ImportError:
    HAS_APP_STORE = False

try:
    from wordcloud import WordCloud
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    HAS_WORDCLOUD = True
except ImportError:
    HAS_WORDCLOUD = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, PieChart, LineChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

C_DARK='1A1A2E'; C_MID='16213E'; C_ACCENT='7B2FBE'; C_GOLD='F5A623'
C_LIGHT='F3EEFF'; C_WHITE='FFFFFF'; C_GRAY='CCCCCC'; C_GREEN='27AE60'
C_LGREEN='E8F5E9'; C_ORANGE='E67E22'; C_RED='C0392B'; C_LRED='FFEBEE'
C_YELLOW='FFF9C4'; C_BLUE='2980B9'

REGIONS = {
    'KR': ('ko', 'kr'),
    'JP': ('ja', 'jp'),
    'US': ('en', 'us'),
    'TW': ('zh_TW', 'tw'),
    'GB': ('en', 'gb'),
}

# 앱스토어 국가 코드
AS_REGIONS = {
    'KR': 'kr',
    'JP': 'jp',
    'US': 'us',
    'TW': 'tw',
    'GB': 'gb',
}

# 스팀 언어 코드
STEAM_LANGS = {
    'KR': 'koreana',
    'JP': 'japanese',
    'US': 'english',
    'TW': 'tchinese',
    'GB': 'english',
}

I18N = {
    'KR': {
        'title':'🎮 구글 플레이 리뷰 분석기',
        'setting_title':'🌐 설정',
        'ui_lang_lbl':'UI 언어',
        'doc_lang_lbl':'문서 언어',
        'made_by':'Made by',
        'url_label':'📱 앱 URL 또는 패키지명',
        'url_ph':'https://play.google.com/store/apps/details?id=com.example.app',
        'mode_label':'📋 수집 방식',
        'mode_count':'개수 지정',
        'mode_period':'기간 지정',
        'count_label':'수집 개수',
        'from_label':'시작일',
        'to_label':'종료일',
        'platform_label':'📱 플랫폼',
        'platform_gp':'🤖 구글 플레이',
        'platform_as':'🍎 앱스토어 (iOS)',
        'platform_st':'🎮 스팀 (Steam)',

        'steam_url_ph':'https://store.steampowered.com/app/1234567/게임명  또는  앱 ID(숫자)',
        'err_no_steamid':'올바른 스팀 URL 또는 앱 ID(숫자)를 입력해주세요.',
        'steam_lang_label':'🌐 리뷰 언어',
        'steam_notice':'🎮 스팀 모드 — Steam 공식 API 사용 (무료, 키 불필요)\n\n💡 이론상 무제한 수집 가능 (배치당 100건)',
        'metric_rec':'👍 추천',
        'metric_norec':'👎 비추천',
        'steam_rec':'👍 추천','steam_norec':'👎 비추천',
        'steam_playtime':'플레이 시간(분)',
        'c_playtime':'플레이시간(분)',
        'region_label':'🌏 수집 국가',
        'appstore_url_ph':'https://apps.apple.com/kr/app/앱이름/id123456789  또는  앱 ID(숫자)',
        'err_no_appstore':'app-store-scraper가 설치되지 않았습니다.',
        'err_no_appid':'올바른 앱스토어 URL 또는 앱 ID(숫자)를 입력해주세요.',
        'btn_start':'▶  분석 시작',
        'btn_dl':'📥 엑셀 다운로드',
        'btn_csv':'📄 CSV 다운로드',
        'result_title':'📊 분석 결과 요약',
        'metric_total':'총 리뷰',
        'metric_avg':'평균 평점',
        'metric_pos':'긍정(4~5★)',
        'metric_neg':'부정(1~2★)',
        'unit_count':'건',
        'err_no_id':'올바른 앱 URL 또는 패키지명을 입력해주세요.',
        'err_date':'종료일이 시작일보다 빠릅니다.',
        'err_no_pkg':'google-play-scraper가 설치되지 않았습니다.',
        'err_no_data':'수집된 리뷰가 없습니다.',
        'log_app':'📱 앱명: {} | ⭐ 평점: {} | 💬 총 리뷰: {:,}',
        'log_collect':'🔄 수집 중... {:,}개 (배치 {}회, 최신: {})',
        'log_done':'✅ {:,}개 수집 완료',
        'log_short':'⚠️ 목표 {:,}개 중 {:,}개만 수집됨 (API 제한 또는 리뷰 부족)',
        'log_excel':'📊 엑셀 생성 중...',
        'log_finish':'🎉 분석 완료!',
        'prog_collect':'수집 중...',
        'prog_excel':'엑셀 생성 중...',
        'prog_done':'완료!',
        'retry_msg':'🔁 재시도 {}/3... ({})',
        'fail_msg':'❌ 수집 실패',
        'region_KR':'KR (한국)',
        'region_JP':'JP (日本)',
        'region_US':'US (미국)',
        'region_TW':'TW (대만)',
        'region_GB':'GB (영국)',
        'sh_dash':'📊 대시보드','sh_op':'🗣️ 여론분석',
        'sh_raw':'📋 전체 리뷰','sh_stat':'📈 통계',
        'c_id':'리뷰ID','c_user':'사용자','c_score':'평점','c_content':'내용',
        'c_date':'작성일','c_month':'작성월','c_like':'좋아요',
        'c_reply':'개발사답변','c_rdate':'답변일','c_ver':'앱버전',
        'dash_ttl':'🎮  구글 플레이 리뷰 분석 대시보드',
        'op_ttl':'🗣️  키워드 여론 분석 리포트',
        'stat_ttl':'📈 상세 통계 분석',
        'kpi_avg':'⭐ 평균 평점','kpi_total':'📝 총 리뷰',
        'kpi_pos':'😊 긍정(4~5★)','kpi_neu':'😐 중립(3★)',
        'kpi_neg':'😡 부정(1~2★)','kpi_rep':'💬 개발사답변',
        'ch_dist':'📊 평점별 리뷰 분포',
        'ch_pie':'🥧 긍정/중립/부정 비율',
        'ch_trend':'📈 월별 리뷰 수 & 평균 평점 추이',
        'ch_top10':'👍 좋아요 Top 10 리뷰',
        'ins_ttl':'💡 종합 인사이트 — 한눈에 보기',
        'ins_cat':'구분','ins_eval':'평가',
        'kw_ttl':'🔍 키워드별 상세 여론 분석',
        'neg_ch':'📊 부정 키워드 빈도','pos_ch':'📊 긍정 키워드 빈도',
        'col_rank':'순위','col_kw':'키워드','col_sent':'감성',
        'col_cnt':'언급량','col_sum':'핵심 요약','col_ex':'대표 리뷰 예시',
        'neg_lbl':'🔴 부정','pos_lbl':'🟢 긍정',
        'stat_sc':'★ 평점별 통계','stat_mo':'📅 월별 추이',
        'hd_cnt':'리뷰 수','hd_ratio':'비율(%)','hd_alike':'평균 좋아요',
        'hd_mlike':'최대 좋아요','hd_rcnt':'답변 수','hd_rrate':'답변율(%)',
        'hd_avg':'평균 평점','hd_month':'월','hd_date':'날짜',
        'pie_pos':'긍정(4~5★)','pie_neu':'중립(3★)','pie_neg':'부정(1~2★)',
        'rev_cnt':'리뷰 수','avg_sc':'평균평점',
        'sum_pfx':'📌  종합 : ',
        'tone_pos':'긍정적','tone_mix':'혼재','tone_neg':'부정적 여론 우세',
        'unit_reviews':'건',
        'dash_sub_fmt':'  {} ~ {}   |   {:,}건   |   {}',
        'op_sub_fmt':'  {:,}건  |  긍정 {:.1f}%  /  부정 {:.1f}%  |  {}',
        'kw_neg_sum':'부정 리뷰({:,}건) 중 {:,}건에서 언급. {} 관련 불만이 주요 이슈.',
        'kw_pos_sum':'긍정 리뷰({:,}건) 중 {:,}건에서 언급. {} 관련 만족도가 높음.',
        'lv_very_many':'매우 많음','lv_many':'많음','lv_normal':'보통','lv_few':'적음',
        'mood_pos':'전반적으로 민심이 우호적이며 긍정 여론({:.1f}%)이 우세',
        'mood_mix':'긍정({:.1f}%)과 부정({:.1f}%) 여론이 팽팽하게 혼재',
        'mood_neg':'부정 여론({:.1f}%)이 우세하며 유저 불만이 높은 상태',
        'one_line_fmt':'{} 등 콘텐츠 만족도는 높으나, {} 관련 불만이 지속 제기되고 있음. 평균 평점 {:.2f}점 — {}.',
        'rank_suffix':'위',
        'neg_kw':{
            '버그·오류':['버그','오류','에러','오작동','먹통','안됨','안돼'],
            '최적화·렉·발열':['렉','버벅','최적화','발열','느려','프레임','끊김'],
            '튕김·종료':['튕기','종료','팅기','강제종료','꺼짐'],
            '과금·뽑기':['과금','뽑기','확률','가챠','현질','비싸','천장'],
            '밸런스':['밸런스','너프','강캐','약캐','사기'],
            '조작·UI':['조작','ui','버튼','불편','작아','안눌'],
            '업데이트·운영':['업데이트','운영','공지','패치','방치','최악'],
            '스토리·콘텐츠':['스토리','콘텐츠','스킵','반복','지루','노잼'],
        },
        'pos_kw':{
            '그래픽·아트':['그래픽','그림','아트','예쁘','이쁘','퀄리티','일러스트'],
            '캐릭터':['캐릭터','캐릭','귀엽','매력','최애'],
            '스토리·세계관':['스토리','세계관','설정','몰입','감동','흥미'],
            '게임성·전투':['전투','전략','재밌','꿀잼','갓겜','중독','타격'],
            '과금·운영':['무과금','무료','보상','이벤트','관대','넉넉'],
        },
    },
    'JP': {
        'title':'🎮 Google Play レビュー分析ツール',
        'setting_title':'🌐 設定',
        'ui_lang_lbl':'UI言語',
        'doc_lang_lbl':'文書言語',
        'made_by':'Made by',
        'url_label':'📱 アプリURL またはパッケージ名',
        'url_ph':'https://play.google.com/store/apps/details?id=com.example.app',
        'mode_label':'📋 収集方法',
        'mode_count':'件数指定',
        'mode_period':'期間指定',
        'count_label':'収集件数',
        'from_label':'開始日',
        'to_label':'終了日',
        'platform_label':'📱 プラットフォーム',
        'platform_gp':'🤖 Google Play',
        'platform_as':'🍎 App Store (iOS)',
        'platform_st':'🎮 Steam',

        'steam_url_ph':'https://store.steampowered.com/app/1234567/ゲーム名  または  アプリID(数字)',
        'err_no_steamid':'正しいSteam URLまたはアプリID(数字)を入力してください。',
        'steam_lang_label':'🌐 レビュー言語',
        'steam_notice':'🎮 Steamモード — Steam公式API使用 (無料、キー不要)\n\n💡 理論上無制限収集可能 (バッチあたり100件)',
        'metric_rec':'👍 推薦',
        'metric_norec':'👎 非推薦',
        'steam_rec':'👍 推薦','steam_norec':'👎 非推薦',
        'steam_playtime':'プレイ時間(分)',
        'c_playtime':'プレイ時間(分)',
        'region_label':'🌏 収集国',
        'appstore_url_ph':'https://apps.apple.com/jp/app/アプリ名/id123456789  または  アプリID(数字)',
        'err_no_appstore':'app-store-scraperがインストールされていません。',
        'err_no_appid':'正しいApp StoreのURLまたはアプリID(数字)を入力してください。',
        'btn_start':'▶  分析開始',
        'btn_dl':'📥 Excelダウンロード',
        'btn_csv':'📄 CSVダウンロード',
        'result_title':'📊 分析結果サマリー',
        'metric_total':'総レビュー',
        'metric_avg':'平均評価',
        'metric_pos':'肯定(4~5★)',
        'metric_neg':'否定(1~2★)',
        'unit_count':'件',
        'err_no_id':'正しいアプリURLまたはパッケージ名を入力してください。',
        'err_date':'終了日が開始日より前になっています。',
        'err_no_pkg':'google-play-scraperがインストールされていません。',
        'err_no_data':'収集されたレビューがありません。',
        'log_app':'📱 アプリ名: {} | ⭐ 評価: {} | 💬 総レビュー: {:,}',
        'log_collect':'🔄 収集中... {:,}件 (バッチ{}回、最新: {})',
        'log_done':'✅ {:,}件収集完了',
        'log_short':'⚠️ 目標{:,}件中{:,}件のみ収集 (API制限またはレビュー不足)',
        'log_excel':'📊 Excel生成中...',
        'log_finish':'🎉 分析完了！',
        'prog_collect':'収集中...',
        'prog_excel':'Excel生成中...',
        'prog_done':'完了！',
        'retry_msg':'🔁 リトライ {}/3... ({})',
        'fail_msg':'❌ 収集失敗',
        'region_KR':'KR (韓国)',
        'region_JP':'JP (日本)',
        'region_US':'US (アメリカ)',
        'region_TW':'TW (台湾)',
        'region_GB':'GB (イギリス)',
        'sh_dash':'📊 ダッシュボード','sh_op':'🗣️ 世論分析',
        'sh_raw':'📋 全レビュー','sh_stat':'📈 統計',
        'c_id':'レビューID','c_user':'ユーザー','c_score':'評価','c_content':'内容',
        'c_date':'投稿日','c_month':'投稿月','c_like':'いいね',
        'c_reply':'デベロッパー返信','c_rdate':'返信日','c_ver':'バージョン',
        'dash_ttl':'🎮  Google Play レビュー分析ダッシュボード',
        'op_ttl':'🗣️  キーワード世論分析レポート',
        'stat_ttl':'📈 詳細統計分析',
        'kpi_avg':'⭐ 平均評価','kpi_total':'📝 総レビュー',
        'kpi_pos':'😊 肯定(4~5★)','kpi_neu':'😐 中立(3★)',
        'kpi_neg':'😡 否定(1~2★)','kpi_rep':'💬 返信あり',
        'ch_dist':'📊 評価別レビュー分布',
        'ch_pie':'🥧 肯定/中立/否定の割合',
        'ch_trend':'📈 月別レビュー数と平均評価の推移',
        'ch_top10':'👍 いいね Top 10 レビュー',
        'ins_ttl':'💡 総合インサイト — 一目でわかる',
        'ins_cat':'区分','ins_eval':'評価',
        'kw_ttl':'🔍 キーワード別詳細分析',
        'neg_ch':'📊 否定キーワード頻度','pos_ch':'📊 肯定キーワード頻度',
        'col_rank':'順位','col_kw':'キーワード','col_sent':'感情',
        'col_cnt':'言及数','col_sum':'要約','col_ex':'代表レビュー例',
        'neg_lbl':'🔴 否定','pos_lbl':'🟢 肯定',
        'stat_sc':'★ 評価別統計','stat_mo':'📅 月別推移',
        'hd_cnt':'レビュー数','hd_ratio':'割合(%)','hd_alike':'平均いいね',
        'hd_mlike':'最大いいね','hd_rcnt':'返信数','hd_rrate':'返信率(%)',
        'hd_avg':'平均評価','hd_month':'月','hd_date':'日付',
        'pie_pos':'肯定(4~5★)','pie_neu':'中立(3★)','pie_neg':'否定(1~2★)',
        'rev_cnt':'レビュー数','avg_sc':'平均評価',
        'sum_pfx':'📌  総合 : ',
        'tone_pos':'肯定的','tone_mix':'混在','tone_neg':'否定的世論優勢',
        'unit_reviews':'件',
        'dash_sub_fmt':'  {} ~ {}   |   {:,}件   |   {}',
        'op_sub_fmt':'  {:,}件  |  肯定 {:.1f}%  /  否定 {:.1f}%  |  {}',
        'kw_neg_sum':'否定レビュー({:,}件)中{:,}件で言及。{} 関連の不満が主要課題。',
        'kw_pos_sum':'肯定レビュー({:,}件)中{:,}件で言及。{} への満足度が高い。',
        'lv_very_many':'非常に多い','lv_many':'多い','lv_normal':'普通','lv_few':'少ない',
        'mood_pos':'全体的に民心は好意的で肯定的な世論({:.1f}%)が優勢',
        'mood_mix':'肯定({:.1f}%)と否定({:.1f}%)の世論が拮抗',
        'mood_neg':'否定的な世論({:.1f}%)が優勢でユーザーの不満が高い状態',
        'one_line_fmt':'{} 等のコンテンツ満足度は高いが、{} への不満が続いている。平均評価 {:.2f}点 — {}。',
        'rank_suffix':'位',
        'neg_kw':{
            'バグ・エラー':['バグ','エラー','不具合','フリーズ','落ちる'],
            '最適化・重さ・発熱':['重い','カクカク','最適化','発熱','遅い'],
            'クラッシュ・終了':['クラッシュ','強制終了','落ちる','止まる'],
            '課金・ガチャ':['課金','ガチャ','確率','天井','高い'],
            'バランス':['バランス','ナーフ','強すぎ','弱すぎ'],
            '操作性・UI':['操作','ui','ボタン','使いにくい'],
            '運営・アップデート':['運営','アップデート','放置','最悪'],
            'ストーリー・コンテンツ':['ストーリー','コンテンツ','つまらない'],
        },
        'pos_kw':{
            'グラフィック・アート':['グラフィック','絵','アート','綺麗','クオリティ'],
            'キャラクター':['キャラ','可愛い','魅力','推し'],
            'ストーリー・世界観':['ストーリー','世界観','設定','没入','感動'],
            'ゲーム性・戦闘':['戦闘','戦略','面白い','神ゲー','爽快'],
            '課金・運営':['無課金','無料','報酬','イベント'],
        },
    },
}

def _fill(c):  return PatternFill('solid', fgColor=c)
def _al(h='center',v='center',wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
def _bd():
    s=Side(style='thin',color=C_GRAY); return Border(left=s,right=s,top=s,bottom=s)
def _w(ws,r,c,val,bold=False,sz=10,fg='000000',bg=None,h='left',v='center',wrap=False,b=True,it=False):
    cell=ws.cell(row=r,column=c,value=val)
    cell.font=Font(bold=bold,size=sz,color=fg,name='Arial',italic=it)
    if bg: cell.fill=_fill(bg)
    cell.alignment=Alignment(horizontal=h,vertical=v,wrap_text=wrap)
    if b: cell.border=_bd()
    return cell
def _mw(ws,r1,c1,r2,c2,val,**kw):
    ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
    return _w(ws,r1,c1,val,**kw)
def _hr(ws,row,texts,bg=C_DARK,fg=C_WHITE,sz=10,height=22):
    ws.row_dimensions[row].height=height
    for col,txt in enumerate(texts,1): _w(ws,row,col,txt,bold=True,sz=sz,fg=fg,bg=bg,h='center')
def _sec(ws,row,c1,c2,title,bg=C_ACCENT,h=26):
    ws.merge_cells(start_row=row,start_column=c1,end_row=row,end_column=c2)
    _w(ws,row,c1,f'  {title}',bold=True,sz=13,fg=C_WHITE,bg=bg,h='left',v='center',b=False)
    ws.row_dimensions[row].height=h
def _cw(ws,m):
    for col,w in m.items(): ws.column_dimensions[col].width=w

def _hidden(ws, row, col, val):
    '''차트용 숨김 데이터 — 흰색 글씨로 안 보이게 처리'''
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(color='FFFFFF', size=9, name='Arial')
    return c

def evaluate_5_insights(df, lang):
    texts_neg=' '.join(df[df['평점']<=2]['내용'].dropna().tolist()).lower()
    texts_pos=' '.join(df[df['평점']>=4]['내용'].dropna().tolist()).lower()
    def count(text,kws): return sum(1 for k in kws if k in text)
    if lang=='KR':
        c_pos=count(texts_pos,['그래픽','스토리','원작','아트','세계관','캐릭터','퀄리티','감동'])
        c_neg=count(texts_neg,['스토리','콘텐츠','부실','빈약','스킵','노잼'])
        if c_pos>=5 and c_neg<3: ce='✅ 합격점 — 그래픽·스토리·IP 재현 만족도 높음'; cc=C_GREEN
        elif c_neg>=5: ce='❌ 부족 — 스토리·콘텐츠 관련 불만 다수'; cc=C_RED
        else: ce='⚠️ 보통 — 일부 만족, 개선 여지 있음'; cc=C_ORANGE
        t_neg=count(texts_neg,['렉','버그','최적화','발열','튕기','오류','사운드','끊김'])
        if t_neg>=10: te='❌ 심각한 수준 — 최적화·버그·사운드 불만 매우 많음'; tc=C_RED
        elif t_neg>=5: te='⚠️ 주의 필요 — 기술적 이슈 다수 보고됨'; tc=C_ORANGE
        else: te='✅ 양호 — 기술적 불만 적음'; tc=C_GREEN
        e_neg=count(texts_neg,['과금','뽑기','확률','재화','천장','현질','비싸'])
        if e_neg>=8: ee='⚠️ 강한 불만 — 뽑기 단가·재화 수급에 강한 불만'; ec=C_ORANGE
        elif e_neg>=4: ee='⚠️ 불만 있음 — 과금 구조 개선 요구 존재'; ec=C_ORANGE
        else: ee='✅ 양호 — 과금 관련 불만 적음'; ec=C_GREEN
        m_neg=count(texts_neg,['모바일','폰으로','조작','버튼','ui','터치'])
        if m_neg>=6: pe='📵 모바일보다 PC/콘솔 권장 여론 형성'; pc=C_BLUE
        elif m_neg>=3: pe='⚠️ 모바일 최적화 개선 요구'; pc=C_ORANGE
        else: pe='✅ 플랫폼 불만 적음'; pc=C_GREEN
        b_neg=count(texts_neg,['운영','공지','방치','노답','최악','환불'])
        if b_neg>=5: be='⚠️ 브랜드 불신 — 운영사 관련 부정 여론 존재'; bc=C_ORANGE
        elif count(texts_pos,['운영','감사','친절'])>=3: be='✅ 브랜드 신뢰 — 운영 관련 긍정 반응'; bc=C_GREEN
        else: be='😐 중립 — 브랜드 관련 언급 적음'; bc=C_BLUE
        return [('콘텐츠 (IP·그래픽·스토리)',ce,cc),('기술 (최적화·버그·사운드)',te,tc),
                ('경제 (뽑기·재화 수급)',ee,ec),('플랫폼',pe,pc),('브랜드',be,bc)]
    else:
        c_pos=count(texts_pos,['グラフィック','ストーリー','アート','世界観','キャラ','クオリティ'])
        c_neg=count(texts_neg,['ストーリー','コンテンツ','つまらない'])
        if c_pos>=5 and c_neg<3: ce='✅ 合格 — グラフィック・ストーリー・IP再現の満足度高い'; cc=C_GREEN
        elif c_neg>=5: ce='❌ 不足 — ストーリー・コンテンツへの不満多数'; cc=C_RED
        else: ce='⚠️ 普通 — 一部満足、改善余地あり'; cc=C_ORANGE
        t_neg=count(texts_neg,['重い','バグ','最適化','発熱','クラッシュ','エラー'])
        if t_neg>=10: te='❌ 深刻 — 最適化・バグ・サウンドへの不満が非常に多い'; tc=C_RED
        elif t_neg>=5: te='⚠️ 要注意 — 技術的問題の報告多数'; tc=C_ORANGE
        else: te='✅ 良好 — 技術的不満少ない'; tc=C_GREEN
        e_neg=count(texts_neg,['課金','ガチャ','確率','天井','高い'])
        if e_neg>=8: ee='⚠️ 強い不満 — ガチャ単価・資源供給への強い不満'; ec=C_ORANGE
        elif e_neg>=4: ee='⚠️ 不満あり — 課金構造の改善要求あり'; ec=C_ORANGE
        else: ee='✅ 良好 — 課金関連不満少ない'; ec=C_GREEN
        m_neg=count(texts_neg,['スマホ','モバイル','操作','ボタン','ui'])
        if m_neg>=6: pe='📵 モバイルよりPC/コンソール推奨の世論形成'; pc=C_BLUE
        elif m_neg>=3: pe='⚠️ モバイル最適化改善要求'; pc=C_ORANGE
        else: pe='✅ プラットフォーム不満少ない'; pc=C_GREEN
        b_neg=count(texts_neg,['運営','放置','最悪','返金'])
        if b_neg>=5: be='⚠️ ブランド不信 — 運営への否定的世論あり'; bc=C_ORANGE
        else: be='😐 中立 — ブランド関連言及少ない'; bc=C_BLUE
        return [('コンテンツ (IP・グラフィック・ストーリー)',ce,cc),('技術 (最適化・バグ・サウンド)',te,tc),
                ('経済 (ガチャ・資源供給)',ee,ec),('プラットフォーム',pe,pc),('ブランド',be,bc)]

def analyze_kw(df, t):
    # 여론분석 시 이상 리뷰 제외 (전체리뷰 시트에는 영향 없음)
    df_analysis, removed = filter_abnormal_reviews(df)
    if removed > 0:
        df = df_analysis  # 분석용 df만 교체
    # ── 평점 40% + sentiment_score 60% 조합 버킷 분류
    # 1단계: 간이 sentiment_score 사전 (버킷 분류용)
    _POS_KW = ['재밌','좋아','최고','갓겜','꿀잼','대박','짱','추천','만족','감동',
               '몰입','좋음','좋네','좋다','재미있','흥미','예쁘','퀄리티','굿',
               '굳','잼남','잼있','존잼','인생겜','명작','강추','중독','힐링']
    _NEG_KW = ['버그','최악','짜증','불편','렉','망겜','서운','아쉽','실망','후회',
               '제발','비싸','확률','과금','현질','뽑기','런함','접음','쓰레기',
               '사기','기만','방치','운영','환불','호구','흑우','도박','인플레',
               '이격','믿어본다','한번만','결국','역시나','기대이하','낙담','허탈']
    _NEGATION = ['없어요','없음','없다','없어','안 ','안됨','안돼','못 ','전혀','하나도']

    def _quick_score(text, score_base):
        '''평점 기반 점수(score_base)와 텍스트 분석 조합'''
        t_low = str(text).lower()
        txt_score = 0
        for k in _POS_KW:
            if k in t_low: txt_score += 1
        for k in _NEG_KW:
            if k in t_low:
                # 부정어 근처면 긍정으로
                idx = t_low.find(k)
                surr = t_low[max(0,idx-6):idx+len(k)+6]
                if any(n in surr for n in _NEGATION): txt_score += 1
                else: txt_score -= 1
        # 평점 40% + 텍스트 60% 조합
        # 평점 5→+2, 4→+1, 3→0, 2→-1, 1→-2
        rating_score = (score_base - 3)  # -2 ~ +2
        combined = rating_score * 0.4 + txt_score * 0.6
        return combined

    # 조합 점수 기반 버킷 분류
    all_texts = df['내용'].dropna()
    neg_indices = []
    pos_indices = []
    for idx, row in df.iterrows():
        if pd.isna(row['내용']) or str(row['내용']).strip() == '':
            continue
        score = _quick_score(row['내용'], row['평점'])
        if score < -0.3:
            neg_indices.append(idx)
        elif score > 0.3:
            pos_indices.append(idx)

    neg_tx = df.loc[neg_indices, '내용'].dropna()
    pos_tx = df.loc[pos_indices, '내용'].dropna()
    neg_total=len(neg_tx); pos_total=len(pos_tx)
    # ── 규칙 기반 감성분석 고도화 (#7)
    # 게임 슬랭 포함 확장 부정 표현 사전
    NEG_EXPR=[
        # 직접 불만
        '없애','별로','최악','짜증','불편','아쉽','문제','버그','오류',
        '싫','노잼','지루','힘들','망','안됨','안돼','못하','에러','튕',
        '렉','느려','발열','뻥','과금','현질','뽑기','비싸','천장','불만',
        '환불','삭제','망겜','없애셈','고쳐','해주세요','해줘요','개선해',
        'ㅡㅡ','ㅠ','ㅜ','갈증','낚이','실망','문의','뭡니까','말았다',
        # 게임 커뮤니티 슬랭 부정
        '런함','런했','접음','접었','꼬접','탈출','도망','지움','삭제함',
        '흑우','봉','호구','호갱','호구됨','봉됨',
        '없데이트','없뎃','노업','노업뎃',
        '수금','뽑아먹','등골','빨아먹','뜯어먹','갈취',
        '기싸움','통보','묵살','유기','방치','무시','불통','소통없',
        '조작겜','확률조작','운겜','운빨','사행성','도박','카지노',
        '섭종','폭망','말아먹','망했','나락','쓰레기','폐기','폐급',
        '뒤통수','통수','사기','기만','거짓말','약속안','약속어기',
        '무능','개판','엉망','최하','최저','ㄹㅈㄷ','ㅈ망','ㅈ같',
        '하지마','하지마세요','비추','비추천','추천금지',
        '열받','화남','뿔남','분노','짜증폭발','열이받',
        '돈낭비','시간낭비','후회','아깝','아까워','아깝다',
        # 추가 보완 키워드
        '서운','서운하','아쉽네','아쉬워','아쉽다','아쉬운',
        '제발','주십시오','주세요','해줬으면','해줬으면 좋',
        '이격','도배','인플레','픽업','픽뚫','천장없',
        '믿어본다','마지막','한번만','한 번만','한번더','한 번더',
        '기대이하','기대 이하','실망이','실망스','낙담','허탈',
        '결국','결국엔','역시나','역시','또다시','또 다시',
    ]
    # 게임 슬랭 포함 확장 긍정 표현 사전
    POS_EXPR=[
        # 기본 긍정
        '재밌','좋아','최고','훌륭','완벽','갓','꿀잼','대박','짱','추천',
        '만족','즐거','신나','감동','몰입','좋음','좋네','좋다','재미있',
        '흥미','멋지','예쁘','이쁘','퀄리티','매력',
        # 게임 커뮤니티 슬랭 긍정
        '갓겜','인생겜','명작','갓작','꿀','꿀템','꿀재','존잼','핵잼',
        '잼남','잼있','잼써','잼네','재밋','재밌네','재밌어','재밌다',
        '굳굳','굿굿','굿게임','굿겜','갓','레전드','레전','ㄹㅇ좋',
        '강추','강력추천','완전추천','진짜추천',
        '중독','빠져','못끊','계속하','계속 하','오래하',
        '잘만든','잘 만든','퀄높','퀄이높','퀄좋','퀄이좋',
        '재미짐','재미있음','재미있네','재미있어',
        '힐링','낭만','추억','감성','따뜻','포근',
    ]
    # 역접어 패턴 (뒤에 오는 감정이 최종 감정)
    REVERSAL_KW=['한데','지만','는데','근데','그러나','하지만','그런데',
                 '이지만','이긴','긴 하','긴하','이긴 하','이긴하']

    # 부정어 목록 (#5 부정어+키워드 조합 강화)
    NEGATION_WORDS = [
        '없어요','없음','없네요','없다','없는','없어','없고',
        '안 ','안되','안됨','안돼','못 ','못함','못해','못하',
        '전혀','하나도','거의','별로','노 ','노~','ㄴㄴ',
    ]

    def has_negation_near(text, keyword, window=8):
        '''키워드 앞뒤 window글자 내에 부정어가 있는지 확인'''
        idx = text.find(keyword)
        if idx == -1: return False
        surrounding = text[max(0, idx-window) : idx+len(keyword)+window]
        return any(nw in surrounding for nw in NEGATION_WORDS)

    def sentiment_score(text):
        '''감성 점수 계산: 양수=긍정, 음수=부정, 0=중립'''
        t_lower = text.lower()
        score = 0
        # 역접어 위치 감지
        reversal_pos = -1
        for rw in REVERSAL_KW:
            idx = t_lower.find(rw)
            if idx != -1:
                reversal_pos = idx
                break
        # 역접어 있으면 앞/뒤 분리
        if reversal_pos > 0:
            before = t_lower[:reversal_pos]
            after  = t_lower[reversal_pos:]
            for p in POS_EXPR:
                if p in before: score += 1
            for n in NEG_EXPR:
                # 부정어+부정키워드 조합이면 오히려 긍정
                if n in before:
                    if has_negation_near(before, n): score += 1
                    else: score -= 1
            for p in POS_EXPR:
                if p in after: score += 2
            for n in NEG_EXPR:
                if n in after:
                    if has_negation_near(after, n): score += 2
                    else: score -= 2
        else:
            for p in POS_EXPR:
                if p in t_lower: score += 2
            for n in NEG_EXPR:
                if n in t_lower:
                    # 부정어 + 부정키워드 = 긍정으로 재분류
                    if has_negation_near(t_lower, n): score += 2
                    else: score -= 2
        return score

    def cnt_neg(texts, kw_dict, used):
        res={}
        for lbl,kws in kw_dict.items():
            c=0; ex=[]
            for tx in texts:
                s=str(tx).lower()
                if any(k in s for k in kws):
                    c+=1
                    raw=str(tx).strip()
                    if len(ex)<3 and len(raw)>=10 and raw not in used:
                        ex.append(raw[:85]); used.add(raw)
            res[lbl]={'count':c,'examples':ex}
        return res

    def cnt_pos(texts, kw_dict, used):
        res={}
        for lbl,kws in kw_dict.items():
            matched=[]
            for tx in texts:
                s=str(tx).lower()
                hit=[k for k in kws if k in s]
                if hit: matched.append((str(tx).strip(), hit))
            c=len(matched)
            ex_best=[]; ex_good=[]; ex_any=[]
            for raw,hit in matched:
                if len(raw)<8 or len(raw)>150 or raw.count('?')>=2 or raw in used: continue
                kv=any(k in raw for k in hit)
                sc=sentiment_score(raw)
                if kv and sc>=2 and len(ex_best)<3: ex_best.append(raw[:90])
                elif kv and sc>=0 and len(ex_good)<3: ex_good.append(raw[:90])
                elif kv and len(ex_any)<3: ex_any.append(raw[:90])
            ex_final=[]
            for pool in [ex_best,ex_good,ex_any]:
                for item in pool:
                    if len(ex_final)>=3: break
                    if item not in ex_final and item not in used:
                        ex_final.append(item); used.add(item)
                if len(ex_final)>=3: break
            res[lbl]={'count':c,'examples':ex_final}
        return res

    # 다국어 감성사전 적용 (#6)
    # session_state에서 region_code 참조
    _region = st.session_state.get('_region_code', 'KR')
    _en_neg, _en_pos = get_kw_dicts_by_region(_region, '')
    if _en_neg:
        neg_kw_use = _en_neg
        pos_kw_use = _en_pos
    else:
        neg_kw_use = t['neg_kw']
        pos_kw_use = t['pos_kw']

    used_neg=set(); used_pos=set()
    nr=cnt_neg(neg_tx, neg_kw_use, used_neg)
    pr=cnt_pos(pos_tx, pos_kw_use, used_pos)
    ns=sorted(nr.items(), key=lambda x:x[1]['count'], reverse=True)
    ps=sorted(pr.items(), key=lambda x:x[1]['count'], reverse=True)

    def lv(c, is_neg):
        if is_neg: return t['lv_very_many'] if c>=30 else (t['lv_many'] if c>=15 else (t['lv_normal'] if c>=5 else t['lv_few']))
        else:      return t['lv_very_many'] if c>=40 else (t['lv_many'] if c>=20 else (t['lv_normal'] if c>=5 else t['lv_few']))

    rows=[]
    for i,(lbl,d) in enumerate(ns,1):
        if d['count']==0: continue
        ex='  /  '.join([f'"{e}"' for e in d['examples']]) or '-'
        rows.append({'rank':f'🔴 {i}{t["rank_suffix"]}','kw':lbl,'sent':'neg','cnt':d['count'],
                     'cnt_pct':f'{lv(d["count"],True)} ({d["count"]:,}{t["unit_reviews"]})',
                     'sum':t['kw_neg_sum'].format(neg_total,d['count'],lbl),'ex':ex})
    for i,(lbl,d) in enumerate(ps,1):
        if d['count']==0: continue
        ex='  /  '.join([f'"{e}"' for e in d['examples']]) or '-'
        rows.append({'rank':f'🟢 {i}{t["rank_suffix"]}','kw':lbl,'sent':'pos','cnt':d['count'],
                     'cnt_pct':f'{lv(d["count"],False)} ({d["count"]:,}{t["unit_reviews"]})',
                     'sum':t['kw_pos_sum'].format(pos_total,d['count'],lbl),'ex':ex})

    total=len(df); avg=df['평점'].mean()
    pos=(df['평점']>=4).sum(); neg=(df['평점']<=2).sum()
    top_neg=ns[0][0] if ns else '-'; top_pos=ps[0][0] if ps else '-'
    pos_r=pos/total; neg_r=neg/total
    if pos_r>0.6:   mood=t['mood_pos'].format(pos_r*100)
    elif pos_r>0.4: mood=t['mood_mix'].format(pos_r*100, neg_r*100)
    else:           mood=t['mood_neg'].format(neg_r*100)
    one=t['one_line_fmt'].format(top_pos, top_neg, avg, mood)
    return rows, one, ns, ps

def mk_dash(ws, df, t):
    ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:P3'); c=ws['A1']; c.value=t['dash_ttl']
    c.font=Font(bold=True,size=20,color=C_WHITE,name='Arial')
    c.fill=_fill(C_DARK); c.alignment=_al()
    for r in [1,2,3]: ws.row_dimensions[r].height=44
    ws.merge_cells('A4:P4'); s=ws['A4']
    s.value=t['dash_sub_fmt'].format(df['작성일'].min(),df['작성일'].max(),len(df),datetime.now().strftime('%Y-%m-%d %H:%M'))
    s.font=Font(size=9,color='BBBBBB',name='Arial'); s.fill=_fill(C_MID); s.alignment=_al('left')
    ws.row_dimensions[4].height=18; ws.row_dimensions[5].height=10
    total=len(df); avg=df['평점'].mean()
    pos=(df['평점']>=4).sum(); neg=(df['평점']<=2).sum()
    neu=(df['평점']==3).sum()
    rep=(df['개발사답변'].notna()&(df['개발사답변']!='')).sum()
    kpis=[(t['kpi_avg'],f'{avg:.2f}',C_GOLD,'★/5.00'),
          (t['kpi_total'],f'{total:,}',C_DARK,''),
          (t['kpi_pos'],f'{pos:,}',C_GREEN,f'{pos/total*100:.1f}%'),
          (t['kpi_neu'],f'{neu:,}',C_BLUE,f'{neu/total*100:.1f}%'),
          (t['kpi_neg'],f'{neg:,}',C_RED,f'{neg/total*100:.1f}%'),
          (t['kpi_rep'],f'{rep:,}',C_ACCENT,f'{rep/total*100:.1f}%')]
    for (c1,c2),(lbl,val,col,sub) in zip(zip([1,3,5,8,11,14],[2,4,7,10,13,16]),kpis):
        for r in [6,7,8,9]: ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
        lc=ws.cell(row=6,column=c1,value=lbl); lc.font=Font(bold=True,size=9,color=C_WHITE,name='Arial')
        lc.fill=_fill(col); lc.alignment=_al()
        vc=ws.cell(row=7,column=c1,value=val); vc.font=Font(bold=True,size=20,color=C_WHITE,name='Arial')
        vc.fill=_fill(col); vc.alignment=_al()
        sc=ws.cell(row=8,column=c1,value=sub); sc.font=Font(size=9,color=C_WHITE,name='Arial')
        sc.fill=_fill(col); sc.alignment=_al()
        ws.cell(row=9,column=c1).fill=_fill(col)
        for r,h in [(6,18),(7,38),(8,16),(9,6)]: ws.row_dimensions[r].height=h
    dist=df['평점'].value_counts().sort_index()
    mo=df.groupby('작성월').agg(cnt=('평점','count'),avg=('평점','mean')).reset_index()

    _hidden(ws,1,19,t['c_score']); _hidden(ws,1,20,t['rev_cnt'])
    for i,star in enumerate([1,2,3,4,5],2):
        _hidden(ws,i,19,f'{star}★'); _hidden(ws,i,20,int(dist.get(star,0)))
    _hidden(ws,8,19,''); _hidden(ws,8,20,t['rev_cnt'])
    for i,(k,v) in enumerate([(t['pie_pos'],int((df['평점']>=4).sum())),(t['pie_neu'],int((df['평점']==3).sum())),(t['pie_neg'],int((df['평점']<=2).sum()))],9):
        _hidden(ws,i,19,k); _hidden(ws,i,20,v)
    _hidden(ws,13,19,''); _hidden(ws,13,20,t['rev_cnt']); _hidden(ws,13,21,t['avg_sc'])
    for i,r in enumerate(mo.itertuples(),14):
        _hidden(ws,i,19,r.작성월); _hidden(ws,i,20,r.cnt); _hidden(ws,i,21,round(r.avg,2))
    _sec(ws,11,1,8,t['ch_dist'])
    bar=BarChart(); bar.type='col'; bar.style=10; bar.title=None; bar.legend=None
    bar.y_axis.title=t['rev_cnt']; bar.width=15; bar.height=12
    bar.add_data(Reference(ws,min_col=20,min_row=1,max_row=6),titles_from_data=True)
    bar.set_categories(Reference(ws,min_col=19,min_row=2,max_row=6))
    for idx,c in enumerate(['C0392B','E67E22','F5A623','2ECC71','27AE60']):
        pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=c; bar.series[0].dPt.append(pt)
    ws.add_chart(bar,'A12')
    _sec(ws,11,9,16,t['ch_pie'])
    pie=PieChart(); pie.style=10; pie.title=None; pie.width=15; pie.height=12
    pie.add_data(Reference(ws,min_col=20,min_row=8,max_row=11),titles_from_data=True)
    pie.set_categories(Reference(ws,min_col=19,min_row=9,max_row=11))
    for idx,c in enumerate(['27AE60','F5A623','C0392B']):
        pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=c; pie.series[0].dPt.append(pt)
    ws.add_chart(pie,'I12')
    ws.row_dimensions[28].height=10; _sec(ws,29,1,16,t['ch_trend'])
    n=len(mo); cats=Reference(ws,min_col=19,min_row=14,max_row=13+n)
    bar2=BarChart(); bar2.type='col'; bar2.style=10; bar2.title=None
    bar2.width=32; bar2.height=13; bar2.y_axis.title=t['rev_cnt']; bar2.y_axis.axId=100
    d_bar=Reference(ws,min_col=20,min_row=13,max_row=13+n)
    bar2.add_data(d_bar,titles_from_data=True); bar2.set_categories(cats)
    bar2.series[0].graphicalProperties.solidFill=C_ACCENT
    line2=LineChart(); line2.style=10; line2.title=None
    line2.y_axis.title=t['avg_sc']; line2.y_axis.axId=200
    line2.y_axis.crosses='max'; line2.y_axis.crossAx=100
    line2.y_axis.scaling.min=1.0; line2.y_axis.scaling.max=5.0
    d_line=Reference(ws,min_col=21,min_row=13,max_row=13+n)
    line2.add_data(d_line,titles_from_data=True); line2.set_categories(cats)
    line2.series[0].graphicalProperties.line.solidFill=C_GOLD
    line2.series[0].graphicalProperties.line.width=28000
    line2.series[0].marker.symbol='circle'; line2.series[0].marker.size=6
    line2.series[0].marker.graphicalProperties.solidFill=C_GOLD
    bar2+=line2; ws.add_chart(bar2,'A30')
    ws.row_dimensions[46].height=10; _sec(ws,47,1,16,t['ch_top10'])
    _hr(ws,48,[t['c_user'],t['c_score'],t['c_like'],t['c_content'],t['hd_date']],bg=C_MID)
    ws.merge_cells('D48:O48')
    _w(ws,48,16,t['hd_date'],bold=True,sz=10,fg=C_WHITE,bg=C_MID,h='center')
    for i,(_,r) in enumerate(df.nlargest(10,'좋아요').iterrows(),49):
        bg=C_LIGHT if i%2==0 else C_WHITE; ws.row_dimensions[i].height=36
        _w(ws,i,1,r['사용자'],sz=9,bg=bg,h='center')
        sc=r['평점']; sc_c=C_GREEN if sc>=4 else (C_ORANGE if sc==3 else C_RED)
        _w(ws,i,2,f'{sc}★',bold=True,sz=11,fg=sc_c,bg=bg,h='center')
        _w(ws,i,3,r['좋아요'],sz=9,bg=bg,h='center')
        ws.merge_cells(start_row=i,start_column=4,end_row=i,end_column=15)
        _w(ws,i,4,r['내용'][:120],sz=9,bg=bg,h='left',wrap=True)
        _w(ws,i,16,r['작성일'],sz=9,bg=bg,h='center')
    _cw(ws,{c:w for c,w in zip('ABCDEFGHIJKLMNOP',[14,7,8,10,8,8,8,8,8,8,8,8,8,8,8,12])})

def mk_opinion(ws, df, t, doc_lang):
    ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:L3'); c=ws['A1']; c.value=t['op_ttl']
    c.font=Font(bold=True,size=20,color=C_WHITE,name='Arial')
    c.fill=_fill(C_DARK); c.alignment=_al()
    for r in [1,2,3]: ws.row_dimensions[r].height=44
    ws.merge_cells('A4:L4'); s=ws['A4']
    total=len(df); pos=(df['평점']>=4).sum(); neg=(df['평점']<=2).sum()
    s.value=t['op_sub_fmt'].format(total,pos/total*100,neg/total*100,datetime.now().strftime('%Y-%m-%d'))
    s.font=Font(size=9,color='BBBBBB',name='Arial'); s.fill=_fill(C_MID); s.alignment=_al('left')
    ws.row_dimensions[4].height=18; ws.row_dimensions[5].height=12
    _sec(ws,6,1,12,t['ins_ttl'],bg=C_MID)
    hrow=7; ws.row_dimensions[hrow].height=20
    _mw(ws,hrow,1,hrow,4,t['ins_cat'],bold=True,sz=10,fg=C_WHITE,bg=C_DARK,h='center')
    _mw(ws,hrow,5,hrow,12,t['ins_eval'],bold=True,sz=10,fg=C_WHITE,bg=C_DARK,h='center')
    categories=evaluate_5_insights(df,doc_lang)
    for i,(cat,ev,color) in enumerate(categories,hrow+1):
        bg=C_LIGHT if i%2==0 else C_WHITE; ws.row_dimensions[i].height=30
        _mw(ws,i,1,i,4,cat,bold=True,sz=10,fg=C_WHITE,bg=color,h='center',v='center')
        _mw(ws,i,5,i,12,ev,bold=True,sz=11,fg=color,bg=bg,h='left',v='center')
    ws.row_dimensions[hrow+len(categories)+1].height=14
    # 스팀이면 스팀 전용 분석 사용
    if '추천여부' in df.columns:
        kw_rows,one_line,ns,ps=analyze_kw_steam(df,t)
    else:
        kw_rows,one_line,ns,ps=analyze_kw(df,t)
    sum_row=hrow+len(categories)+2
    ws.merge_cells(start_row=sum_row,start_column=1,end_row=sum_row,end_column=12)
    c=ws.cell(row=sum_row,column=1,value=f'{t["sum_pfx"]}{one_line}')
    c.font=Font(bold=True,sz=11,color=C_WHITE,name='Arial')
    c.fill=_fill(C_ACCENT); c.alignment=_al('left',wrap=True)
    ws.row_dimensions[sum_row].height=38; ws.row_dimensions[sum_row+1].height=14
    tbl=sum_row+2; _sec(ws,tbl,1,12,t['kw_ttl'])
    h2=tbl+1; ws.row_dimensions[h2].height=22
    for (c1,c2),lbl in [((1,1),t['col_rank']),((2,2),t['col_kw']),((3,3),t['col_sent']),
                         ((4,4),t['col_cnt']),((5,8),t['col_sum']),((9,12),t['col_ex'])]:
        _mw(ws,h2,c1,h2,c2,lbl,bold=True,sz=10,fg=C_WHITE,bg=C_MID,h='center')
    drow=h2+1
    for item in kw_rows:
        is_neg=item['sent']=='neg'; sc=C_RED if is_neg else C_GREEN; sb='FFF0F0' if is_neg else 'F0FFF4'
        ws.row_dimensions[drow].height=52
        _mw(ws,drow,1,drow,1,item['rank'],bold=True,sz=11,fg=C_WHITE,bg=sc,h='center',v='center')
        _mw(ws,drow,2,drow,2,item['kw'],bold=True,sz=10,fg=sc,bg=sb,h='center',v='center',wrap=True)
        _mw(ws,drow,3,drow,3,t['neg_lbl'] if is_neg else t['pos_lbl'],bold=True,sz=9,fg=sc,bg=sb,h='center',v='center')
        _mw(ws,drow,4,drow,4,item['cnt_pct'],sz=9,fg='444444',bg=sb,h='center',v='center',wrap=True)
        _mw(ws,drow,5,drow,8,item['sum'],sz=9,fg='222222',bg=sb,h='left',v='center',wrap=True)
        _mw(ws,drow,9,drow,12,item['ex'],sz=9,fg='555555',bg=sb,h='left',v='center',wrap=True,it=True)
        drow+=1
    ws.row_dimensions[drow].height=14
    _hidden(ws,1,14,t['col_kw']); _hidden(ws,1,15,t['col_cnt'])
    for i,(lbl,d) in enumerate(ns[:6],2):
        _hidden(ws,i,14,lbl); _hidden(ws,i,15,d['count'])
    _hidden(ws,9,14,t['col_kw']); _hidden(ws,9,15,t['col_cnt'])
    for i,(lbl,d) in enumerate(ps[:5],10):
        _hidden(ws,i,14,lbl); _hidden(ws,i,15,d['count'])
    cr=drow+1; nn=min(6,len(ns)); np2=min(5,len(ps))
    _sec(ws,cr,1,6,t['neg_ch'],bg=C_RED)
    nb=BarChart(); nb.type='bar'; nb.style=10; nb.title=None; nb.legend=None
    nb.x_axis.title=t['col_cnt']; nb.width=14; nb.height=11
    nb.add_data(Reference(ws,min_col=15,min_row=1,max_row=1+nn),titles_from_data=True)
    nb.set_categories(Reference(ws,min_col=14,min_row=2,max_row=1+nn))
    for idx in range(nn):
        pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=C_RED; nb.series[0].dPt.append(pt)
    ws.add_chart(nb,f'A{cr+1}')
    _sec(ws,cr,7,12,t['pos_ch'],bg=C_GREEN)
    pb=BarChart(); pb.type='bar'; pb.style=10; pb.title=None; pb.legend=None
    pb.x_axis.title=t['col_cnt']; pb.width=14; pb.height=11
    pb.add_data(Reference(ws,min_col=15,min_row=9,max_row=9+np2),titles_from_data=True)
    pb.set_categories(Reference(ws,min_col=14,min_row=10,max_row=9+np2))
    for idx in range(np2):
        pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=C_GREEN; pb.series[0].dPt.append(pt)
    ws.add_chart(pb,f'G{cr+1}')
    _cw(ws,{'A':10,'B':18,'C':10,'D':10,'E':14,'F':14,'G':14,'H':14,'I':16,'J':16,'K':16,'L':16})

def mk_raw(ws, df, t):
    ws.sheet_view.showGridLines=False; ws.freeze_panes='A2'
    cols=[t['c_id'],t['c_user'],t['c_score'],t['c_content'],t['c_date'],
          t['c_month'],t['c_like'],t['c_reply'],t['c_rdate'],t['c_ver']]
    dk=['리뷰ID','사용자','평점','내용','작성일','작성월','좋아요','개발사답변','답변일','앱버전']
    _hr(ws,1,cols,height=20)
    for i,row in enumerate(dataframe_to_rows(df[dk],index=False,header=False),2):
        sc=row[2]; bg=C_LGREEN if sc>=4 else (C_YELLOW if sc==3 else C_LRED)
        ws.row_dimensions[i].height=14
        for j,val in enumerate(row,1):
            c=ws.cell(row=i,column=j,value=val)
            c.fill=_fill(bg); c.border=_bd(); c.font=Font(size=9,name='Arial')
            c.alignment=_al('left','center',wrap=(j in [4,8]))
            if j==3:
                cc=C_GREEN if sc>=4 else (C_ORANGE if sc==3 else C_RED)
                c.font=Font(bold=True,size=9,color=cc,name='Arial'); c.alignment=_al()
    ws.auto_filter.ref=f'A1:{get_column_letter(len(cols))}1'
    _cw(ws,{'A':18,'B':12,'C':6,'D':50,'E':12,'F':10,'G':8,'H':35,'I':12,'J':10})

def mk_stats(ws, df, t):
    ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:H1'); c=ws['A1']; c.value=t['stat_ttl']
    c.font=Font(bold=True,size=15,color=C_WHITE,name='Arial')
    c.fill=_fill(C_DARK); c.alignment=_al(); ws.row_dimensions[1].height=32
    ws.row_dimensions[2].height=10; _sec(ws,3,1,8,t['stat_sc'])
    _hr(ws,4,[t['c_score'],t['hd_cnt'],t['hd_ratio'],t['hd_alike'],t['hd_mlike'],t['hd_rcnt'],t['hd_rrate'],''],height=20)
    for i,star in enumerate([5,4,3,2,1],5):
        sub=df[df['평점']==star]; cnt=len(sub)
        bg=C_LGREEN if star>=4 else (C_YELLOW if star==3 else C_LRED)
        cc=C_GREEN if star>=4 else (C_ORANGE if star==3 else C_RED)
        rc=(sub['개발사답변'].notna()&(sub['개발사답변']!='')).sum()
        vals=[f'{star}★',cnt,round(cnt/len(df)*100,1),
              round(sub['좋아요'].mean(),1) if cnt else 0,
              int(sub['좋아요'].max()) if cnt else 0,int(rc),
              round(rc/cnt*100,1) if cnt else 0,'●']
        ws.row_dimensions[i].height=22
        for j,v in enumerate(vals,1):
            c=ws.cell(row=i,column=j,value=v); c.fill=_fill(bg); c.border=_bd(); c.alignment=_al()
            c.font=Font(bold=(j in [1,8]),size=11 if j==1 else (14 if j==8 else 10),
                        color=cc if j in [1,8] else '000000',name='Arial')
    ws.row_dimensions[10].height=12; _sec(ws,11,1,8,t['stat_mo'])
    _hr(ws,12,[t['hd_month'],t['hd_cnt'],t['hd_avg'],'5★','4★','3★','2★','1★'],height=20)
    for i,(mo,g) in enumerate(df.groupby('작성월'),13):
        bg=C_LIGHT if i%2==0 else C_WHITE; ws.row_dimensions[i].height=20
        for j,v in enumerate([mo,len(g),round(g['평점'].mean(),2),
                               int((g['평점']==5).sum()),int((g['평점']==4).sum()),
                               int((g['평점']==3).sum()),int((g['평점']==2).sum()),
                               int((g['평점']==1).sum())],1):
            c=ws.cell(row=i,column=j,value=v); c.fill=_fill(bg); c.border=_bd()
            c.font=Font(size=10,name='Arial'); c.alignment=_al()
    _cw(ws,{'A':12,'B':10,'C':12,'D':8,'E':8,'F':8,'G':8,'H':8})

def mk_criteria(ws, t):
    ws.sheet_view.showGridLines=False
    is_kr = (t is I18N['KR'])

    # 타이틀
    title_txt = '📐 분석 기준 시트' if is_kr else '📐 分析基準シート'
    ws.merge_cells('A1:F1')
    c=ws['A1']; c.value=title_txt
    c.font=Font(bold=True,size=15,color=C_WHITE,name='Arial')
    c.fill=_fill(C_DARK); c.alignment=_al(); ws.row_dimensions[1].height=32
    ws.row_dimensions[2].height=10

    # ── 섹션1: 감성 점수 기준
    sec1 = '⭐ 감성 점수 기준' if is_kr else '⭐ 感情スコア基準'
    _sec(ws,3,1,6,sec1)
    headers = ['구분','기준','설명'] if is_kr else ['区分','基準','説明']
    _hr(ws,4,headers+['','',''],bg=C_MID)
    score_rows_kr = [
        ('분류 방식','평점 40% + 텍스트 60%','평점만으로 분류하지 않고 리뷰 내용을 함께 반영'),
        ('긍정 분류','조합 점수 > +0.3','긍정 키워드 多, 부정어 적음 → 긍정 버킷'),
        ('부정 분류','조합 점수 < -0.3','부정 키워드 多, 긍정 표현 없음 → 부정 버킷'),
        ('중립','조합 점수 -0.3~+0.3','긍정·부정 혼재 또는 판단 어려움 → 제외'),
        ('역접어 처리','한데/지만/근데 등','역접어 뒤 내용을 최종 감정으로 판단 (가중치 2배)'),
        ('부정어 조합','없어요/안/못/전혀 등','부정 키워드 앞뒤에 부정어 있으면 긍정으로 재분류'),
        ('예시','평점 4점 + 불만 내용','→ 부정으로 정확히 분류 (기존엔 평점만 보고 긍정 오분류)'),
    ]
    score_rows_jp = [
        ('分類方式','評価40% + テキスト60%','評価だけでなくレビュー内容も反映'),
        ('肯定分類','総合スコア > +0.3','肯定KW多、否定語少 → 肯定バケット'),
        ('否定分類','総合スコア < -0.3','否定KW多、肯定表現なし → 否定バケット'),
        ('中立','スコア -0.3~+0.3','肯定・否定混在または判断困難 → 除外'),
        ('逆接語処理','けど/が/でも等','逆接語以降の内容を最終感情として判断(重み2倍)'),
        ('否定語組合せ','ない/ず/ません等','否定KW前後に否定語があれば肯定に再分類'),
        ('例','評価4点 + 不満内容','→ 否定に正確に分類(旧:評価だけ見て誤分類)'),
    ]
    score_rows = score_rows_kr if is_kr else score_rows_jp
    for i,(a,b,c_txt) in enumerate(score_rows, 5):
        bg = C_LIGHT if i%2==0 else C_WHITE
        ws.row_dimensions[i].height=22
        _w(ws,i,1,a,bold=True,sz=10,bg=bg,h='center')
        _w(ws,i,2,b,sz=10,bg=bg,h='center')
        ws.merge_cells(start_row=i,start_column=3,end_row=i,end_column=6)
        _w(ws,i,3,c_txt,sz=10,bg=bg,h='left',wrap=True)
    ws.row_dimensions[10].height=12

    # ── 섹션2: 부정 키워드
    sec2 = '🔴 부정 키워드 목록' if is_kr else '🔴 否定キーワード一覧'
    _sec(ws,11,1,6,sec2,bg=C_RED)
    h2 = ['카테고리','키워드 목록'] if is_kr else ['カテゴリ','キーワード一覧']
    _hr(ws,12,h2+['','','',''],bg=C_MID)
    row=13
    for lbl,kws in t['neg_kw'].items():
        bg=C_LRED if row%2==0 else C_WHITE
        ws.row_dimensions[row].height=20
        _w(ws,row,1,lbl,bold=True,sz=10,fg=C_RED,bg=bg,h='center')
        ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=6)
        _w(ws,row,2,'  /  '.join(kws),sz=9,bg=bg,h='left')
        row+=1
    ws.row_dimensions[row].height=12; row+=1

    # ── 섹션3: 긍정 키워드
    sec3 = '🟢 긍정 키워드 목록' if is_kr else '🟢 肯定キーワード一覧'
    _sec(ws,row,1,6,sec3,bg=C_GREEN); row+=1
    h3 = ['카테고리','키워드 목록'] if is_kr else ['カテゴリ','キーワード一覧']
    _hr(ws,row,h3+['','','',''],bg=C_MID); row+=1
    for lbl,kws in t['pos_kw'].items():
        bg=C_LGREEN if row%2==0 else C_WHITE
        ws.row_dimensions[row].height=20
        _w(ws,row,1,lbl,bold=True,sz=10,fg=C_GREEN,bg=bg,h='center')
        ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=6)
        _w(ws,row,2,'  /  '.join(kws),sz=9,bg=bg,h='left')
        row+=1
    ws.row_dimensions[row].height=12; row+=1

    # ── 섹션4: 부정어 목록
    sec4 = '🚫 부정어 목록 (이 단어+부정키워드 = 긍정 재분류)' if is_kr else '🚫 否定語一覧 (否定語+否定KW = 肯定に再分類)'
    _sec(ws,row,1,6,sec4,bg=C_BLUE); row+=1
    neg_words_kr = ['없어요','없음','없네요','없다','안','못','안됨','안돼','노','전혀','하나도','거의','별로 안','전혀 안']
    neg_words_jp = ['ない','ず','ません','なし','全然','ほとんど','あまり','全く']
    neg_words = neg_words_kr if is_kr else neg_words_jp
    ws.row_dimensions[row].height=22
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=6)
    _w(ws,row,1,'  /  '.join(neg_words),sz=10,bg=C_LIGHT,h='left',wrap=True)
    row+=1
    ws.row_dimensions[row].height=12  # 여백 행

    _cw(ws,{'A':20,'B':50,'C':15,'D':15,'E':15,'F':15})

def mk_dash_steam(ws, df, t):
    '''스팀 전용 대시보드 시트'''
    ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:P3'); c=ws['A1']
    c.value='🎮  Steam 리뷰 분석 대시보드' if t is I18N.get('KR',t) else '🎮  Steam レビュー分析ダッシュボード'
    c.font=Font(bold=True,size=20,color=C_WHITE,name='Arial')
    c.fill=_fill('1B2838'); c.alignment=_al()
    for r in [1,2,3]: ws.row_dimensions[r].height=44
    ws.merge_cells('A4:P4'); s=ws['A4']
    s.value=t['dash_sub_fmt'].format(df['작성일'].min(),df['작성일'].max(),len(df),datetime.now().strftime('%Y-%m-%d %H:%M'))
    s.font=Font(size=9,color='BBBBBB',name='Arial'); s.fill=_fill(C_MID); s.alignment=_al('left')
    ws.row_dimensions[4].height=18; ws.row_dimensions[5].height=10
    total=len(df)
    rec=(df['평점']==5).sum(); norec=(df['평점']==1).sum()
    avg_pt=df['플레이시간'].mean() if '플레이시간' in df.columns else 0
    rep=(df['개발사답변'].notna()&(df['개발사답변']!='')).sum()
    rec_rate=rec/total*100 if total else 0
    kpis=[
        ('👍 추천률', f'{rec_rate:.1f}%', '1B6A3C', f'{rec:,}건'),
        ('📝 총 리뷰', f'{total:,}', C_DARK, ''),
        ('👍 추천', f'{rec:,}', C_GREEN, f'{rec_rate:.1f}%'),
        ('👎 비추천', f'{norec:,}', C_RED, f'{norec/total*100:.1f}%'),
        ('⏱️ 평균플레이', f'{avg_pt/60:.0f}h', C_BLUE, f'{avg_pt:.0f}분'),
        ('💬 개발사답변', f'{rep:,}', C_ACCENT, f'{rep/total*100:.1f}%'),
    ]
    for (c1,c2),(lbl,val,col,sub) in zip(zip([1,3,5,8,11,14],[2,4,7,10,13,16]),kpis):
        for r in [6,7,8,9]: ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
        lc=ws.cell(row=6,column=c1,value=lbl); lc.font=Font(bold=True,size=9,color=C_WHITE,name='Arial')
        lc.fill=_fill(col); lc.alignment=_al()
        vc=ws.cell(row=7,column=c1,value=val); vc.font=Font(bold=True,size=20,color=C_WHITE,name='Arial')
        vc.fill=_fill(col); vc.alignment=_al()
        sc=ws.cell(row=8,column=c1,value=sub); sc.font=Font(size=9,color=C_WHITE,name='Arial')
        sc.fill=_fill(col); sc.alignment=_al()
        ws.cell(row=9,column=c1).fill=_fill(col)
        for r,h in [(6,18),(7,38),(8,16),(9,6)]: ws.row_dimensions[r].height=h
    # 차트용 숨김 데이터
    mo=df.groupby('작성월').agg(cnt=('평점','count'),rec=('평점',lambda x:(x==5).sum())).reset_index()
    _hidden(ws,1,19,'추천/비추천'); _hidden(ws,1,20,'건수')
    _hidden(ws,2,19,'👍 추천'); _hidden(ws,2,20,int(rec))
    _hidden(ws,3,19,'👎 비추천'); _hidden(ws,3,20,int(norec))
    _hidden(ws,5,19,'월'); _hidden(ws,5,20,'리뷰수'); _hidden(ws,5,21,'추천수')
    for i,r in enumerate(mo.itertuples(),6):
        _hidden(ws,i,19,r.작성월); _hidden(ws,i,20,r.cnt); _hidden(ws,i,21,int(r.rec))
    _sec(ws,11,1,8,'📊 추천/비추천 분포')
    pie=PieChart(); pie.style=10; pie.title=None; pie.width=15; pie.height=12
    pie.add_data(Reference(ws,min_col=20,min_row=1,max_row=3),titles_from_data=True)
    pie.set_categories(Reference(ws,min_col=19,min_row=2,max_row=3))
    for idx,c in enumerate(['27AE60','C0392B']):
        pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=c; pie.series[0].dPt.append(pt)
    ws.add_chart(pie,'A12')
    n=len(mo); _sec(ws,29,1,16,'📈 월별 리뷰 수 & 추천 추이')
    bar2=BarChart(); bar2.type='col'; bar2.style=10; bar2.title=None
    bar2.width=32; bar2.height=13; bar2.y_axis.title='리뷰수'; bar2.y_axis.axId=100
    d_bar=Reference(ws,min_col=20,min_row=5,max_row=5+n)
    bar2.add_data(d_bar,titles_from_data=True)
    bar2.set_categories(Reference(ws,min_col=19,min_row=6,max_row=5+n))
    bar2.series[0].graphicalProperties.solidFill='1B6A3C'
    line2=LineChart(); line2.style=10; line2.title=None
    line2.y_axis.title='추천수'; line2.y_axis.axId=200
    line2.y_axis.crosses='max'; line2.y_axis.crossAx=100
    d_line=Reference(ws,min_col=21,min_row=5,max_row=5+n)
    line2.add_data(d_line,titles_from_data=True)
    line2.set_categories(Reference(ws,min_col=19,min_row=6,max_row=5+n))
    line2.series[0].graphicalProperties.line.solidFill=C_GOLD
    line2.series[0].graphicalProperties.line.width=28000
    bar2+=line2; ws.add_chart(bar2,'A30')
    ws.row_dimensions[46].height=10; _sec(ws,47,1,16,'👍 좋아요 Top 10 리뷰')
    _hr(ws,48,['사용자','추천','플레이시간','내용','날짜'],bg=C_MID)
    ws.merge_cells('D48:O48')
    _w(ws,48,16,'날짜',bold=True,sz=10,fg=C_WHITE,bg=C_MID,h='center')
    for i,(_,r) in enumerate(df.nlargest(10,'좋아요').iterrows(),49):
        rec_flag = r['평점']==5
        bg=C_LGREEN if rec_flag else C_LRED; ws.row_dimensions[i].height=36
        _w(ws,i,1,r['사용자'],sz=9,bg=bg,h='center')
        sc_c=C_GREEN if rec_flag else C_RED
        _w(ws,i,2,'👍' if rec_flag else '👎',bold=True,sz=14,fg=sc_c,bg=bg,h='center')
        pt_h = r.get('플레이시간',0)//60 if '플레이시간' in r.index else 0
        _w(ws,i,3,f'{pt_h}h',sz=9,bg=bg,h='center')
        ws.merge_cells(start_row=i,start_column=4,end_row=i,end_column=15)
        _w(ws,i,4,r['내용'][:120],sz=9,bg=bg,h='left',wrap=True)
        _w(ws,i,16,r['작성일'],sz=9,bg=bg,h='center')
    _cw(ws,{c:w for c,w in zip('ABCDEFGHIJKLMNOP',[14,6,10,10,8,8,8,8,8,8,8,8,8,8,8,12])})

def mk_raw_steam(ws, df, t):
    ws.sheet_view.showGridLines=False; ws.freeze_panes='A2'
    cols=['리뷰ID','사용자','추천여부','내용','작성일','작성월','좋아요','플레이시간(분)','개발사답변','답변일']
    dk=['리뷰ID','사용자','추천여부','내용','작성일','작성월','좋아요','플레이시간','개발사답변','답변일']
    _hr(ws,1,cols,height=20)
    for i,row in enumerate(df[dk].itertuples(index=False),2):
        is_rec = str(row[2]).startswith('👍')
        bg=C_LGREEN if is_rec else C_LRED
        ws.row_dimensions[i].height=14
        for j,val in enumerate(row,1):
            c=ws.cell(row=i,column=j,value=val)
            c.fill=_fill(bg); c.border=_bd(); c.font=Font(size=9,name='Arial')
            c.alignment=_al('left','center',wrap=(j in [4,9]))
            if j==3:
                cc=C_GREEN if is_rec else C_RED
                c.font=Font(bold=True,size=10,color=cc,name='Arial'); c.alignment=_al()
    ws.auto_filter.ref=f'A1:{get_column_letter(len(cols))}1'
    _cw(ws,{'A':18,'B':12,'C':10,'D':50,'E':12,'F':10,'G':8,'H':12,'I':35,'J':12})

def mk_stats_steam(ws, df, t):
    ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:H1'); c=ws['A1']
    c.value='📈 스팀 상세 통계' if t is I18N.get('KR',t) else '📈 Steam詳細統計'
    c.font=Font(bold=True,size=15,color=C_WHITE,name='Arial')
    c.fill=_fill(C_DARK); c.alignment=_al(); ws.row_dimensions[1].height=32
    ws.row_dimensions[2].height=10
    _sec(ws,3,1,8,'👍👎 추천/비추천 통계')
    _hr(ws,4,['구분','리뷰수','비율(%)','평균플레이(h)','평균좋아요','답변수','답변율(%)',''],height=20)
    total=len(df)
    for i,(label,mask) in enumerate([('👍 추천',df['평점']==5),('👎 비추천',df['평점']==1)],5):
        sub=df[mask]; cnt=len(sub)
        bg=C_LGREEN if i==5 else C_LRED; cc=C_GREEN if i==5 else C_RED
        rc=(sub['개발사답변'].notna()&(sub['개발사답변']!='')).sum()
        avg_pt=sub['플레이시간'].mean()/60 if cnt and '플레이시간' in sub.columns else 0
        vals=[label,cnt,round(cnt/total*100,1),round(avg_pt,1),
              round(sub['좋아요'].mean(),1) if cnt else 0,
              int(rc),round(rc/cnt*100,1) if cnt else 0,'●']
        ws.row_dimensions[i].height=22
        for j,v in enumerate(vals,1):
            c=ws.cell(row=i,column=j,value=v); c.fill=_fill(bg); c.border=_bd(); c.alignment=_al()
            c.font=Font(bold=(j in [1,8]),size=10,color=cc if j in [1,8] else '000000',name='Arial')
    ws.row_dimensions[7].height=12; _sec(ws,8,1,8,'📅 월별 추이')
    _hr(ws,9,['월','리뷰수','추천수','비추천수','추천률(%)','','',''],height=20)
    for i,(mo,g) in enumerate(df.groupby('작성월'),10):
        bg=C_LIGHT if i%2==0 else C_WHITE; ws.row_dimensions[i].height=20
        rec=int((g['평점']==5).sum()); norec=int((g['평점']==1).sum())
        for j,v in enumerate([mo,len(g),rec,norec,round(rec/len(g)*100,1) if len(g) else 0],1):
            c=ws.cell(row=i,column=j,value=v); c.fill=_fill(bg); c.border=_bd()
            c.font=Font(size=10,name='Arial'); c.alignment=_al()
    _cw(ws,{'A':12,'B':10,'C':10,'D':10,'E':12,'F':8,'G':8,'H':8})

def mk_criteria_steam(ws, t):
    ws.sheet_view.showGridLines=False
    is_kr = (t is I18N.get('KR', t))
    title_txt = '📐 스팀 분석 기준 시트' if is_kr else '📐 Steam分析基準シート'
    ws.merge_cells('A1:F1'); c=ws['A1']; c.value=title_txt
    c.font=Font(bold=True,size=15,color=C_WHITE,name='Arial')
    c.fill=_fill('1B2838'); c.alignment=_al(); ws.row_dimensions[1].height=32
    ws.row_dimensions[2].height=10
    _sec(ws,3,1,6,'⭐ 스팀 감성 분류 기준' if is_kr else '⭐ Steam感情分類基準')
    _hr(ws,4,['구분','기준','설명','','',''] if is_kr else ['区分','基準','説明','','',''],bg=C_MID)
    rows_kr = [
        ('분류 방식','추천여부 기반','스팀은 추천(👍)/비추천(👎)이 명확 → 그대로 버킷 분류'),
        ('추천(👍)','voted_up = True','긍정 버킷 → 긍정 키워드 분석 대상'),
        ('비추천(👎)','voted_up = False','부정 버킷 → 부정 키워드 분석 대상'),
        ('역접어 처리','한데/지만/했는데 등','역접어 뒤 내용을 최종 감정으로 판단'),
        ('부정어 조합','없어요/안/못/전혀 등','부정KW+부정어 = 긍정으로 재분류'),
        ('플레이시간','author.playtime_forever','Steam API 제공 — 분단위, 60으로 나누면 시간'),
        ('구글플레이 차이','추천/비추천 2단계','구글플레이는 1~5점, 스팀은 추천/비추천만'),
    ]
    rows_jp = [
        ('分類方式','推薦可否ベース','Steamは推薦(👍)/非推薦(👎)が明確 → そのままバケット分類'),
        ('推薦(👍)','voted_up = True','肯定バケット → 肯定KW分析対象'),
        ('非推薦(👎)','voted_up = False','否定バケット → 否定KW分析対象'),
        ('逆接語処理','けど/が/でも等','逆接語以降の内容を最終感情として判断'),
        ('否定語組合せ','ない/ず/ません等','否定KW+否定語 = 肯定に再分類'),
        ('プレイ時間','author.playtime_forever','Steam API提供 — 分単位'),
        ('GP比較','推薦/非推薦2段階','Google Playは1~5点、Steamは推薦/非推薦のみ'),
    ]
    score_rows = rows_kr if is_kr else rows_jp
    for i,(a,b,c_txt) in enumerate(score_rows, 5):
        bg = C_LIGHT if i%2==0 else C_WHITE
        ws.row_dimensions[i].height=22
        _w(ws,i,1,a,bold=True,sz=10,bg=bg,h='center')
        _w(ws,i,2,b,sz=10,bg=bg,h='center')
        ws.merge_cells(start_row=i,start_column=3,end_row=i,end_column=6)
        _w(ws,i,3,c_txt,sz=10,bg=bg,h='left',wrap=True)
    row = 5 + len(score_rows) + 2
    _sec(ws,row,1,6,'🔴 부정 키워드' if is_kr else '🔴 否定キーワード',bg=C_RED); row+=1
    _hr(ws,row,['카테고리','키워드','','','',''] if is_kr else ['カテゴリ','キーワード','','','',''],bg=C_MID); row+=1
    neg_kw_items = [
        ('버그·오류·크래시','버그, 오류, 에러, 크래시, 튕기, 강제종료'),
        ('최적화·성능','렉, 버벅, 최적화, 발열, 프레임, 끊김, 무거'),
        ('핵·치트','핵, 치터, 치트, 어뷰징, 핵유저'),
        ('밸런스·패치','밸런스, 너프, 사기, 패치, 운영, 방치'),
        ('환불·과금','환불, 과금, 비싸, 사기, 돈'),
        ('스토리·콘텐츠','스토리, 콘텐츠, 반복, 지루, 노잼, 부족'),
    ] if is_kr else [
        ('バグ・エラー','バグ, エラー, クラッシュ, 落ちる'),
        ('最適化','重い, カクカク, 最適化, 発熱'),
        ('チート','チート, ハック, 不正'),
        ('バランス','バランス, ナーフ, 運営, 放置'),
        ('返金','返金, 課金, 高い'),
        ('ストーリー','ストーリー, コンテンツ, つまらない'),
    ]
    for a,b in neg_kw_items:
        bg=C_LRED if row%2==0 else C_WHITE; ws.row_dimensions[row].height=20
        _w(ws,row,1,a,bold=True,sz=10,fg=C_RED,bg=bg,h='center')
        ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=6)
        _w(ws,row,2,b,sz=9,bg=bg,h='left'); row+=1
    row+=1
    _sec(ws,row,1,6,'🟢 긍정 키워드' if is_kr else '🟢 肯定キーワード',bg=C_GREEN); row+=1
    _hr(ws,row,['카테고리','키워드','','','',''] if is_kr else ['カテゴリ','キーワード','','','',''],bg=C_MID); row+=1
    pos_kw_items = [
        ('그래픽·비주얼','그래픽, 비주얼, 아트, 예쁘, 퀄리티'),
        ('게임성·전투','전투, 전략, 재밌, 꿀잼, 갓겜, 중독'),
        ('스토리·세계관','스토리, 세계관, 몰입, 감동, 흥미'),
        ('멀티·커뮤니티','친구, 멀티, 파티, 같이, 함께'),
        ('가성비·업데이트','무료, 업데이트, 컨텐츠, 보상, 이벤트'),
    ] if is_kr else [
        ('グラフィック','グラフィック, アート, 綺麗, クオリティ'),
        ('ゲーム性','戦闘, 戦略, 面白い, 神ゲー, 爽快'),
        ('ストーリー','ストーリー, 世界観, 没入, 感動'),
        ('マルチ','友達, マルチ, パーティ, 一緒'),
        ('コスパ','無料, アップデート, 報酬, イベント'),
    ]
    for a,b in pos_kw_items:
        bg=C_LGREEN if row%2==0 else C_WHITE; ws.row_dimensions[row].height=20
        _w(ws,row,1,a,bold=True,sz=10,fg=C_GREEN,bg=bg,h='center')
        ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=6)
        _w(ws,row,2,b,sz=9,bg=bg,h='left'); row+=1
    _cw(ws,{'A':22,'B':45,'C':12,'D':12,'E':12,'F':12})

def gen_excel_bytes(df, doc_lang, is_steam=False):
    t=I18N[doc_lang]; wb=Workbook()
    if is_steam:
        ws1=wb.active; ws1.title='📊 대시보드' if doc_lang=='KR' else '📊 ダッシュボード'
        mk_dash_steam(ws1,df,t)
        ws2=wb.create_sheet(t['sh_op']); mk_opinion(ws2,df,t,doc_lang)
        ws3=wb.create_sheet('📋 전체 리뷰' if doc_lang=='KR' else '📋 全レビュー')
        mk_raw_steam(ws3,df,t)
        ws4=wb.create_sheet(t['sh_stat']); mk_stats_steam(ws4,df,t)
        sh_crit = '📐 분석기준(Steam)' if doc_lang=='KR' else '📐 分析基準(Steam)'
        ws5=wb.create_sheet(sh_crit); mk_criteria_steam(ws5,t)
    else:
        ws1=wb.active; ws1.title=t['sh_dash']; mk_dash(ws1,df,t)
        ws2=wb.create_sheet(t['sh_op']);   mk_opinion(ws2,df,t,doc_lang)
        ws3=wb.create_sheet(t['sh_raw']);  mk_raw(ws3,df,t)
        ws4=wb.create_sheet(t['sh_stat']); mk_stats(ws4,df,t)
        sh_crit = '📐 분석기준' if doc_lang=='KR' else '📐 分析基準'
        ws5=wb.create_sheet(sh_crit); mk_criteria(ws5,t)
    ws1.sheet_properties.tabColor='1B2838' if is_steam else C_ACCENT
    ws2.sheet_properties.tabColor=C_GOLD
    ws3.sheet_properties.tabColor=C_GREEN
    ws4.sheet_properties.tabColor=C_BLUE
    ws5.sheet_properties.tabColor='9B59B6'
    # 워드클라우드 시트 추가
    if HAS_WORDCLOUD:
        sh_wc = '☁️ 워드클라우드' if doc_lang=='KR' else '☁️ ワードクラウド'
        ws6 = wb.create_sheet(sh_wc)
        ws6.sheet_view.showGridLines = False
        ws6.merge_cells('A1:N1'); c=ws6['A1']
        c.value = '☁️  키워드 워드클라우드' if doc_lang=='KR' else '☁️  キーワードワードクラウド'
        c.font=Font(bold=True,size=15,color=C_WHITE,name='Arial')
        c.fill=_fill(C_DARK); c.alignment=_al(); ws6.row_dimensions[1].height=32
        ws6.row_dimensions[2].height=10
        # 부정 워드클라우드
        _sec(ws6,3,1,7,'🔴 부정 리뷰 키워드' if doc_lang=='KR' else '🔴 否定レビューキーワード',bg=C_RED)
        neg_wc = gen_wordcloud_image(df, is_neg=True)
        if neg_wc:
            insert_wordcloud_to_excel(ws6, neg_wc, 'A4')
        # 긍정 워드클라우드
        _sec(ws6,3,8,14,'🟢 긍정 리뷰 키워드' if doc_lang=='KR' else '🟢 肯定レビューキーワード',bg=C_GREEN)
        pos_wc = gen_wordcloud_image(df, is_neg=False)
        if pos_wc:
            insert_wordcloud_to_excel(ws6, pos_wc, 'H4')
        ws6.sheet_properties.tabColor='3498DB'
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

# ══════════════════════════════════════════
# 이상 리뷰 감지 및 필터링 (#8)
# ══════════════════════════════════════════
def filter_abnormal_reviews(df):
    '''의미없는 단어로만 구성된 리뷰 필터링'''
    def is_abnormal(text):
        if pd.isna(text) or str(text).strip() == '':
            return True
        t = str(text).strip()
        # 5글자 미만 단순 감탄사
        if len(t) <= 3:
            return True
        # 동일 문자 반복 (ㅋㅋㅋㅋ, ㅎㅎㅎㅎ, ....., !!!! 등)
        unique_chars = set(t.replace(' ',''))
        if len(unique_chars) <= 2 and len(t) >= 4:
            return True
        # 의미없는 단순 패턴
        meaningless = ['ㅇㅇ','ㄱㄱ','ㄴㄴ','ㅠㅠ','ㅜㅜ','ㅋㅋ','ㅎㅎ',
                       '굿','굳','good','gud','ok','okay','👍','👎',
                       '..','...','....','!!','!!!']
        if t.lower() in [m.lower() for m in meaningless]:
            return True
        return False

    before = len(df)
    df_filtered = df[~df['내용'].apply(is_abnormal)].copy()
    after = len(df_filtered)
    removed = before - after
    return df_filtered, removed

# ══════════════════════════════════════════
# 워드클라우드 생성 (#3)
# ══════════════════════════════════════════
def get_korean_font():
    '''한국어 폰트 경로 반환 — 없으면 나눔고딕 다운로드'''
    import os, urllib.request
    FONT_PATH = '/tmp/NanumGothic.ttf'
    FONT_URL  = 'https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf'

    # 이미 있으면 바로 반환
    if os.path.exists(FONT_PATH):
        return FONT_PATH

    # 시스템 폰트 먼저 확인
    candidates = [
        '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
        '/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf',
        '/usr/share/fonts/truetype/unfonts-core/UnBatang.ttf',
    ]
    for fp in candidates:
        if os.path.exists(fp):
            return fp

    # 없으면 다운로드
    try:
        urllib.request.urlretrieve(FONT_URL, FONT_PATH)
        if os.path.exists(FONT_PATH):
            return FONT_PATH
    except Exception:
        pass
    return None

def gen_wordcloud_image(df, region_code='KR', is_neg=True):
    '''워드클라우드 이미지 생성 → bytes 반환'''
    if not HAS_WORDCLOUD:
        return None
    try:
        import os
        # 감성 기준으로 텍스트 분리
        if is_neg:
            texts = df[df['평점'] <= 2]['내용'].dropna().tolist()
        else:
            texts = df[df['평점'] >= 4]['내용'].dropna().tolist()

        if not texts:
            return None

        text_all = ' '.join(texts)

        # 불용어
        stopwords = set([
            '게임','이게','그게','이거','저거','그거','있어','없어','해요',
            '해서','하고','하는','하면','되는','되어','이라','이런','그런',
            '저런','같아','같은','같이','때문','정말','진짜','너무','매우',
            '조금','좀더','더욱','아주','이제','그냥','그래','그리고',
            'the','and','for','this','that','with','have','from',
        ])

        # 한국어 폰트 가져오기
        font_path = get_korean_font()

        wc_kwargs = dict(
            width=800, height=400,
            background_color='white',
            max_words=80,
            stopwords=stopwords,
            colormap='Reds' if is_neg else 'Greens',
            prefer_horizontal=0.7,
        )
        if font_path:
            wc_kwargs['font_path'] = font_path

        wc = WordCloud(**wc_kwargs).generate(text_all)

        buf = io.BytesIO()
        plt.figure(figsize=(10, 5))
        plt.imshow(wc, interpolation='bilinear')
        plt.axis('off')
        plt.tight_layout(pad=0)
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

def insert_wordcloud_to_excel(ws, img_bytes, anchor='A1'):
    '''엑셀 시트에 워드클라우드 이미지 삽입'''
    if not img_bytes:
        return
    try:
        from openpyxl.drawing.image import Image as XLImage
        buf = io.BytesIO(img_bytes)
        img = XLImage(buf)
        img.width = 600; img.height = 300
        ws.add_image(img, anchor)
    except Exception:
        pass

# ══════════════════════════════════════════
# 다국어 감성사전 확장 (#6)
# ══════════════════════════════════════════
EN_NEG_KW = {
    'Bug·Error'    : ['bug','error','crash','freeze','glitch','broken','lag'],
    'Performance'  : ['laggy','slow','fps','optimization','heating','stutter'],
    'Gacha·Pay'    : ['gacha','p2w','pay to win','expensive','overpriced','whale'],
    'Balance'      : ['imbalanced','nerf','op','broken','unfair'],
    'Operations'   : ['update','devs','support','abandoned','dead game'],
    'Story·Content': ['boring','repetitive','short','lack of content'],
}
EN_POS_KW = {
    'Graphics·Art' : ['beautiful','gorgeous','art','visual','stunning'],
    'Gameplay'     : ['fun','addictive','smooth','satisfying','engaging'],
    'Story'        : ['story','lore','immersive','emotional','deep'],
    'F2P·Generous' : ['free','generous','f2p friendly','no pay wall'],
    'Community'    : ['community','multiplayer','coop','friends'],
}

TW_NEG_KW = {
    '錯誤·當機'    : ['bug','錯誤','當機','閃退','卡頓','lag'],
    '最佳化·效能'  : ['lag','卡','優化','發熱','慢'],
    '抽卡·課金'    : ['課金','抽卡','機率','天井','貴','氪金'],
    '平衡'         : ['不平衡','削弱','外掛','作弊'],
    '營運'         : ['更新','營運','廢棄','死遊戲'],
    '故事·內容'    : ['無聊','重複','內容不足'],
}
TW_POS_KW = {
    '畫面·美術'    : ['美','畫質','美術','精緻'],
    '遊戲性'       : ['好玩','有趣','上癮','爽'],
    '故事'         : ['劇情','故事','感動','沉浸'],
    '課金友善'     : ['不課金','免費','佛心','慷慨'],
    '社群'         : ['多人','朋友','公會','一起玩'],
}

def get_kw_dicts_by_region(region_code, lang_code):
    '''수집 국가에 맞는 키워드 사전 반환'''
    if region_code == 'US' or lang_code == 'en':
        return EN_NEG_KW, EN_POS_KW
    elif region_code == 'TW' or lang_code == 'zh_TW':
        return TW_NEG_KW, TW_POS_KW
    return None, None  # KR/JP는 기존 I18N 사전 사용

def gen_csv_bytes(df, doc_lang='KR'):
    t = I18N[doc_lang]
    dk=['리뷰ID','사용자','평점','내용','작성일','작성월','좋아요','개발사답변','답변일','앱버전']
    col_names=[t['c_id'],t['c_user'],t['c_score'],t['c_content'],t['c_date'],
               t['c_month'],t['c_like'],t['c_reply'],t['c_rdate'],t['c_ver']]
    out_df = df[dk].copy()
    out_df.columns = col_names
    buf=io.StringIO()
    out_df.to_csv(buf, index=False)
    return buf.getvalue().encode('utf-8-sig')  # BOM 포함 UTF-8 (엑셀 호환)

def parse_app_id(text):
    '''구글 플레이 패키지명 추출'''
    text=text.strip()
    m=re.search(r'id=([a-zA-Z0-9._]+)',text)
    if m: return m.group(1)
    if re.match(r'^[a-zA-Z][a-zA-Z0-9._]+$',text): return text
    return None

# ══════════════════════════════════════════
# 스팀 수집 함수
# ══════════════════════════════════════════
def parse_steam_id(text):
    text = text.strip()
    import re as _re
    m = _re.search(r'store[.]steampowered[.]com/app/(\d+)', text)
    if m: return m.group(1)
    if _re.match(r'^\d+$', text): return text
    return None

def fetch_steam_reviews(app_id, language='koreana', how_many=1000, mode='count',
                        dt_from=None, dt_to=None):
    import urllib.request, json
    all_reviews = []; cursor = '*'; batch = 0
    while True:
        url = (f'https://store.steampowered.com/appreviews/{app_id}'
               f'?json=1&language={language}&review_type=all'
               f'&purchase_type=all&num_per_page=100&cursor={urllib.parse.quote(cursor)}'
               f'&filter=recent')
        try:
            import urllib.parse
            url = (f'https://store.steampowered.com/appreviews/{app_id}'
                   f'?json=1&language={language}&review_type=all'
                   f'&purchase_type=all&num_per_page=100'
                   f'&cursor={urllib.parse.quote(str(cursor))}&filter=recent')
            req = urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as r:
                data = json.loads(r.read())
        except Exception as e:
            break
        if data.get('success') != 1: break
        reviews = data.get('reviews', [])
        if not reviews: break
        for rv in reviews:
            ts = rv.get('timestamp_created', 0)
            at = datetime.fromtimestamp(ts) if ts else None
            if mode == 'period' and at:
                if at < dt_from: return all_reviews
                if at > dt_to: continue
            all_reviews.append({
                'reviewId'    : rv.get('recommendationid',''),
                'userName'    : rv.get('author',{}).get('steamid',''),
                'recommended' : rv.get('voted_up', False),
                'score'       : 1 if rv.get('voted_up') else 0,
                'content'     : rv.get('review','').replace('\n',' '),
                'at'          : at,
                'playtime'    : rv.get('author',{}).get('playtime_forever', 0),
                'thumbsUpCount': rv.get('votes_up', 0),
                'replyContent': rv.get('developer_response','') or '',
                'repliedAt'   : None,
                'reviewCreatedVersion': '',
            })
        batch += 1
        new_cursor = data.get('cursor','')
        if not new_cursor or new_cursor == cursor: break
        cursor = new_cursor
        if mode == 'count' and len(all_reviews) >= how_many: break
    return all_reviews[:how_many] if mode == 'count' else all_reviews

def build_df_steam(raw):
    rows = []
    for r in raw:
        at = r.get('at')
        rows.append({
            '리뷰ID'    : str(r.get('reviewId','')),
            '사용자'    : r.get('userName',''),
            '평점'      : 5 if r.get('recommended') else 1,
            '추천여부'  : '👍 추천' if r.get('recommended') else '👎 비추천',
            '내용'      : r.get('content','').replace('\n',' '),
            '작성일'    : at.strftime('%Y-%m-%d') if at else '',
            '작성월'    : at.strftime('%Y-%m') if at else '',
            '좋아요'    : r.get('thumbsUpCount', 0),
            '플레이시간': r.get('playtime', 0),
            '개발사답변': r.get('replyContent',''),
            '답변일'    : '',
            '앱버전'    : '',
        })
    return pd.DataFrame(rows).sort_values('작성일', ascending=False).reset_index(drop=True)

def analyze_kw_steam(df, t):
    '''스팀 전용 여론 분석 — 추천여부 기반 버킷 분류'''
    # 여론분석 시 이상 리뷰 제외 (전체리뷰에는 영향 없음)
    df_analysis, removed = filter_abnormal_reviews(df)
    if removed > 0:
        df = df_analysis
    # 스팀은 추천/비추천이 명확해서 그걸 기준으로 버킷 분류
    neg_tx = df[df['평점']==1]['내용'].dropna()
    pos_tx = df[df['평점']==5]['내용'].dropna()
    neg_total = len(neg_tx); pos_total = len(pos_tx)

    NEG_EXPR = ['버그','최악','짜증','불편','렉','망겜','서운','아쉽','실망','후회',
                '제발','비싸','확률','과금','런함','접음','쓰레기','사기','기만',
                '방치','운영','환불','도박','인플레','믿어본다','한번만','낙담',
                '최적화','발열','튕기','오류','에러','끊김','크래시','프레임',
                '핵','치터','치트','어뷰징','밸런스','너프','핵유저']
    POS_EXPR = ['재밌','좋아','최고','갓겜','꿀잼','대박','추천','만족','감동',
                '몰입','좋음','좋다','재미있','퀄리티','굿','굳','존잼','강추',
                '중독','힐링','명작','인생겜','레전드','완벽','훌륭']
    REVERSAL_KW = ['한데','지만','는데','근데','그러나','하지만','그런데',
                   '이지만','이긴','긴하','했는데','했지만']
    NEGATION_WORDS = ['없어요','없음','없다','없어','안 ','안됨','안돼',
                      '못 ','전혀','하나도','거의']

    def has_negation_near(text, keyword, window=8):
        idx = text.find(keyword)
        if idx == -1: return False
        surrounding = text[max(0,idx-window):idx+len(keyword)+window]
        return any(nw in surrounding for nw in NEGATION_WORDS)

    def sentiment_score(text):
        t_lower = text.lower(); score = 0
        reversal_pos = -1
        for rw in REVERSAL_KW:
            idx = t_lower.find(rw)
            if idx != -1: reversal_pos = idx; break
        if reversal_pos > 0:
            before = t_lower[:reversal_pos]; after = t_lower[reversal_pos:]
            for p in POS_EXPR:
                if p in before: score += 1
            for n in NEG_EXPR:
                if n in before:
                    score += 1 if has_negation_near(before,n) else -1
            for p in POS_EXPR:
                if p in after: score += 2
            for n in NEG_EXPR:
                if n in after:
                    score += 2 if has_negation_near(after,n) else -2
        else:
            for p in POS_EXPR:
                if p in t_lower: score += 2
            for n in NEG_EXPR:
                if n in t_lower:
                    score += 2 if has_negation_near(t_lower,n) else -2
        return score

    def cnt_kw(texts, kw_dict, used, is_neg):
        res = {}
        for lbl, kws in kw_dict.items():
            c = 0; ex = []
            for tx in texts:
                s = str(tx).lower()
                if any(k in s for k in kws):
                    c += 1
                    raw = str(tx).strip()
                    sc = sentiment_score(raw)
                    if len(ex) < 3 and len(raw) >= 10 and raw not in used:
                        if is_neg and sc <= 0:
                            ex.append(raw[:85]); used.add(raw)
                        elif not is_neg and sc >= 0:
                            ex.append(raw[:85]); used.add(raw)
            res[lbl] = {'count': c, 'examples': ex}
        return res

    steam_neg_kw = {
        '버그·오류·크래시': ['버그','오류','에러','크래시','튕기','강제종료'],
        '최적화·성능':      ['렉','버벅','최적화','발열','프레임','끊김','무거'],
        '핵·치트':          ['핵','치터','치트','어뷰징','핵유저'],
        '밸런스·패치':      ['밸런스','너프','사기','패치','운영','방치'],
        '환불·과금':        ['환불','과금','비싸','사기','돈'],
        '스토리·콘텐츠':    ['스토리','콘텐츠','반복','지루','노잼','부족'],
    } if t is I18N.get('KR', t) else {
        'バグ・エラー':     ['バグ','エラー','クラッシュ','落ちる'],
        '最適化・性能':     ['重い','カクカク','最適化','発熱','フレーム'],
        'チート':           ['チート','ハック','不正'],
        'バランス・パッチ': ['バランス','ナーフ','運営','放置'],
        '返金・課金':       ['返金','課金','高い'],
        'ストーリー':       ['ストーリー','コンテンツ','つまらない'],
    }
    steam_pos_kw = {
        '그래픽·비주얼':    ['그래픽','비주얼','아트','예쁘','퀄리티'],
        '게임성·전투':      ['전투','전략','재밌','꿀잼','갓겜','중독'],
        '스토리·세계관':    ['스토리','세계관','몰입','감동','흥미'],
        '멀티·커뮤니티':    ['친구','멀티','파티','같이','함께'],
        '가성비·업데이트':  ['무료','업데이트','컨텐츠','보상','이벤트'],
    } if t is I18N.get('KR', t) else {
        'グラフィック':     ['グラフィック','アート','綺麗','クオリティ'],
        'ゲーム性':         ['戦闘','戦略','面白い','神ゲー','爽快'],
        'ストーリー':       ['ストーリー','世界観','没入','感動'],
        'マルチ':           ['友達','マルチ','パーティ','一緒'],
        'コスパ':           ['無料','アップデート','報酬','イベント'],
    }

    used_neg = set(); used_pos = set()
    nr = cnt_kw(neg_tx, steam_neg_kw, used_neg, True)
    pr = cnt_kw(pos_tx, steam_pos_kw, used_pos, False)
    ns = sorted(nr.items(), key=lambda x:x[1]['count'], reverse=True)
    ps = sorted(pr.items(), key=lambda x:x[1]['count'], reverse=True)

    def lv(c, is_neg):
        if is_neg: return t['lv_very_many'] if c>=30 else (t['lv_many'] if c>=15 else (t['lv_normal'] if c>=5 else t['lv_few']))
        else:      return t['lv_very_many'] if c>=40 else (t['lv_many'] if c>=20 else (t['lv_normal'] if c>=5 else t['lv_few']))

    rows = []
    for i,(lbl,d) in enumerate(ns,1):
        if d['count']==0: continue
        ex='  /  '.join([f'"{e}"' for e in d['examples']]) or '-'
        rows.append({'rank':f'🔴 {i}{t["rank_suffix"]}','kw':lbl,'sent':'neg','cnt':d['count'],
                     'cnt_pct':f'{lv(d["count"],True)} ({d["count"]:,}{t["unit_reviews"]})',
                     'sum':t['kw_neg_sum'].format(neg_total,d['count'],lbl),'ex':ex})
    for i,(lbl,d) in enumerate(ps,1):
        if d['count']==0: continue
        ex='  /  '.join([f'"{e}"' for e in d['examples']]) or '-'
        rows.append({'rank':f'🟢 {i}{t["rank_suffix"]}','kw':lbl,'sent':'pos','cnt':d['count'],
                     'cnt_pct':f'{lv(d["count"],False)} ({d["count"]:,}{t["unit_reviews"]})',
                     'sum':t['kw_pos_sum'].format(pos_total,d['count'],lbl),'ex':ex})

    total=len(df); rec=(df['평점']==5).sum(); norec=(df['평점']==1).sum()
    top_neg=ns[0][0] if ns else '-'; top_pos=ps[0][0] if ps else '-'
    pos_r=rec/total if total else 0; neg_r=norec/total if total else 0
    if pos_r>0.6:   mood=t['mood_pos'].format(pos_r*100)
    elif pos_r>0.4: mood=t['mood_mix'].format(pos_r*100, neg_r*100)
    else:           mood=t['mood_neg'].format(neg_r*100)
    one=t['one_line_fmt'].format(top_pos, top_neg, pos_r*5, mood)
    return rows, one, ns, ps

def parse_appstore_id(text):
    '''앱스토어 앱 ID 추출 (숫자)'''
    text=text.strip()
    # URL 패턴: /id123456789
    m=re.search(r'/id(\d+)', text)
    if m: return m.group(1)
    # 순수 숫자
    if re.match(r'^\d+$', text): return text
    return None

def fetch_appstore_reviews(app_id, country, how_many=500):
    import urllib.request, math
    import xml.etree.ElementTree as ET
    all_reviews = []
    max_page = min(math.ceil(how_many / 50), 10)
    NS = {
        'atom': 'http://www.w3.org/2005/Atom',
        'im':   'http://itunes.apple.com/rss',
    }
    for page in range(1, max_page + 1):
        url = (f'https://itunes.apple.com/{country}/rss/customerreviews/'
               f'page={page}/id={app_id}/sortby=mostrecent/xml')
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as r:
                raw = r.read()
            root = ET.fromstring(raw)
            entries = root.findall('atom:entry', NS)
            if not entries:
                break
            found = 0
            for e in entries:
                rating_el = e.find('im:rating', NS)
                if rating_el is None:
                    continue
                author  = e.findtext('atom:author/atom:name', '', NS)
                rating  = int(rating_el.text or 0)
                content_el = e.find('atom:content', NS)
                content = content_el.text if content_el is not None else ''
                title_el = e.find('atom:title', NS)
                title   = title_el.text if title_el is not None else ''
                updated = e.findtext('atom:updated', '', NS)[:10]
                version = e.findtext('im:version', '', NS)
                rev_id  = e.findtext('atom:id', '', NS)
                all_reviews.append({
                    'id': rev_id, 'userName': author,
                    'rating': rating,
                    'review': f'{title} {content}'.strip(),
                    'date': updated, 'version': version,
                })
                found += 1
            if found == 0:
                break
            if len(all_reviews) >= how_many:
                break
        except Exception:
            break
    return all_reviews[:how_many]

def build_df_appstore(raw):
    '''앱스토어 리뷰 → DataFrame (구글 플레이와 동일 컬럼 구조)'''
    rows=[]
    for r in raw:
        at = r.get('date')
        if isinstance(at, str):
            try: at = datetime.strptime(at[:10], '%Y-%m-%d')
            except: at = None
        rows.append({
            '리뷰ID'    : str(r.get('id','')),
            '사용자'    : r.get('userName','') or r.get('name',''),
            '평점'      : int(r.get('rating', r.get('score', 0))),
            '내용'      : (r.get('review','') or r.get('content','')).replace('\n',' '),
            '작성일'    : at.strftime('%Y-%m-%d') if at else '',
            '작성월'    : at.strftime('%Y-%m') if at else '',
            '좋아요'    : 0,          # 앱스토어는 좋아요 미제공
            '개발사답변': '',          # 앱스토어는 개발사 답변 미제공
            '답변일'    : '',
            '앱버전'    : r.get('version',''),
        })
    return pd.DataFrame(rows).sort_values('작성일',ascending=False).reset_index(drop=True)

def build_df(raw):
    rows=[]
    for r in raw:
        at=r.get('at'); ra=r.get('repliedAt')
        rows.append({'리뷰ID':r.get('reviewId',''),'사용자':r.get('userName',''),
                     '평점':r.get('score',0),'내용':(r.get('content','') or '').replace('\n',' '),
                     '작성일':at.strftime('%Y-%m-%d') if at else '',
                     '작성월':at.strftime('%Y-%m') if at else '',
                     '좋아요':r.get('thumbsUpCount',0),
                     '개발사답변':(r.get('replyContent','') or '').replace('\n',' '),
                     '답변일':ra.strftime('%Y-%m-%d') if ra else '',
                     '앱버전':r.get('reviewCreatedVersion','')})
    return pd.DataFrame(rows).sort_values('작성일',ascending=False).reset_index(drop=True)

# ══════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════
st.set_page_config(page_title='리뷰 분석기 | VIC GAME STUDIOS', page_icon='🎮', layout='wide')
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700&display=swap');
    html,body,[class*="css"]{font-family:'Noto Sans KR',sans-serif;}
    .stApp{background-color:#1A1A2E;color:white;}
    .block-container{padding-top:2rem;max-width:900px;}
    div[data-testid="stMetricValue"]{font-size:28px;font-weight:bold;}
    .stButton>button{background:linear-gradient(135deg,#7B2FBE,#A855F7);color:white;border:none;
        border-radius:8px;padding:12px 32px;font-size:16px;font-weight:bold;width:100%;}
    .stButton>button:hover{background:linear-gradient(135deg,#A855F7,#7B2FBE);}
    .stDownloadButton>button{background:linear-gradient(135deg,#27AE60,#2ECC71);color:white;
        border:none;border-radius:8px;padding:12px 32px;font-size:16px;font-weight:bold;width:100%;}
    .log-box{background:#16213E;border:1px solid #2A2A4A;border-radius:8px;padding:16px;
        font-family:monospace;font-size:12px;color:#CCC;height:280px;overflow-y:auto;white-space:pre-wrap;}
</style>
""", unsafe_allow_html=True)

# 사이드바 - UI 언어 먼저 선택 후 전체 적용
with st.sidebar:
    ui_lang_sel = st.selectbox('UI Language / UI 언어', ['KR (한국어)', 'JP (日本語)'], index=0)
    ui_code = 'KR' if ui_lang_sel.startswith('KR') else 'JP'
    t = I18N[ui_code]
    doc_lang_sel = st.selectbox(t['doc_lang_lbl'], ['KR (한국어)', 'JP (日本語)'], index=0)
    doc_code = 'KR' if doc_lang_sel.startswith('KR') else 'JP'
    st.markdown('---')
    st.markdown(f'**{t["made_by"]}**')
    st.markdown('VIC GAME STUDIOS')
    st.markdown('일본사업실 박경원')
    st.markdown('---')
    PATCH_NOTES = {
        'KR': '''
**v3.0** *(현재 버전)*
- ☁️ 워드클라우드 시각화
- 🌏 영어/중국어 감성사전
- 🧹 이상 리뷰 필터
- 🍎 워드클라우드 폰트 수정

**v2.9**
- ☁️ 워드클라우드 시각화 추가
- 🌏 영어/중국어(번체) 감성사전 추가
- 🧹 이상 리뷰 자동 필터링 추가

**v2.8**
- 🎮 스팀(Steam) 리뷰 수집 지원 추가
- 📐 스팀 전용 분석 기준 시트 추가

**v2.7**
- 🎯 평점 40% + 텍스트 60% 조합 분류 도입
- 📚 부정 키워드 사전 보완 (서운/제발/이격 등)

**v2.6**
- 🚫 부정어+키워드 조합 감성분석 강화
  - "버그 없어요" → 긍정 재분류
  - "렉 안걸려요" → 긍정 재분류
- 📋 분석 기준 시트 추가
- 🔒 여론분석 시트 숨김 데이터 처리
- 📁 파일명 형식 변경 (게임명_국가_마켓_날짜)

**v2.5**
- 🔒 다운로드 후 결과창 유지 개선
- 📊 대시보드 숨김 데이터 노출 수정
- 🍎 앱스토어 500건 제한 안내 문구 추가

**v2.4**
- 🍎 앱스토어(iOS) 리뷰 수집 지원 추가
- 📄 CSV 문서 언어(KR/JP) 반영

**v2.3**
- 🔍 규칙 기반 감성분석 고도화
- 📄 리뷰 전체 CSV 다운로드 기능 추가
- 📋 패치 노트 UI 추가
        ''',
        'JP': '''
**v3.0** *(現在バージョン)*
- ☁️ ワードクラウド可視化
- 🌏 多言語感情辞書
- 🧹 異常レビューフィルター

**v2.9**
- ☁️ ワードクラウド可視化追加
- 🌏 英語/中国語(繁体)感情辞書追加
- 🧹 異常レビュー自動フィルタリング追加

**v2.8**
- 🎮 Steam レビュー収集対応
- 📐 Steam専用分析基準シート追加

**v2.7**
- 🎯 評価40% + テキスト60% 組み合わせ分類導入
- 📚 否定キーワード辞書補強

**v2.6**
- 🚫 否定語+キーワード組み合わせ強化
  - 「バグないです」→ 肯定に再分類
  - 「ラグかからない」→ 肯定に再分類
- 📋 分析基準シート追加
- 🔒 世論分析シート非表示データ処理
- 📁 ファイル名形式変更 (ゲーム名_国_マーケット_日付)

**v2.5**
- 🔒 ダウンロード後も結果画面を維持
- 📊 ダッシュボード非表示データ修正
- 🍎 App Store 500件制限案内追加

**v2.4**
- 🍎 App Store(iOS) レビュー収集対応
- 📄 CSV文書言語(KR/JP)反映

**v2.3**
- 🔍 ルールベース感情分析高度化
- 📄 レビューCSVダウンロード機能追加
- 📋 パッチノートUI追加
        '''
    }
    with st.expander('📋 패치 노트 / Patch Notes'):
        st.markdown(PATCH_NOTES[ui_code])

st.markdown(f'# {t["title"]}')
st.markdown('---')

# 플랫폼 선택
platform = st.radio(
    t['platform_label'],
    [t['platform_gp'], t['platform_as'], t['platform_st']],
    horizontal=True
)
is_appstore = (platform == t['platform_as'])
is_steam    = (platform == t['platform_st'])

if is_appstore:
    url_input = st.text_input(t['url_label'], placeholder=t['appstore_url_ph'])
elif is_steam:
    url_input = st.text_input(t['url_label'], placeholder=t['steam_url_ph'])
else:
    url_input = st.text_input(t['url_label'], placeholder=t['url_ph'])

col1, col2 = st.columns([1, 2])
with col1:
    mode = st.radio(t['mode_label'], [t['mode_count'], t['mode_period']], horizontal=True)
with col2:
    if mode == t['mode_count']:
        count_val = st.selectbox(t['count_label'], [100, 300, 500, 1000, 2000, 3000], index=3)
    else:
        dc1, dc2 = st.columns(2)
        with dc1: dt_from = st.date_input(t['from_label'], value=date.today()-timedelta(days=90))
        with dc2: dt_to   = st.date_input(t['to_label'],   value=date.today())

col3, _ = st.columns(2)
with col3:
    region_options = [t['region_KR'],t['region_JP'],t['region_US'],t['region_TW'],t['region_GB']]
    region_sel  = st.selectbox(t['region_label'], region_options)
    region_code = region_sel[:2]

if is_appstore:
    if ui_code == 'KR':
        st.info('🍎 앱스토어 모드 — 앱스토어 URL 또는 숫자 ID를 입력해주세요.\n\n⚠️ Apple RSS API 정책상 최대 **500건**까지만 수집 가능합니다.')
    else:
        st.info('🍎 App Storeモード — URLまたは数字IDを入力してください。\n\n⚠️ Apple RSS APIの制限により、最大 **500件** まで収集可能です。')
elif is_steam:
    st.info(t['steam_notice'])

st.markdown('---')
btn_start = st.button(t['btn_start'], use_container_width=True)

log_placeholder    = st.empty()
progress_bar       = st.empty()

if btn_start:
    is_count_mode = (mode == t['mode_count'])

    logs = []
    def add_log(msg):
        logs.append(msg)
        log_placeholder.markdown(f'<div class="log-box">{"<br>".join(logs[-30:])}</div>',unsafe_allow_html=True)

    if mode == t['mode_period']:
        if dt_from > dt_to: st.error(t['err_date']); st.stop()
        dt_from_dt = datetime(dt_from.year,dt_from.month,dt_from.day,0,0,0)
        dt_to_dt   = datetime(dt_to.year,dt_to.month,dt_to.day,23,59,59)

    target = count_val if is_count_mode else 99999
    prog = progress_bar.progress(0, text=t['prog_collect'])

    # ══════════════════════════
    # 🎮 스팀 수집
    # ══════════════════════════
    if is_steam:
        steam_id = parse_steam_id(url_input)
        if not steam_id: st.error(t['err_no_steamid']); st.stop()
        lang_st = STEAM_LANGS.get(region_code, 'koreana')
        add_log(f'🎮 스팀 수집 시작 | ID: {steam_id} | 언어: {lang_st} | 목표: {target if is_count_mode else "기간"}')
        prog.progress(10, text=t['prog_collect'])
        try:
            raw_reviews = fetch_steam_reviews(
                steam_id, language=lang_st,
                how_many=target,
                mode='count' if is_count_mode else 'period',
                dt_from=dt_from_dt if not is_count_mode else None,
                dt_to=dt_to_dt if not is_count_mode else None,
            )
        except Exception as e:
            st.error(f'❌ 수집 실패: {e}'); st.stop()
        add_log(t['log_done'].format(len(raw_reviews)))
        if not raw_reviews: st.error(t['err_no_data']); st.stop()
        prog.progress(80, text=t['prog_excel'])
        add_log(t['log_excel'])
        df = build_df_steam(raw_reviews)
        fname_prefix = f'steam_{steam_id}'
        excel_bytes = gen_excel_bytes(df, doc_code, is_steam=True)
        prog.progress(100, text=t['prog_done'])
        add_log(t['log_finish'])
        market_name = 'Steam'
        date_str = datetime.now().strftime('%y%m%d')
        fname_base = f'{fname_prefix}_{region_code}_{market_name}_{date_str}'
        st.session_state['result'] = {
            'df': df, 'excel_bytes': excel_bytes,
            'csv_bytes': gen_csv_bytes(df, doc_code),
            'fname_base': fname_base,
            'avg': df['평점'].mean(),
            'pos': int((df['평점']==5).sum()),
            'neg': int((df['평점']==1).sum()),
            'total': len(df),
            'is_steam': True,
        }

    # ══════════════════════════════
    # 🍎 앱스토어 수집
    # ══════════════════════════════
    elif is_appstore:
        as_id = parse_appstore_id(url_input)
        if not as_id: st.error(t['err_no_appid']); st.stop()
        pass  # requests는 기본 내장

        country_as = AS_REGIONS.get(region_code, 'us')
        add_log(f'🍎 앱스토어 수집 시작 | ID: {as_id} | 국가: {region_code}')
        try:
            how_many = target if is_count_mode else 500
            raw_reviews = fetch_appstore_reviews(as_id, country_as, how_many)
            add_log(f'✅ {len(raw_reviews):,}개 수집 완료')
        except Exception as e:
            add_log(f'❌ 수집 실패: {e}'); st.stop()

        # 기간 필터
        if not is_count_mode:
            filtered = []
            for rv in raw_reviews:
                at = rv.get('date')
                if isinstance(at, str):
                    try: at = datetime.strptime(at[:10], '%Y-%m-%d')
                    except: continue
                if at and at >= dt_from_dt and at <= dt_to_dt:
                    filtered.append(rv)
            raw_reviews = filtered
            add_log(f'📅 기간 필터 후 {len(raw_reviews):,}개')

        if not raw_reviews: st.error(t['err_no_data']); st.stop()
        prog.progress(80, text=t['prog_excel'])
        add_log(t['log_excel'])
        df = build_df_appstore(raw_reviews)
        fname_prefix = f'appstore_{as_id}'

    # ══════════════════════════════
    # 🤖 구글 플레이 수집
    # ══════════════════════════════
    else:
        app_id = parse_app_id(url_input)
        if not app_id: st.error(t['err_no_id']); st.stop()
        if not HAS_SCRAPER: st.error(t['err_no_pkg']); st.stop()

        lang_c, country_c = REGIONS[region_code]
        try:
            info = gp_app(app_id, lang=lang_c, country=country_c)
            add_log(t['log_app'].format(info.get('title',''),round(info.get('score',0),2),info.get('reviews',0)))
        except Exception as e:
            add_log(f'⚠️ {e}')

        all_r=[]; seen=set(); token=None; batch_num=0; stop=False; empty_streak=0

        while not stop:
            if is_count_mode and len(all_r)>=target: break
            ok=False; batch=[]
            for retry in range(1,4):
                try:
                    batch, token = reviews(app_id,lang=lang_c,country=country_c,
                        sort=Sort.NEWEST,count=200,continuation_token=token)
                    ok=True; break
                except Exception as e:
                    add_log(t['retry_msg'].format(retry,str(e)[:50])); time.sleep(retry*3)
            if not ok: add_log(t['fail_msg']); break
            if not batch:
                empty_streak+=1
                if empty_streak>=3: add_log(t['log_done'].format(len(all_r))); break
                time.sleep(2); continue
            empty_streak=0
            new_batch=[r for r in batch if r.get('reviewId') not in seen]
            for r in new_batch: seen.add(r.get('reviewId'))
            if not new_batch: break
            if not is_count_mode:
                filtered=[]
                for rv in new_batch:
                    at=rv.get('at')
                    if at is None: continue
                    rv_dt=at.replace(tzinfo=None) if (hasattr(at,'tzinfo') and at.tzinfo) else at
                    if rv_dt<dt_from_dt: stop=True; break
                    if rv_dt<=dt_to_dt: filtered.append(rv)
                all_r.extend(filtered)
                if stop: break
            else:
                all_r.extend(new_batch)
            batch_num+=1
            latest=batch[-1].get('at','')
            if hasattr(latest,'strftime'): latest=latest.strftime('%Y-%m-%d')
            add_log(t['log_collect'].format(len(all_r),batch_num,latest))
            pct=min(int(len(all_r)/target*80),80) if is_count_mode else min(batch_num*3,80)
            prog.progress(pct, text=f'{len(all_r):,}{t["unit_count"]} {t["prog_collect"]}')
            if is_count_mode and len(all_r)>=target: break
            if token is None:
                if is_count_mode and len(all_r)<target:
                    time.sleep(3); empty_streak+=1
                    if empty_streak>=3: break
                else: break
            time.sleep(1.0)

        all_r = all_r[:target] if is_count_mode else all_r
        collected = len(all_r)
        if collected<target and is_count_mode: add_log(t['log_short'].format(target,collected))
        else: add_log(t['log_done'].format(collected))
        if not all_r: st.error(t['err_no_data']); st.stop()

        prog.progress(85, text=t['prog_excel'])
        add_log(t['log_excel'])
        df = build_df(all_r)
        fname_prefix = app_id.split('.')[-1]

    excel_bytes = gen_excel_bytes(df, doc_code)
    prog.progress(100, text=t['prog_done'])
    add_log(t['log_finish'])

    # session_state에 결과 저장 (다운로드 후에도 유지)
    # #2 파일명: 게임이름_국가_마켓명_날짜
    market_name = 'AppStore' if is_appstore else 'GooglePlay'
    date_str = datetime.now().strftime('%y%m%d')
    fname_base = f'{fname_prefix}_{region_code}_{market_name}_{date_str}'
    st.session_state['_region_code'] = region_code
    st.session_state['result'] = {
        'df': df,
        'excel_bytes': excel_bytes,
        'csv_bytes': gen_csv_bytes(df, doc_code),
        'fname_base': fname_base,
        'avg': df['평점'].mean(),
        'pos': int((df['평점']>=4).sum()),
        'neg': int((df['평점']<=2).sum()),
        'total': len(df),
    }

# ── 결과 표시 (session_state 기반 — 다운로드 후에도 유지)
if 'result' in st.session_state:
    r = st.session_state['result']
    st.markdown('---')
    st.markdown(f'### {t["result_title"]}')
    m1,m2,m3,m4 = st.columns(4)
    m1.metric(t['metric_total'], f'{r["total"]:,}{t["unit_count"]}')
    if r.get('is_steam'):
        rec_rate = r["pos"]/r["total"]*100 if r["total"] else 0
        m2.metric('👍 추천률' if ui_code=='KR' else '👍 推薦率', f'{rec_rate:.1f}%')
        m3.metric('👍 추천' if ui_code=='KR' else '👍 推薦', f'{r["pos"]:,}{t["unit_count"]}')
        m4.metric('👎 비추천' if ui_code=='KR' else '👎 非推薦', f'{r["neg"]:,}{t["unit_count"]}')
    else:
        m2.metric(t['metric_avg'],   f'{r["avg"]:.2f} ★')
        m3.metric(t['metric_pos'],   f'{r["pos"]:,}{t["unit_count"]} ({r["pos"]/r["total"]*100:.1f}%)')
        m4.metric(t['metric_neg'],   f'{r["neg"]:,}{t["unit_count"]} ({r["neg"]/r["total"]*100:.1f}%)')
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(label=t['btn_dl'], data=r['excel_bytes'],
            file_name=f'{r["fname_base"]}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True)
    with dl2:
        st.download_button(label=t['btn_csv'], data=r['csv_bytes'],
            file_name=f'{r["fname_base"]}.csv',
            mime='text/csv',
            use_container_width=True)
