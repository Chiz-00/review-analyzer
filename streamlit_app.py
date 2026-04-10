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
        'region_label':'🌏 수집 국가',
        'btn_start':'▶  분석 시작',
        'btn_dl':'📥 엑셀 다운로드',
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
        'region_label':'🌏 収集国',
        'btn_start':'▶  分析開始',
        'btn_dl':'📥 Excelダウンロード',
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
    neg_tx=df[df['평점']<=2]['내용'].dropna()
    pos_tx=df[df['평점']>=4]['내용'].dropna()
    neg_total=len(neg_tx); pos_total=len(pos_tx)
    NEG_EXPR=['없애','별로','최악','짜증','불편','아쉽','문제','버그','오류',
              '싫','노잼','지루','힘들','망','안됨','안돼','못하','에러','튕',
              '렉','느려','발열','뻥','과금','현질','뽑기','비싸','천장','불만',
              '환불','삭제','망겜','없애셈','고쳐','해주세요','해줘요','개선해',
              'ㅡㅡ','ㅠ','ㅜ','갈증','낚이','실망','문의','뭡니까','말았다']
    POS_EXPR=['재밌','좋아','최고','훌륭','완벽','갓','꿀잼','대박','짱','추천',
              '만족','즐거','신나','감동','몰입','좋음','좋네','좋다','재미있',
              '흥미','멋지','예쁘','이쁘','퀄리티','매력']

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
                rl=raw.lower()
                has_neg=any(n in rl for n in NEG_EXPR)
                has_pos=any(p in rl for p in POS_EXPR)
                kv=any(k in raw for k in hit)
                if kv and has_pos and not has_neg and len(ex_best)<3: ex_best.append(raw[:90])
                elif kv and not has_neg and len(ex_good)<3: ex_good.append(raw[:90])
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

    used_neg=set(); used_pos=set()
    nr=cnt_neg(neg_tx, t['neg_kw'], used_neg)
    pr=cnt_pos(pos_tx, t['pos_kw'], used_pos)
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
    ws.cell(row=1,column=19,value=t['c_score']); ws.cell(row=1,column=20,value=t['rev_cnt'])
    for i,star in enumerate([1,2,3,4,5],2):
        ws.cell(row=i,column=19,value=f'{star}★'); ws.cell(row=i,column=20,value=int(dist.get(star,0)))
    ws.cell(row=8,column=19,value=''); ws.cell(row=8,column=20,value=t['rev_cnt'])
    for i,(k,v) in enumerate([(t['pie_pos'],int((df['평점']>=4).sum())),(t['pie_neu'],int((df['평점']==3).sum())),(t['pie_neg'],int((df['평점']<=2).sum()))],9):
        ws.cell(row=i,column=19,value=k); ws.cell(row=i,column=20,value=v)
    mo=df.groupby('작성월').agg(cnt=('평점','count'),avg=('평점','mean')).reset_index()
    ws.cell(row=13,column=19,value=''); ws.cell(row=13,column=20,value=t['rev_cnt']); ws.cell(row=13,column=21,value=t['avg_sc'])
    for i,r in enumerate(mo.itertuples(),14):
        ws.cell(row=i,column=19,value=r.작성월); ws.cell(row=i,column=20,value=r.cnt); ws.cell(row=i,column=21,value=round(r.avg,2))
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
    ws.cell(row=1,column=14,value=t['col_kw']); ws.cell(row=1,column=15,value=t['col_cnt'])
    for i,(lbl,d) in enumerate(ns[:6],2):
        ws.cell(row=i,column=14,value=lbl); ws.cell(row=i,column=15,value=d['count'])
    ws.cell(row=9,column=14,value=t['col_kw']); ws.cell(row=9,column=15,value=t['col_cnt'])
    for i,(lbl,d) in enumerate(ps[:5],10):
        ws.cell(row=i,column=14,value=lbl); ws.cell(row=i,column=15,value=d['count'])
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

def gen_excel_bytes(df, doc_lang):
    t=I18N[doc_lang]; wb=Workbook()
    ws1=wb.active; ws1.title=t['sh_dash']; mk_dash(ws1,df,t)
    ws2=wb.create_sheet(t['sh_op']);   mk_opinion(ws2,df,t,doc_lang)
    ws3=wb.create_sheet(t['sh_raw']);  mk_raw(ws3,df,t)
    ws4=wb.create_sheet(t['sh_stat']); mk_stats(ws4,df,t)
    ws1.sheet_properties.tabColor=C_ACCENT; ws2.sheet_properties.tabColor=C_GOLD
    ws3.sheet_properties.tabColor=C_GREEN;  ws4.sheet_properties.tabColor=C_BLUE
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

def parse_app_id(text):
    text=text.strip()
    m=re.search(r'id=([a-zA-Z0-9._]+)',text)
    if m: return m.group(1)
    if re.match(r'^[a-zA-Z][a-zA-Z0-9._]+$',text): return text
    return None

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

st.markdown(f'# {t["title"]}')
st.markdown('---')

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

st.markdown('---')
btn_start = st.button(t['btn_start'], use_container_width=True)

log_placeholder    = st.empty()
progress_bar       = st.empty()
result_placeholder = st.empty()

if btn_start:
    app_id = parse_app_id(url_input)
    if not app_id: st.error(t['err_no_id']); st.stop()
    if not HAS_SCRAPER: st.error(t['err_no_pkg']); st.stop()
    if mode == t['mode_period']:
        if dt_from > dt_to: st.error(t['err_date']); st.stop()
        dt_from_dt = datetime(dt_from.year,dt_from.month,dt_from.day,0,0,0)
        dt_to_dt   = datetime(dt_to.year,dt_to.month,dt_to.day,23,59,59)

    lang_c, country_c = REGIONS[region_code]
    logs = []
    def add_log(msg):
        logs.append(msg)
        log_placeholder.markdown(f'<div class="log-box">{"<br>".join(logs[-30:])}</div>',unsafe_allow_html=True)

    try:
        info = gp_app(app_id, lang=lang_c, country=country_c)
        add_log(t['log_app'].format(info.get('title',''),round(info.get('score',0),2),info.get('reviews',0)))
    except Exception as e:
        add_log(f'⚠️ {e}')

    all_r=[]; seen=set(); token=None; batch_num=0; stop=False; empty_streak=0
    is_count_mode = (mode == t['mode_count'])
    target = count_val if is_count_mode else 99999
    prog = progress_bar.progress(0, text=t['prog_collect'])

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
    excel_bytes = gen_excel_bytes(df, doc_code)
    prog.progress(100, text=t['prog_done'])
    add_log(t['log_finish'])

    with result_placeholder.container():
        st.markdown('---')
        st.markdown(f'### {t["result_title"]}')
        m1,m2,m3,m4 = st.columns(4)
        avg=df['평점'].mean(); pos=(df['평점']>=4).sum(); neg=(df['평점']<=2).sum()
        m1.metric(t['metric_total'], f'{len(df):,}{t["unit_count"]}')
        m2.metric(t['metric_avg'],   f'{avg:.2f} ★')
        m3.metric(t['metric_pos'],   f'{pos:,}{t["unit_count"]} ({pos/len(df)*100:.1f}%)')
        m4.metric(t['metric_neg'],   f'{neg:,}{t["unit_count"]} ({neg/len(df)*100:.1f}%)')
        fname=f'{app_id.split(".")[-1]}_review_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
        st.download_button(label=t['btn_dl'],data=excel_bytes,file_name=fname,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True)
