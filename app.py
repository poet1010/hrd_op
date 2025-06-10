import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import re

# 페이지 설정
st.set_page_config(
    page_title="HRD 운영 실적 대시보드",
    page_icon="📊",
    layout="wide"
)

# 데이터 로드 함수
@st.cache_data
def load_data():
    try:
        # 데이터 로드
        df = pd.read_excel("24년_운영실적.xlsx", sheet_name="Sheet1")
        
        # # 데이터프레임 정보 출력
        # st.write("데이터프레임 컬럼:", df.columns.tolist())
        # st.write("데이터프레임 샘플:", df.head())
        
        # 필수 컬럼 확인
        required_columns = ['시작월', '카테고리1', '과정유형', '담당자', '과정명', 
                          '참석인원', '이수인원']
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"필수 컬럼이 누락되었습니다: {missing_columns}")
            return None
            
        # 시작월 데이터 처리
        def convert_month(month_str):
            try:
                # '01월' 형식의 문자열에서 숫자만 추출
                month_num = re.sub(r'[^0-9]', '', str(month_str))
                if month_num:
                    # 2024년을 기준으로 날짜 생성
                    return f"2024-{month_num.zfill(2)}"
                return "2024-01"  # 기본값 설정
            except:
                return "2024-01"  # 에러 발생 시 기본값 설정

        df['시작월'] = df['시작월'].apply(convert_month)
        
        # 데이터 타입 변환
        df['참석인원'] = pd.to_numeric(df['참석인원'], errors='coerce').fillna(0)
        df['이수인원'] = pd.to_numeric(df['이수인원'], errors='coerce').fillna(0)
        
        # 선택적 컬럼 처리
        optional_columns = {
            '과정만족도': 0,
            '현업적용율': 0,
            '교육일수': 0,
            '교육시간': 0
        }
        
        for col, default_value in optional_columns.items():
            if col not in df.columns:
                df[col] = default_value
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        return df
    except FileNotFoundError:
        st.error("데이터 파일을 찾을 수 없습니다. '24년_운영실적.xlsx' 파일이 프로젝트 루트 디렉토리에 있는지 확인해주세요.")
        return None
    except Exception as e:
        st.error(f"데이터 로드 중 오류가 발생했습니다: {str(e)}")
        return None

# 데이터 로드
df = load_data()

if df is None:
    st.stop()

# 사이드바 필터
st.sidebar.title("필터")

# 월 필터
months = sorted(df['시작월'].unique())
selected_month = st.sidebar.selectbox("시작월", ["전체"] + list(months))

# 카테고리 필터
categories1 = sorted(df['카테고리1'].unique())
selected_category1 = st.sidebar.selectbox("카테고리1", ["전체"] + list(categories1))

# 과정 유형 필터
course_types = sorted(df['과정유형'].unique())
selected_course_type = st.sidebar.selectbox("과정유형", ["전체"] + list(course_types))

# 담당자 필터
managers = sorted(df['담당자'].unique())
selected_manager = st.sidebar.selectbox("담당자", ["전체"] + list(managers))

# 데이터 필터링
filtered_df = df.copy()
if selected_month != "전체":
    filtered_df = filtered_df[filtered_df['시작월'] == selected_month]
if selected_category1 != "전체":
    filtered_df = filtered_df[filtered_df['카테고리1'] == selected_category1]
if selected_course_type != "전체":
    filtered_df = filtered_df[filtered_df['과정유형'] == selected_course_type]
if selected_manager != "전체":
    filtered_df = filtered_df[filtered_df['담당자'] == selected_manager]

# 메인 타이틀
st.title("HRD 운영 실적 대시보드")

# 1. 전체 운영 실적 개요 (KPI 카드)
st.header("1. 전체 운영 실적 개요")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("총 과정 수", len(filtered_df))
    st.metric("총 참석인원", filtered_df['참석인원'].sum())

with col2:
    st.metric("총 이수인원", filtered_df['이수인원'].sum())
    completion_rate = (filtered_df['이수인원'].sum() / filtered_df['참석인원'].sum() * 100) if filtered_df['참석인원'].sum() > 0 else 0
    st.metric("전체 수료율", f"{completion_rate:.1f}%")

with col3:
    if filtered_df['과정만족도'].sum() > 0:
        st.metric("평균 과정 만족도", f"{filtered_df['과정만족도'].mean():.1f}")
    if filtered_df['현업적용율'].sum() > 0:
        st.metric("평균 현업 적용율", f"{filtered_df['현업적용율'].mean():.1f}%")

with col4:
    if filtered_df['교육일수'].sum() > 0:
        st.metric("총 교육 일수", filtered_df['교육일수'].sum())
    if filtered_df['교육시간'].sum() > 0:
        st.metric("총 교육 시간", filtered_df['교육시간'].sum())

# 2. 월별 실적 분석
st.header("2. 월별 실적 분석")

# 월별 데이터 집계
monthly_data = filtered_df.groupby('시작월').agg({
    '참석인원': 'sum',
    '이수인원': 'sum',
    '과정만족도': 'mean',
    '현업적용율': 'mean'
}).reset_index()

# 월별 참석/이수 인원 추이
fig1 = px.line(monthly_data, x='시작월', y=['참석인원', '이수인원'],
               title='월별 참석/이수 인원 추이',
               labels={'value': '인원 수', 'variable': '구분'})
st.plotly_chart(fig1, use_container_width=True)

# 월별 만족도/적용율 추이 (데이터가 있는 경우에만)
if monthly_data['과정만족도'].sum() > 0 or monthly_data['현업적용율'].sum() > 0:
    fig2 = px.line(monthly_data, x='시작월', y=['과정만족도', '현업적용율'],
                   title='월별 만족도/적용율 추이',
                   labels={'value': '비율 (%)', 'variable': '구분'})
    st.plotly_chart(fig2, use_container_width=True)

# 3. 카테고리별 실적 분석
st.header("3. 카테고리별 실적 분석")

# 성과지표 컬럼 미존재시 NaN으로 추가 및 숫자형 변환 (fillna(0) 제거)
score_5 = ['과정만족도', '교육내용', '교육방법']
score_pct = ['긍정응답율', '과정NPS', '현업적용']
성과지표 = score_5 + score_pct
for col in 성과지표:
    if col not in filtered_df.columns:
        filtered_df[col] = np.nan
    filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce')

# 카테고리별 데이터 집계
category_data = filtered_df.groupby('카테고리1').agg({
    '참석인원': 'sum',
    '이수인원': 'sum',
    '과정만족도': 'mean',
    '교육내용': 'mean',
    '교육방법': 'mean',
    '긍정응답율': 'mean',
    '과정NPS': 'mean',
    '현업적용': 'mean'
}).reset_index()

# 카테고리별 참석인원
fig3 = px.bar(category_data, x='카테고리1', y='참석인원',
              title='카테고리별 참석인원',
              labels={'참석인원': '참석인원 수'})
st.plotly_chart(fig3, use_container_width=True)

# 카테고리별 과정성과 지표 바차트
for col in score_5:
    if col in category_data.columns:
        if (category_data[col].notna().sum() > 0 and category_data[col].sum() > 0):
            fig = px.bar(
                category_data, x='카테고리1', y=col,
                title=f'카테고리별 {col} (평균, 5점 척도)',
                labels={col: f'{col} (0~5점)'},
                range_y=[2, 5]
            )
            # 바 상단에 값 표시
            fig.update_traces(
                text=category_data[col].round(2),
                textposition='outside'
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"'{col}' 데이터가 없습니다.")

for col in score_pct:
    if col in category_data.columns:
        # 0~1 스케일을 0~100%로 변환
        category_data[f'{col}_pct'] = category_data[col] * 100
        if (category_data[col].sum() > 0):
            fig = px.bar(category_data, x='카테고리1', y=f'{col}_pct',
                         title=f'카테고리별 {col} (평균, %)',
                         labels={f'{col}_pct': f'{col} (%)'},
                         range_y=[0,100])
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"'{col}' 데이터가 없습니다.")

# 4. 과정 유형별 실적 분석
st.header("4. 과정 유형별 실적 분석")

# 과정 유형별 데이터 집계
course_type_data = filtered_df.groupby('과정유형').agg({
    '참석인원': 'sum',
    '이수인원': 'sum',
    '과정만족도': 'mean',
    '현업적용율': 'mean'
}).reset_index()

# 과정 유형별 참석인원 비중
fig5 = px.pie(course_type_data, values='참석인원', names='과정유형',
              title='과정 유형별 참석인원 비중')
st.plotly_chart(fig5, use_container_width=True)

# 과정 유형별 수료율
course_type_data['수료율'] = (course_type_data['이수인원'] / course_type_data['참석인원'] * 100)
fig6 = px.bar(course_type_data, x='과정유형', y='수료율',
              title='과정 유형별 수료율',
              labels={'수료율': '수료율 (%)'})
st.plotly_chart(fig6, use_container_width=True)

# 5. 담당자별 실적 분석
st.header("5. 담당자별 실적 분석")

# 담당자별 데이터 집계
manager_data = filtered_df.groupby('담당자').agg({
    '과정명': 'count',
    '참석인원': 'sum',
    '이수인원': 'sum',
    '과정만족도': 'mean',
    '현업적용율': 'mean'
}).reset_index()

# 담당자별 관리 과정 수
fig7 = px.bar(manager_data, x='담당자', y='과정명',
              title='담당자별 관리 과정 수',
              labels={'과정명': '과정 수'})
st.plotly_chart(fig7, use_container_width=True)

# 담당자별 수료율
manager_data['수료율'] = (manager_data['이수인원'] / manager_data['참석인원'] * 100)
fig8 = px.bar(manager_data, x='담당자', y='수료율',
              title='담당자별 수료율',
              labels={'수료율': '수료율 (%)'})
st.plotly_chart(fig8, use_container_width=True)

# 상세 데이터 테이블
st.header("상세 데이터")
st.dataframe(filtered_df, use_container_width=True) 