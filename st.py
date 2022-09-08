import streamlit as st

st.title('데이터 분석가를 꿈꾸는 장현우')
from PIL import Image

image = Image.open('jang.jpg')

st.image(image, width =200)

st.write("데이터 분석이도 심리학의 기반이라 생각하기에 이리로 오게 되었습니다.")

st.subheader("나를 소개합니다")
selected_item = st.radio("학력", ("2017. 04 ~ 2019. 02", "2009. 02 ~ 2013. 02", "2009"))

if selected_item == "2017. 04 ~ 2019. 02":
    st.write("학점은행제 심리학과")
elif selected_item == "2009. 02 ~ 2013. 02":
    st.write("송호대학교 호텔관광과")
    
    # imageh = Image.open('highs.jpg')
    # st.image(imageh, width =200)
    
elif selected_item == "2009":
    st.write("세명컴퓨터고등학교 전기과")
    
    
if st.button("경력"):
      st.write("유니에스  (2018. 01 ~ 재직중)")
      st.write(" 청년내일채움공제 담당업무를 하고 있습니다.(노무,사무,상담)")