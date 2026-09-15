import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os

# Matplotlib 캐시 디렉토리 내 fontlist 파일 삭제 후 재생성
cache_dir = fm.get_cachedir()
for file in os.listdir(cache_dir):
    if 'fontlist' in file:
        try:
            os.remove(os.path.join(cache_dir, file))
        except:
            pass

# 나눔고딕 또는 맑은 고딕 설정 예시
plt.rc('font', family='Malgun Gothic') # 윈도우인 경우
# plt.rc('font', family='NanumGothic') # 리눅스인 경우
plt.rcParams['axes.unicode_minus'] = False
