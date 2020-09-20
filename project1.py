# 필요한 모듈과 라이브러리를 로딩
from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd    
import time, os, math, random
import cx_Oracle


# 사용자에게 크롤링 방식 묻기
print("=" *80)
print("         쿠팡 크롤러입니다.")
print("=" *80)
print('\n')
print(''' 1. 카테고리별 BEST SELLER       2. 키워드 검색''')
print('\n')
choice = int(input('0.어떤 방식으로 크롤링을 수행하시겠습니까? : '))
print("=" *80)
print('\n')

# 크롤링할 데이터를 바로 DB에 집어넣을 것인지, 우선 보류 후 나중에 결정할지를 묻기
print('''0.크롤링 데이터를 데이터베이스에 저장하시겠습니까?
  => 저장한다(Y), 저장하지 않는다(N), 데이터를 보고 결정(L)''')

# 위에서 입력한 값에서 지정된 키 이외의 값을 입력했을 경우 DB에 저장하지 않는 것을 DEFAULT 값으로
db_save = str(input('1-1.DB 저장 여부를 입력해주세요 (Y/N/L) : '))
if db_save == 'y' or db_save == 'Y' :
    print('입력값 : ',db_save,' => 크롤링 데이터를 DB에 저장합니다.')
elif  db_save == 'l' or db_save == 'L' :
    print('입력값 : ',db_save,' => DB 저장을 보류합니다.')
else : 
    print('입력값 : ',db_save,' => 크롤링 데이터를 파일로만 추출합니다.')
    db_save = 'N'

# DB 저장을 선택했을 경우, 데이터를 저장할 TABLE 이름과 그 TABLE을 구성
if db_save == 'Y' or db_save == 'y':

    db_name = str(input('''1-2.데이터를 저장할 table명을 입력해주세요. (DEFAULT TABLE : CRAWLING_TABLE) :'''))
    if db_name == "" : 
        db_name = "CRAWLING_TABLE"

    print('\n')
    print('입력된 테이블명 : ',db_name)
    print('='*80)

    # 제약조건의 경우 재고와 크롤링 날짜는 디폴트값을 부여하고, 다른 값들은 NOT NULL 값 부여
    create_table = '''CREATE TABLE  '''+ db_name + ''' (
                            RANKING NUMBER(4,0)  CONSTRAINT ranking_nn NOT NULL, 
                            TITLE VARCHAR2(256) CONSTRAINT title_nn NOT NULL,
                            P_PRICE NVARCHAR2(10) CONSTRAINT p_price_nn NOT NULL,
                            O_PRICE NVARCHAR2(10) CONSTRAINT o_price_nn NOT NULL,
                            DISCOUNT VARCHAR2(4) CONSTRAINT discount_nn NOT NULL,
                            DELIVERY VARCHAR2(8) CONSTRAINT delivery_nn NOT NULL,
                            STOCK VARCHAR2(8) DEFAULT '재고있음',
                            REVIEW NUMBER(6) CONSTRAINT review_nn NOT NULL,
                            STARS NUMBER(6,1) CONSTRAINT stars_nn NOT NULL,
                            CATEGORY VARCHAR2(20) CONSTRAINT category_nn NOT NULL,
                            CRAWLED DATE DEFAULT sysdate          
                            )''' 

    print('==> 오라클과 연결합니다.')
    print('\n')

    # 오라클 서버의 scott 계정으로 접속
    con = cx_Oracle.connect('scott/tiger@DESKTOP-0ULPS36/testdb')
    cur = con.cursor()
  
    print('===> 연결 완료. 테이블 생성을 시작합니다.')
    print('='*80)

    # 위의 조건대로 테이블을 생성하되 기존에 같은 이름의 테이블이 있을 경우 데이터를 추가할지,
    # 기존 테이블을 drop하고 새로 만들지를 묻기
    try : cur.execute(create_table)
    except :
        option = input('''1-3.이미 테이블이 존재합니다. 
        데이터를 추가하거나(Y), 기존 테이블을 삭제하고 새로 만듭니다(N) : ''' )
        if option == 'N' or option == 'n':
            cur.execute('DROP TABLE '+ db_name)
            
            # 데이터를 반복적으로 수집하는 과정에서 간혹 발생하는 오라클 접속 오류를 조정
            try : cur.execute(create_table)
            except :
                print("1-4.오라클 연결 시드를 재조정합니다.")
                cur.execute('SET oracle_sid=testdb')
                con = cx_Oracle.connect('scott/tiger@DESKTOP-0ULPS36/testdb')
                cur = con.cursor()

                cur.execute(create_table)

        else : # 실수로 다른 것을 입력했을 때 기존 데이터가 삭제되는 경우를 대비해서 디폴트값을 붙여쓰기로 지정
            pass

    print('====> 테이블 생성 완료, 크롤링을 시작합니다.')        
    print("=" *80)
    print('\n')

else : pass

# 카테고리별 베스트 셀러 추출 방식 선택
if choice == 1 : 

    # 크롤링할 카테고리 보여주고 선택
    print("=" *80)
    print(" 쿠팡 사이트의 카테고리별 Best Seller 상품 정보 추출하기 ")
    print("=" *80)
    print('\n')
    print('''
    1.패션의류/잡화 - 남성     2.패션의류/잡화 - 여성      3.뷰티               4.출산/유아동
    5.식품                     6.주방용품                  7.생활용품           8.홈인테리어
    9.가전/디지털              10.스포츠/레저              11.자동차용품        12.도서/음반/DVD
    13.완구/취미               14.문구/오피스              15.반려동물용품      16.헬스/건강식품
    ''')
    print('\n')

    # ctg index를 벗어나는 경우에 한해 무한루프문 구성
    while True : 
        try :
            ctg = int(input('2-1.위 분야 중에서 자료를 수집할 분야의 번호를 선택하세요 :'))
            if ctg > 16 or ctg < 1 : pass
            else : break
        except :
            continue

    # 한 번에 최대한 많은 양의 크롤링을 수행하기 위해 한 페이지당 게시글 수를 높여 회당 1080개로 최대화
    # 또한 역으로 1080개가 넘는 숫자를 입력했을 경우 이를 1080개로 조정하는 문구와 함께 cnt값 조정
    while True : 
        try :
            cnt = int(input('2-2.크롤링 할 건수는 몇건입니까? (최대 1,080개): '))
            if cnt > 1080 : 
                print("요청하신 건수가 최대건수보다 많습니다. 1080개로 조정합니다.")
                cnt = 1080
            break
        except :
            continue

    # 페이지당 게시물 수로 크롤링할 페이지 계싼  
    page_cnt = math.ceil(cnt/120)
    
    f_dir = input("2-3.파일을 저장할 폴더명만 쓰세요(기본경로:c:\\temp\\project1\\):")
    if f_dir == '' :
        f_dir = "c:\\temp\\project1\\"
        
    print("\n")

    print("데이터 크롤링을 시작합니다.")

    # 총 작업에 걸리는 시간 및 작업 시간별로 dir를 만들어서 고유의 dir 생성
    n = time.localtime()
    s = '%04d-%02d-%02d' % (n.tm_year, n.tm_mon, n.tm_mday)
    s1 = '%02d-%02d' % (n.tm_hour, n.tm_min)
    s_time = time.time( )

    # 웹사이트 접속 후 해당 메뉴로 이동
    chrome_path = "c:/temp/python/chromedriver.exe"
    driver = webdriver.Chrome(chrome_path)
    query_url='https://www.coupang.com/'
    driver.get(query_url)
    driver.implicitly_wait(5)

    # 분야별 더보기 버튼을 눌러 페이지 열기
    # 페이지를 눌러서 들어가는 과정에서 여러 요인에 의해 멈추는 경우 생김
    # => 오류 방지를 위해 try구문으로 페이지 최신화 및 다시 수행 명령
    try : 
        time.sleep(random.randint(1,4))
        driver.find_element_by_xpath("""//*[@id="header"]/div""").click()
        time.sleep(random.randint(2,4))
    except :
        driver.refresh()
        time.sleep(random.randint(1,4))
        driver.find_element_by_xpath("""//*[@id="header"]/div""").click()
        time.sleep(random.randint(2,4))



    # 남성패션(1번)과 여성패션(2번)의 경우, 패션 항목을 눌러서 들어가는 경로가 추가적으로 필요
    if ctg < 3 :
        driver.find_element_by_xpath('''//*[@id="gnbAnalytics"]/ul[1]/li[1]/a''').click()
        time.sleep(1)

    if ctg == 1 : #남성패션
        ctg1 = '''//*[@id="gnbAnalytics"]/ul[1]/li[1]/div/div/ul/li[2]/a'''
    elif ctg == 2 : #여성패션
        ctg1 = '''//*[@id="gnbAnalytics"]/ul[1]/li[1]/div/div/ul/li[1]/a'''
    else : #그 외의 경우 xpath가 카테고리의 순서에 따라 순차적으로 나열되어 있으므로 이를 반영
        ctg1 = '''//*[@id="gnbAnalytics"]/ul[1]/li['''+ str(int(ctg-1)) +''']/a'''

    #어느 페이지에서든 오버랩되어 표시되는 페이지 번호를 누를 수 있도록 크롤링 다운 함수를 정의
    def scroll_down(driver):
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
        time.sleep(1)

    #카테고리를 눌러 들어가는 과정에서도 에러 방지용 try구문
    try : 
        driver.find_element_by_xpath(ctg1).click( )
        driver.implicitly_wait(3)
    except : 
        time.sleep(1)
        driver.find_element_by_xpath(ctg1).click( )
        driver.implicitly_wait(3)


    # 쿠팡의 경우, 카테고리가 60개로 볼 때보다 120개로 볼때보다 총 게시물이 더 많이 나옴
    # 120개씩 보기로 변경 - 60개 더블 클릭 후 120개 클릭
    try :
        time.sleep(random.randint(1,3))
        scroll_down(driver)
        driver.find_element_by_class_name('selectbox-options').click()
        time.sleep(0.5)
        driver.find_element_by_class_name('selectbox-options').click()
        time.sleep(1)
        driver.find_element_by_xpath('''//*[@id="searchSortingList"]/ul/li[2]''').click()
        driver.implicitly_wait(3)
        time.sleep(random.randint(1,3))
    except :
        driver.refresh()
        time.sleep(1)
        scroll_down(driver)
        driver.find_element_by_class_name('selectbox-options').click()
        time.sleep(0.5)
        driver.find_element_by_class_name('selectbox-options').click()
        time.sleep(1)
        driver.find_element_by_xpath('''//*[@id="searchSortingList"]/ul/li[2]''').click()
        driver.implicitly_wait(3)
        time.sleep(random.randint(1,3))


    # 저장될 파일 경로와 이름을 지정
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    sec = soup.find('div',class_='search-header').find('ul'>'li').find_all('a')[1].get_text()
    sec_name = str(sec.replace('/','+').split())
    sec_name = sec_name.replace('[','').replace(']','').replace("'","")
    query_txt='쿠팡'
    print('scout3', sec_name)

    try : 
        os.makedirs(f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name)
    except : pass
    else : os.chdir(f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name)

    ff_dir=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name
    ff_name=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name+'\\'+s+'-'+query_txt+'-'+sec_name+'.txt'
    fc_name=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name+'\\'+s+'-'+query_txt+'-'+sec_name+'.csv'
    fx_name=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name+'\\'+s+'-'+query_txt+'-'+sec_name+'.xls'

    # 제품 이미지 저장용 폴더 생성
    img_dir = ff_dir+"\\images"
    os.makedirs(img_dir)
    os.chdir(img_dir)

    # 내용을 수집
    print("\n")
    print("===== 곧 수집된 결과를 출력합니다 ^^ ===== ")
    print("\n")

    ranking2=[]        #제품의 판매순위 저장
    title2=[]          #제품 정보 저장
    p_price2=[]        #현재 판매가 저장
    original2 = []     #원가 저장
    discount2 = []     #할인율 저장
    rocket2 = []       #로켓배송여부
    out2 = []          #품절여부
    sat_count2=[]      #상품평 수 저장
    stars2 = []        #상품평점 저장
    category2 = []     #크롤링 분류


    file_no = 0   # 이미지 파일 저장할 때 번호
    count = 1     # 총 게시물 건수 카운트 변수


    scroll_down(driver)   # 현재화면의 가장 아래로 스크롤다운

    #본격적인 크롤링 시작, 각 페이지 별로 파싱해서 소스를 넣어주고 각 게시물 별로 크롤링
    for x in range(1,page_cnt + 1) :
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        item_result = soup.find('ul','baby-product-list')
        item_result2 = item_result.find_all('li')

        for li in item_result2:

            if cnt < count :
                break

            # 제품 이미지 다운로드
            import urllib.request
            import urllib
            
            try :
                photo = li.find('dt','image').find('img')['src']
            except AttributeError :
                continue

            file_no += 1
            full_photo = 'https:' + photo
            urllib.request.urlretrieve(full_photo,str(file_no)+'.jpg')
            time.sleep(0.5)

            if cnt < file_no :
                break

            #제품 내용 추출하며 작업 현황 print 및 메모장에 저장
            f = open(ff_name, 'a',encoding='UTF-8')
            f.write("-----------------------------------------------------"+"\n")
            print("-" *70)

            ranking = count
            print("1.판매순위:",ranking)
            f.write('1.판매순위:'+ str(ranking) + "\n")

            try :
                t = li.find('div',class_='name').get_text()
            except AttributeError :
                title = '제품소개가 없습니다'
                print(title.replace("\n",""))
                f.write('2.제품소개:'+ title + "\n")
            else :
                title = t.replace("\n","").strip()
                print("2.제품소개:", title.replace("\n",""))                  
                f.write('2.제품소개:'+ title + "\n")

            try :
                p_price = li.find('strong','price-value').get_text()
            except :
                p_price = '0'
                print("3.판매가격:", p_price.replace("\n",""))
                f.write('3.판매가격:'+ p_price + "\n")
            else :
                print("3.판매가격:", p_price.replace("\n",""))
                f.write('3.판매가격:'+ p_price + "\n")

            try :
                original = li.find('del','base-price').get_text()
            except  :
                original = p_price
                print("4:원래가격:", original.replace("\n",""))
                f.write('4.원래가격:'+ original + "\n")
            else :
                print("4:원래가격:", original.replace("\n",""))
                f.write('4.원래가격:'+ original + "\n")

            # 명시된 할인율을 가져오되, 명시되어 있지 않은 경우 자체적으로 계산
            try :
                discount = li.find('span','discount-percentage').get_text()
            except  :
                try : 
                    discount = (((original - p_price)/original))*100 
                    discount = str(round(discount,1))
                except :
                    discount = '0 %'
                print("5:할인율:", discount)
                f.write('5.할인율:'+ discount + "\n")
            else :
                print("5:할인율:", discount)
                f.write('5.할인율:'+ discount + "\n")

            try :
                rocket = li.find('span', class_="badge rocket").find('img')['alt']
                rocket = str(rocket)
            except  AttributeError :
                rocket= "일반배송"
                print('6.로켓배송여부: ',rocket)
                f.write('6. 로켓배송여부:'+ rocket + "\n")
            else :
                print('6.로켓배송여부:',rocket)
                f.write('6. 로켓배송여부:'+ rocket + "\n")

            try :
                out = li.find('div','out-of-stock').get_text()
            except AttributeError :
                out= "재고있음"
                print('7.품절여부: ',out)
                f.write('7.품절여부:'+ out + "\n")
            else :
                out = "품절"
                print('7.품절여부:',out)
                f.write('7.품절여부:'+ out + "\n")

            try :
                sat_count_1 = li.find('span','rating-total-count').get_text()
                sat_count_2 = sat_count_1.replace("(","").replace(")","")
            except  :
                sat_count_2='0'
                print('8.상품평 수: ',sat_count_2)
                f.write('8.상품평 수:'+ sat_count_2 + "\n")
            else :
                print('8.상품평 수:',sat_count_2)
                f.write('8.상품평 수:'+ sat_count_2 + "\n")

            try :
                stars1 = li.find('em','rating').get_text()
            except  :
                stars1='0'
                print('9.상품평점: ',stars1)
                f.write('9.상품평점:'+ stars1 + "\n")
            else :
                print('9.상품평점:',stars1)
                f.write('9.상품평점:'+ stars1 + "\n")

            category = sec_name
            print("10.카테고리:",category)
            f.write('10.카테고리:'+ str(category) + "\n")


            print("-" *70)

            #상단광고의 경우 구분(각 페이지별 상위 4개 항목에 대한 리뷰 수가 50개 이하인 경우 상단노출 광고로 간주)
            #이를 통해 추후 상단 광고로 노출된 상품이 어느정도의 광고효과가 있었는지 추적 가능
            if count % 120 < 5 and int(sat_count_2) < 51 :
                title = title + '   *** 상단광고 ***'

            f.close( )             
            time.sleep(0.5)

            #추출한 데이터를 리스트화
            ranking2.append(ranking)
            title2.append(title.replace("\n",""))

            p_price2.append(p_price.replace("\n",""))
            original2.append(original.replace("\n",""))
            discount2.append(discount.replace("\n",""))

            rocket2.append(rocket.replace("\n",""))
            out2.append(out)

            #간혹 indexerror가 발생하는 부분에선 다시금 에러 방지구문
            try :   
                sat_count2.append(sat_count_2)
            except IndexError :
                sat_count2.append(0)

            try :   
                stars2.append(stars1)
            except IndexError :
                stars2.append(0)

            category2.append(str(category))


            count += 1
        x += 1        
        scroll_down(driver)
        scroll_down(driver)

        # 정상작업 이후 이 부분에서 잦은 오류가 발생하므로 이를 방지하기 위해 패스 구문 적용
        try : driver.find_element_by_link_text('%s' %x).click() # 다음 페이지번호 클릭 
        except : pass

# 키워드 검색의 경우
elif choice == 2 :

    #검색할 키워드 입력받기
    keyword = str(input("2-1.크롤링할 키워드를 지정해주세요 : "))
    
    #크롤링할 건수를 숫자 이외의 글자로 잘못입력했을 시 반복 수행
    while True :
        try: 
            cnt = int(input('2-2.크롤링 할 건수는 몇건입니까?: '))
            break
        except : pass
    
    #직접 검색 방식으로 크롤링할 경우에는 한 페이지당 최대 게시물이 72개.
    #동시에 72개씩 볼때 가장 많은 게시물을 크롤링 가능
    page_cnt = math.ceil(cnt/72)
    f_dir = input("2-3.파일을 저장할 폴더명을 지정해주세요(기본경로:c:\\temp\\project1\\):")
    if f_dir == '' :
        f_dir = "c:\\temp\\project1\\"
        
    print("\n")

    print("데이터 크롤링을 시작합니다.")

    # 작업 시간과 고유 dir 등 생성
    n = time.localtime()
    s = '%04d-%02d-%02d' % (n.tm_year, n.tm_mon, n.tm_mday)
    s1 = '%02d-%02d' % (n.tm_hour, n.tm_min)
    s_time = time.time( )

    # 웹사이트 접속 후 해당 메뉴로 이동
    chrome_path = "c:/temp/python/chromedriver.exe"
    driver = webdriver.Chrome(chrome_path)
    query_url= ('https://www.coupang.com/np/search?component=&q={}&channel=user'. format(keyword))
    driver.get(query_url)
    driver.implicitly_wait(5)

    # 키워드 검색의 경우 게시물 보기 방식이 클릭이 아닌 가져다대는 방식으로 actionchains 사용
    # 72개씩 보기로 변경 - 액션 체인으로 마우스 가져다대고 클릭
    # 마찬가지로 페이지 에러 방지용 새로고침 try구문
    try : 
        time.sleep(random.randint(1,3))
        from selenium.webdriver.common.action_chains import ActionChains
        action = ActionChains(driver)
        selectbox = driver.find_element_by_class_name('selectbox-options')
        action.move_to_element(selectbox).perform()
        time.sleep(0.5)
        driver.find_element_by_xpath('''//*[@id="searchSortingList"]/ul/li[4]/label''').click()
        driver.implicitly_wait(3)
        time.sleep(random.randint(1,3))
    except :
        driver.refresh()
        time.sleep(1)
        
        from selenium.webdriver.common.action_chains import ActionChains
        action = ActionChains(driver)
        selectbox = driver.find_element_by_class_name('selectbox-options')
        action.move_to_element(selectbox).perform()
        time.sleep(0.5)
        driver.find_element_by_xpath('''//*[@id="searchSortingList"]/ul/li[4]/label''').click()
        driver.implicitly_wait(3)
        time.sleep(random.randint(1,3))


    # 저장될 파일 경로와 이름을 지정
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    sec_name = str(keyword)
    query_txt='쿠팡'


    try : 
        os.makedirs(f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name)
    except : pass
    else : os.chdir(f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name)

    ff_dir=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name
    ff_name=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name+'\\'+s+'-'+query_txt+'-'+sec_name+'.txt'
    fc_name=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name+'\\'+s+'-'+query_txt+'-'+sec_name+'.csv'
    fx_name=f_dir+s+'\\'+s1+'\\'+query_txt+'-'+sec_name+'\\'+s+'-'+query_txt+'-'+sec_name+'.xls'

    # 제품 이미지 저장용 폴더 생성
    img_dir = ff_dir+"\\images"
    os.makedirs(img_dir)
    os.chdir(img_dir)

    # 내용을 수집
    print("\n")
    print("===== 곧 수집된 결과를 출력합니다 ^^ ===== ")
    print("\n")

    ranking2=[]        #제품의 판매순위 저장
    title2=[]          #제품 정보 저장
    p_price2=[]        #현재 판매가 저장
    original2 = []     #원가 저장
    discount2 = []     #할인율 저장
    rocket2 = []       #로켓배송여부
    out2 = []          #품절여부
    sat_count2=[]      #상품평 수 저장
    stars2 = []        #상품평점 저장
    category2 = []     #크롤링 분류

    file_no = 0   # 이미지 파일 저장할 때 번호
    count = 1     # 총 게시물 건수 카운트 변수


    #각 페이지별 소스를 파싱해서 게시글 단위로 크롤링
    for x in range(1,page_cnt + 1) :
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        item_result = soup.find('div','search-content search-content-with-feedback')
        item_result2 = item_result.find_all('li')

        for li in item_result2:

            if cnt < count :
                break

            # 제품 이미지 다운로드
            import urllib.request
            import urllib
            
            try :
                photo = li.find('dt','image').find('img')['src']
            except AttributeError :
                continue

            file_no += 1
            full_photo = 'https:' + photo
            urllib.request.urlretrieve(full_photo,str(file_no)+'.jpg')
            time.sleep(0.5)

            if cnt < file_no :
                break

            #제품 내용 추출
            f = open(ff_name, 'a',encoding='UTF-8')
            f.write("-----------------------------------------------------"+"\n")
            print("-" *70)

            ranking = count
            print("1.판매순위:",ranking)
            f.write('1.판매순위:'+ str(ranking) + "\n")

            try :
                t = li.find('div',class_='name').get_text().replace("\n","")
            except AttributeError :
                title = '제품소개가 없습니다'
                print(title.replace("\n",""))
                f.write('2.제품소개:'+ title + "\n")
            else :
                title = t.replace("\n","").strip()
                print("2.제품소개:", title.replace("\n","").strip())                  
                f.write('2.제품소개:'+ title + "\n")

            try :
                p_price = li.find('strong','price-value').get_text()
            except :
                p_price = '0'
                print("3.판매가격:", p_price.replace("\n",""))
                f.write('3.판매가격:'+ p_price + "\n")
            else :
                print("3.판매가격:", p_price.replace("\n",""))
                f.write('3.판매가격:'+ p_price + "\n")

            try :
                original = li.find('del','base-price').get_text()
            except  :
                original = p_price
                print("4:원래가격:", original.replace("\n",""))
                f.write('4.원래가격:'+ original + "\n")
            else :
                print("4:원래가격:", original.replace("\n",""))
                f.write('4.원래가격:'+ original + "\n")

            # 할인율을 가져오되, 할인율이 명시되어 있지 않은 경우 실제 가격과 비교해서 할인율을 산출하는 식을 추가
            try :
                discount = li.find('span','instant-discount-rate').get_text().replace("\n","")
            except  :
                try : 
                    discount = (((original - p_price)/original))*100 
                    discount = str(round(discount,1))
                except :
                    discount = '0 %'
                print("5:할인율:", discount)
                f.write('5.할인율:'+ discount + "\n")
            else :
                print("5:할인율:", discount)
                f.write('5.할인율:'+ discount + "\n")

            try :
                rocket = li.find('span', class_="badge rocket").find('img')['alt']
                rocket = str(rocket)
            except  AttributeError :
                rocket= "일반배송"
                print('6.로켓배송여부: ',rocket)
                f.write('6. 로켓배송여부:'+ rocket + "\n")
            else :
                print('6.로켓배송여부:',rocket)
                f.write('6. 로켓배송여부:'+ rocket + "\n")

            try :
                out = li.find('div','out-of-stock').get_text()
            except AttributeError :
                out= "재고있음"
                print('7.품절여부: ',out)
                f.write('7.품절여부:'+ out + "\n")
            else :
                out = "품절"
                print('7.품절여부:',out)
                f.write('7.품절여부:'+ out + "\n")

            try :
                sat_count_1 = li.find('span','rating-total-count').get_text()
                sat_count_2 = sat_count_1.replace("(","").replace(")","")
            except  :
                sat_count_2='0'
                print('8.상품평 수: ',sat_count_2)
                f.write('8.상품평 수:'+ sat_count_2 + "\n")
            else :
                print('8.상품평 수:',sat_count_2)
                f.write('8.상품평 수:'+ sat_count_2 + "\n")

            try :
                stars1 = li.find('em','rating').get_text()
            except  :
                stars1='0'
                print('9.상품평점: ',stars1)
                f.write('9.상품평점:'+ stars1 + "\n")
            else :
                print('9.상품평점:',stars1)
                f.write('9.상품평점:'+ stars1 + "\n")

            category = sec_name
            print("10.카테고리:",category)
            f.write('10.판매순위:'+ str(category) + "\n")


            print("-" *70)

            #상단광고의 경우 구분(각 페이지별 상위 4개 항목에 대한 리뷰 수가 50개 이하인 경우 상단노출 광고로 간주)
            #이를 통해 추후 상단 광고로 노출된 상품이 어느정도의 광고효과가 있었는지 추적 가능
            if count % 72 < 5 and int(sat_count_2) < 51 :
                title = title + '   *** 상단광고 ***'

            f.close( )             
            time.sleep(0.5)

            #추출한 데이터를 리스트화
            ranking2.append(ranking)
            title2.append(title.replace("\n",""))

            p_price2.append(p_price.replace("\n",""))
            original2.append(original.replace("\n",""))
            discount2.append(discount.replace("\n",""))

            rocket2.append(rocket.replace("\n",""))
            out2.append(out)

            try :   
                sat_count2.append(sat_count_2)
            except IndexError :
                sat_count2.append(0)

            try :   
                stars2.append(stars1)
            except IndexError :
                stars2.append(0)

            category2.append(str(category))


            count += 1

        # 페이지 번호를 넘기고, 다음 페이지 번호 클릭
        x += 1          
        try :
            driver.find_element_by_class_name("btn-page").find_element_by_link_text('%s' %x).click() # 다음 페이지번호 클릭
        except :
            pass

        time.sleep(2)  

        
          
# csv , xls 형태로 저장하기              
co_best_seller = pd.DataFrame()
co_best_seller['판매순위']=ranking2
co_best_seller['제품소개']=pd.Series(title2)
co_best_seller['제품판매가']=pd.Series(p_price2)
co_best_seller['원래 가격']=pd.Series(original2)
co_best_seller['할인율']=pd.Series(discount2)
co_best_seller['로켓배송여부']=pd.Series(rocket2)
co_best_seller['품절여부']=pd.Series(out2)
co_best_seller['상품평수']=pd.Series(sat_count2)
co_best_seller['상품평점']=pd.Series(stars2)
co_best_seller['분류']=pd.Series(category2)


# csv 형태로 저장하기
co_best_seller.to_csv(fc_name,encoding="utf-8-sig",index=True)

# 엑셀 형태로 저장하기
co_best_seller.to_excel(fx_name ,index=True)

e_time = time.time( )
t_time = e_time - s_time

count -= 1
print("\n")
print("=" *80)
print("1.요청된 총 %s 건의 리뷰 중에서 실제 크롤링 된 리뷰수는 %s 건입니다" %(cnt,count))
print("2.총 소요시간은 %s 초 입니다 " %round(t_time,1))
print("3.파일 저장 완료: txt 파일명 : %s " %ff_name)
print("4.파일 저장 완료: csv 파일명 : %s " %fc_name)
print("5.파일 저장 완료: xls 파일명 : %s " %fx_name)
print("=" *80)


print("=" *80)
print(" excel에 이미지를 저장 중입니다. 잠시만 기다려주세요. ")
print("=" *80)
print('\n')

# xls 파일에 제품 이미지 삽입하기

import win32com.client as win32   #pywin32 , pypiwin32 설치후 동작
import win32api  #파이썬 프롬프트를 관리자 권한으로 실행해야 에러없음

# 엑셀 작업 - 다소 수작업적인 부분 존재
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fx_name)
sheet = wb.ActiveSheet
sheet.Columns(3).ColumnWidth = 30   # 사진을 넣기 위해서 셀 크기를 키우는 작업 - 사이즈 선정은 일단은 노가다 식으로
row_cnt = cnt+1
sheet.Rows("2:%s" %row_cnt).RowHeight = 120   

ws = wb.Sheets("Sheet1")
col_name2=[]
file_name2=[]

for a in range(2,cnt+2) :
      col_name='C'+str(a)
      col_name2.append(col_name)

for b in range(1,cnt+1) :
      file_name=img_dir+'\\'+str(b)+'.jpg'
      file_name2.append(file_name)
      
for i in range(0,cnt) :
      rng = ws.Range(col_name2[i])
      image = ws.Shapes.AddPicture(file_name2[i], False, True, rng.Left, rng.Top, 130, 100)
      excel.Visible=True
      excel.ActiveWorkbook.Save()

# 웹드라어비 작업 이후 창 종료
driver.close( )

# 엑셀작업 완료 후 엑셀 파일 종료
excel.Quit()

# 처음에 데이터베이스에 저장한다고 한 경우, 바로 수집한 값 넣어주고 추후 분류를 위해 날짜값 추가(CRAWLED)
if db_save == 'y' or db_save == 'Y': 
    rows = [tuple(x) for x in co_best_seller.values]
    execute1 = 'INSERT INTO ' + db_name + ''' (RANKING, TITLE, P_PRICE, O_PRICE, DISCOUNT, DELIVERY, \
    STOCK, REVIEW, STARS, CATEGORY, CRAWLED) VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,SYSDATE)'''
    cur.executemany(execute1,rows)
    con.commit()

    cur.close()
    con.close()
    print("=" *80)
    print("6.DATABASE 저장 완료: TABLE 명 : %s " %(db_name))

# 처음에 데이터 저장을 보류한 경우 다시 처음부터 묻는 과정 거치기
elif  db_save == 'L' or db_save =='l' :
    db_save2 = input('크롤링된 데이터를 DB에 저장하시겠습니까? (Y/N):')

    # 보류 후 데이터베이스에 저장한다고 한 경우 
    if db_save2 =='y' or db_save2 == 'Y' :
        db_name = str(input('''데이터를 저장할 table명을 입력해주세요. (기본 db : CRAWLING_TABLE) :'''))
        if db_name == "" : 
            db_name = "CRAWLING_TABLE"
        create_table = '''CREATE TABLE  '''+ db_name + ''' (
                            RANKING NUMBER(4,0)  CONSTRAINT ranking_nn NOT NULL, 
                            TITLE VARCHAR2(256) CONSTRAINT title_nn NOT NULL,
                            P_PRICE NVARCHAR2(10) CONSTRAINT p_price_nn NOT NULL,
                            O_PRICE NVARCHAR2(10) CONSTRAINT o_price_nn NOT NULL,
                            DISCOUNT VARCHAR2(4) CONSTRAINT discount_nn NOT NULL,
                            DELIVERY VARCHAR2(8) CONSTRAINT delivery_nn NOT NULL,
                            STOCK VARCHAR2(8) DEFAULT '재고있음',
                            REVIEW NUMBER(6) CONSTRAINT review_nn NOT NULL,
                            STARS NUMBER(6,1) CONSTRAINT stars_nn NOT NULL,
                            CATEGORY VARCHAR2(20) CONSTRAINT category_nn NOT NULL,
                            CRAWLED DATE DEFAULT sysdate          
                            )''' 

        #오라클에 연결
        con = cx_Oracle.connect('scott/tiger@DESKTOP-0ULPS36/testdb')
        cur = con.cursor()

        #같은 이름의 테이블이 존재하는 경우
        try : cur.execute(create_table)
        except :
            option = input('''이미 테이블이 존재합니다. 
            데이터를 추가하거나(Y), 기존 테이블을 삭제하고 새로 만듭니다(N) : ''' )
            if option == 'N' or option == 'n':
                cur.execute('DROP TABLE '+ db_name)
                #간혹 발생하는 시드조정오류 자동화
                try : cur.execute(create_table)
                except :
                    print("1-4.오라클 연결 시드를 재조정합니다.")
                    cur.execute('SET oracle_sid=testdb')
                    con = cx_Oracle.connect('scott/tiger@DESKTOP-0ULPS36/testdb')
                    cur = con.cursor()
                    cur.execute(create_table)

            else : # 실수로 다른 것을 입력했을 때 기존 데이터가 삭제되는 경우를 대비해서 디폴트값을 붙여쓰기로 지정
                pass

        rows = [tuple(x) for x in co_best_seller.values]
        execute1 = 'INSERT INTO ' + db_name + ''' (RANKING, TITLE, P_PRICE, O_PRICE,DISCOUNT, DELIVERY, \
        STOCK, REVIEW, STARS, CATEGORY, CRAWLED) VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,SYSDATE)'''
        cur.executemany(execute1,rows)
        con.commit()

        cur.close()
        con.close()
        print("=" *80)
        print("6.DATABASE 저장 완료: TABLE 명 : %s " %(db_name))
        print("=" *80)
        print('\n')
    
    # 보류 후 데이터베이스에 저장하지 않겠다고 한 경우 - 이 경우 txt,csv,xls로만 저장됨
    else : pass

else : pass

print('\n')
print(" 작업이 완료되었습니다. ")
print(" 작업 폴더를 엽니다. ")
print("=" *80)

# 작업 완료 후 작업 폴더 열기
os.startfile(f_dir)


