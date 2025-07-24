from BridgeStatistic import *

def get_bridge_config(n):
    """다리 번호에 따른 설정 반환"""
    if n == 1:
        return {
            'name': 'hangju',
            'stdevlist': ['SL_HJG_HGZ', 'SL_HJG_HGN', 'SL_HJG_HGE', 'SL_HJQ_HBZ', 'SL_HJQ_HBY', 'SL_HJH_HBZ', 'SL_HJK_HBZ', 'SL_HJK_HBY',
                          'SL_HJJ_HBZ', 'SL_HJF_HBX', 'SL_HJE_HBY', 'SL_HJE_HBX', 'SL_HJD_HBX', 'SL_HJC_HBY', 'SL_HJC_HBX', 'SL_HJP_HBY'],
            'wind': ['WG_1S_M', 'WG_2S_M', 'WG_3S_M']
        }
    elif n == 2:
        return {
            'name': 'gayang',
            'stdevlist': ['TA_1L','TA_1T','TA_1V'],
            'wind': None
        }
    elif n == 3:
        return {
            'name': 'worldcup',
            'stdevlist': ['AC_S2M_01_Y','AC_S2M_01_Z','EQK_WBG_X','EQK_WBG_Y','EQK_WBG_Z','EQK_WBQ_Y','EQK_WBQ_Z','EQK_WBH_Z','EQK_WBA_X','EQK_WBA_Y','EQK_WBC_X','EQK_WBC_Y','EQK_WBP_Y'],
            'wind': ['WS','WG_T1T_01_W','WG_T1T_01_S','WG_S1_01_W','WG_S1_01_S','WG_S1_02_S']
        }
    elif n == 4:
        return {
            'name': 'sungsan',
            'stdevlist': ['AC_S3_01'],
            'wind': None
        }
    elif n == 5:
        return {
            'name': 'yanghwa',
            'stdevlist': ['AC_01','AC_02'],
            'wind': None
        }
    elif n == 6:
        return {
            'name': 'seogang',
            'stdevlist': ['AC_01','AC_02'],
            'wind': ['WG_WS_M']
        }
    elif n == 7:
        return {
            'name': 'wonhyo',
            'stdevlist': [],
            'wind': None
        }
    elif n == 8:
        return {
            'name': 'hangang',
            'stdevlist': ['AC_P04P05_01','AC_P04P05_02','AC_P04P05_03','AC_P12P13_01','AC_P12P13_02','AC_P12P13_03'],
            'wind': None
        }
    elif n == 9:
        return {
            'name': 'sungsu',
            'stdevlist': ['AC_P8P9_01','AC_P8P9_02','AC_P8P9_03','AC_P8P9_04'],
            'wind': None
        }
    elif n == 10:
        return {
            'name': 'chongdam',
            'stdevlist': ['AC_P36P37_Z_01','AC_P37P38_Z_01'],
            'wind': ['WG_WS_M']
        }
    elif n == 11:
        return {
            'name': 'olympic',
            'stdevlist': ['SL_OPG_HGZ','SL_OPG_HGN','SL_OPG_HGE','SL_OPQ_HBZ','SL_OPQ_HBY','SL_OPH_HBZ','SL_OPK_HBZ','SL_OPK_HBY', 'SL_OPJ_HBZ', 'SL_OPA_HBZ', 'SL_OPA_HBX', 'SL_OPA_HBY', 'SL_OPI_HBY', 'SL_OPI_HBX', 'SL_OPP_HBY', 'SL_OPC_HBY', 'SL_OPC_HBX', 'SL_OPD_HBX', 'SL_OPB_HBZ', 'SL_OPB_HBX', 'SL_OPB_HBY'],
            'wind': ['WG_1S_M', 'WG_2S_M']
        }
    elif n == 12:
        return {
            'name': 'amsa',
            'stdevlist': ['EST_Z','EST_Y','EST_X','ACC1_Z','ACC2_Y','ACC3_Z','ACC4_Z','ACC5_Y','ACC6_X','ACC7_Z'],
            'wind': ['WS_1_M']
        }
    elif n == 13:
        return {
            'name': 'setgang',
            'stdevlist': ['AC_01_Y','AC_01_Z','AC_02_Y','AC_02_Z'],
            'wind': ['P_01_WS', 'U_01_WS_H']
        }
    elif n == 14:
        return {
            'name': 'cheonho',
            'stdevlist': ['JA3_X', 'JA3_Y', 'JA3_Z', 'JA4_X', 'JA4_Y', 'JA4_Z'],
            'wind': None,
            'min': ['DP_1_4_P10P11','DP_1_2_P10P11','DP_3_4_P10P11','DP_1_4_P11P12','DP_1_2_P11P12','DP_3_4_P11P12']
        }
    elif n == 15:
        return {
            'name': 'yeongdong',
            'stdevlist': ['AC_S1Q2_01_Z', 'AC_S4Q2_01_Z', 'AC_S7Q2_01_Z', 'AC_S10Q2_01_Z', 'AC_S13Q2_01_Z', 'AC_S16Q2_01_Z', 'AC_S17Q2_01_Z'],
            'wind': None
        }
    elif n == 16:
        return {
            'name': 'jamsu',
            'stdevlist': ['JA1_X','JA1_Y','JA1_Z','JA1_T','JA2_X','JA2_Y','JA2_Z'],
            'wind': None
        }
    else:
        raise ValueError("Invalid bridge number. Should be between 1 and 16.")

def get_quarter():
    """분기 선택"""
    print("\n분기를 선택하세요:")
    print("1. Q1 (1분기)")
    print("2. Q2 (2분기)")
    print("3. Q3 (3분기)")
    print("4. Q4 (4분기)")
    
    while True:
        try:
            q_choice = input("분기 번호를 입력하세요 (1-4): ").strip()
            q_choice = int(q_choice)
            if q_choice in [1, 2, 3, 4]:
                return f'Q{q_choice}'
            else:
                print("1~4 사이의 숫자를 입력해주세요.")
        except ValueError:
            print("올바른 숫자를 입력해주세요.")

def process_bridge(n, Q):
    """특정 다리 처리"""
    print(f"\n{'='*60}")
    print(f"다리 번호 {n} 처리 시작...")
    print(f"{'='*60}")
    
    config = get_bridge_config(n)
    br_name = config['name']
    stdevlist = config['stdevlist']
    wind = config.get('wind')
    min_list = config.get('min')
    
    print(f"다리 이름: {br_name}")
    print(f"분기: {Q}")
    print(f"센서 수: {len(stdevlist)}")
    if wind:
        print(f"풍속 센서: {len(wind)}개")
    if min_list:
        print(f"최소값 센서: {len(min_list)}개")
    
    try:
        # BridgeStatistic 객체 생성
        bridge = BridgeStatistic(br_name)
        
        # 데이터 전처리
        print("데이터 전처리 중...")
        if min_list:
            bridge.preprocess_average_data(stdevlist=stdevlist, wind=wind, min=min_list)

        else:
            bridge.preprocess_average_data(stdevlist=stdevlist, wind=wind)
        
        # 산점도 플롯
        print("산점도 생성 중...")
        bridge.plot_scatter(Q)
        
        # 분기 보고서
        print("분기 보고서 생성 중...")
        bridge.plot_quarterly_report(Q)
        
        # 연간 보고서
        print("연간 보고서 생성 중...")
        bridge.plot_yearly_report()
        
        print(f"다리 번호 {n} ({br_name}) 처리 완료!")
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        raise

def main():
    """메인 함수"""
    print("한강대교 통계 분석 프로그램")
    print("="*60)
    print("다리 번호: 1(행주), 2(가양), 3(월드컵), 4(성산), 5(양화)")
    print("          6(서강), 7(원효), 8(한강), 9(성수), 10(청담)")
    print("          11(올림픽), 12(암사), 13(샛강), 14(천호), 15(영동), 16(잠수)")
    print("="*60)
    
    # 분기 선택
    Q = get_quarter()
    
    while True:
        try:
            # 사용자로부터 다리 번호 입력받기
            n = input(f"\n처리할 다리 번호를 입력하세요 (1-16, 종료하려면 'q' 입력): ").strip()
            
            # 종료 조건
            if n.lower() == 'q':
                print("프로그램을 종료합니다.")
                break
            
            # 입력값 검증
            n = int(n)
            if n < 1 or n > 16:
                print("오류: 1~16 사이의 숫자를 입력해주세요.")
                continue
            
            # 다리 처리
            process_bridge(n, Q)
            
            # 다음 다리 처리 여부 확인
            while True:
                next_bridge = input("\n다음 다리를 처리하시겠습니까? (y/n): ").strip().lower()
                if next_bridge in ['y', 'n']:
                    break
                else:
                    print("y 또는 n을 입력해주세요.")
            
            if next_bridge == 'n':
                print("프로그램을 종료합니다.")
                break
                
        except ValueError:
            print("오류: 올바른 숫자를 입력해주세요.")
        except Exception as e:
            print(f"오류 발생: {str(e)}")
            print("다시 시도해주세요.")

if __name__ == "__main__":
    main() 
