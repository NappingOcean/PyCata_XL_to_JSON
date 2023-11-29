'''

구현 목표
1. xlsx 로부터 데이터를 가져와서 JSON 파일을 만든다.

2. JSON 파일로부터 형식만을 취하여 템플릿 xlsx 파일을 만든다.

3. 객체와 배열을 구분하여 데이터를 넣는다.
    - {}로서 표현되는 객체와 []로서 표현되는 배열.
    - 그럼 엑셀은 어떤 식으로 작성해둬야 하는가? 이는 나중에 마크다운 파일에 적어놓을 것.
    - 배열을 저장할 필요가 있는 경우 접두어에 Arr를 붙인다던지?
    - 아니면 애초에 표시 자체를 어떤 괄호를 썼는지 직접 쓰는 게 좋을지도 모르겠다.
    - 그런데 그게 엑셀에서 직관적으로 나오느냐가 문제.
    - 엑셀 Sheet의 구분은 "type" 으로 한다.
    - 1행에는 key 값을 나열. 2행부터 key 값에 따라 value를 넣는다.
    - 문제점: value로 객체 또는 배열이 필요한 경우.
        * 요구하는 바인딩에 따라 /obj, /arr를 접미로 붙인다.
        * 다시 그 뒤에 key 값을 넣는다.
    - 즉 1행의 key를 받을 때 /obj, /arr 사항을 검사하고 추가적인 조치를 취한다.
    - 응? /arr 를 받을 때는 key 값이 필요없다?
    - 그러면 키가 같은 값에 대해 array로 출력하도록 해야 하나.

4. JSON 파일이 들어있는 폴더를 연다.   

5. 이상의 사항들을 이벤트 루프 속에서 처리하는 창을 따로 띄우도록 한다.

'''

#import zone

import openpyxl as opyxl
import json as jo

class File_manager:
    # 생성자. 파일의 이름을 변수로서 받는다. 
    # 이 때 클래스 내부에서는 그 이름을 fn으로서 다룬다.
    def __init__(self, filename):
        self.fn = filename
    
    # xlsx 확장자로부터 데이터를 받는다.
    def load_xlsx(self):
        fn_path = "./xlsx/" + self.fn
        load_wb = opyxl.load_workbook(fn_path, data_only=True)
        return load_wb
    
    # 1행에 있는 키 중에 복잡한 키를 분리하여 리스트화.
    def key_separator(self, keys):

        # 받은 keys를 / 별로 짤라서 리스트로 저장한다.
        key_serial = keys.split('/')
        
        # 각 키마다 있을 수 있는 ':list' 를 제거한다.
        # 이 함수는 나중에 key에서 :list를 이용해야 하면 없앤다.
        # for key in key_serial:
        #     key = key.removesuffix(':list')
        
        return key_serial

    # 읽은 데이터를 json으로 가공한다.
    # 굳이 한 번에 할 필요 없지.
    def xlsx_to_json(self):
        xl_data = self.load_xlsx()
        
        # 시트의 1행을 키로 받고, 열별로 값을 받은 다음, 
        # 각 행마다의 딕셔너리 객체를 만들어서 딕셔너리 list로 만든다.
        xl_dix_list = []
        xl_data_inpy = {}
        #시트 네임마다 차례대로 살핀다.
        for name in xl_data.sheetnames:
            #해당하는 시트의 정보를 xl_sheet에 저장.
            xl_sheet = xl_data[name]
            #1행의 키들을 모두 가져온다.
            xl_keys = xl_sheet[1]

            for column in range(2, xl_sheet.max_column + 1):
                for row in range(1, xl_sheet.max_row + 1):
                    
                    # 시트 1행 중에 비어있지 않은 열만 k로 저장함.
                    # 만약 비어있다면 이전의 k 값을 그대로 쓰겠지.
                    if xl_keys[row] != None:
                        k = xl_keys[row]
                       
                        # list를 value로서 받는 경우.
                        v_list_bl:bool = False
                        k_idx_list = k.rfind(':list')
                        if k_idx_list >= 0:
                            v_list_bl = True

                       # obj 를 value로서 받는 경우. 
                        v_obj_bl:bool = False
                        k_idx_obj = k.rfind('/')
                        if k_idx_obj >= 0:
                            v_obj_bl = True
                            k_list = self.key_separator(k)

                    # 행의 1번열(id)에 값이 존재할 경우.
                    if xl_sheet.cell(1, column).value != None:
                        # 바로 밑 셀과 같은 열에 id가 없거나, 엑셀에서 통짜로 가져온 k 값의 끝에 :list가 포함되어 있을 경우.
                        if (xl_sheet.cell(1, column + 1).value == None) or k.endswith(':list'):
                            v=[xl_sheet.cell(1,column).value]
                            
                            idx = 1
                            v_id_next = xl_sheet.cell(1, column+idx).value
                            # 다음 행의 id가 0이 아니거나 최대 좌표에 도달할 때까지 다음 행의 데이터를 v 배열에 넣는 것을 반복한다.
                            while v_id_next == None:
                                if column+idx == xl_sheet.max_column:
                                    break
                                v_next = xl_sheet.cell(row, column + idx).value
                                v.append(v_next)
                                idx = idx + 1
                        #만약 단일값이고 k 끝에 :list가 없다면.
                        else:
                            v=xl_sheet.cell(row, column).value

                        # k에 /가 포함되어 있을 경우.
                        if v_obj_bl:
                            #임시 루프를 만들고 쫙 넣는다.
                            tmp_dix = {}
                            for k_num in range(len(k_list), 1, -1):
                                tmp_dix.clear()
                                key = k_list[k_num]
                                if k_num == len(k_list):
                                    tmp_dix[key] = v
                                elif k_num < len(k_list):
                                    pre_key = k_list[k_num + 1]
                                    tmp_dix[key] = tmp_dix
                            # 루프 종료 후 tmp_dix를 통으로 먹는다.
                            xl_data_inpy[k_list[0]] = tmp_dix
                        #TODO: k에 / 말고도 :list가 포함되어 있는 경우도 많이 있을 것이다. 그건 다 어떻게 처리할거냐? 하나하나 append를 먹일 수 있나?

                        else:
                            xl_data_inpy[k] = v

                    # id 값이 없다면?
                    else:
                        # 해당 셀이 비어 있지 않다면.
                        if xl_sheet.cell(row, column).value != None:
                             
                            # v가 리스트가 아닐 경우, v를 리스트로 먼저 만들어야 한다.
                            if not isinstance(xl_data_inpy[k],list):
                                xl_data_inpy[k] = [v]

                            ele_v = xl_sheet.cell(row,column).value
                            #이전에 저장된 v에 대해 리스트를 축적시킨다.
                            xl_data_inpy[k].append(ele_v)
                        # id도 비어있고 그 자리의 셀도 비어있으면 그냥 지나간다.
                # row 루프 종료. 
                # xl_data_inpy 딕셔너리에는 
                # 1줄의 column에 있는 데이터가 모두 저장됐다!
                # 그 딕셔너리를 리스트에 추가한다.
                xl_dix_list.append(xl_data_inpy)

    # 복합 딕셔너리에만 사용할 것! 이건 key_serial에만 대응한다.
    def dix_list_in_k(self, dix: dict, key, val):
        if isinstance(key, list):
            if not bool(dix.get(key, False)):
                dix[key] = []
            dix[key].append(val)


    def col_ser_check(self, ws, row, col):
        ser_num = 1
        for check_row in ws.rows[row + 1:]:
            if bool(ws.cell(check_row, col)) and not bool(ws.cell(check_row, 1)):
                ser_num += 1
            else:
                return ser_num

    # vals 는 2차원 리스트로 작성된다.
    def vals_from_xl(self):
        xl_data = self.load_xlsx()
        for name in xl_data.sheetnames:
            xl_sheet = xl_data[name]
            vals = []
            # 첫 줄 row 는 key가 있는 곳이다! 
            # 따라서 rows[0]은 배제한다.
            for row in xl_sheet.rows[1:]:
                
                val_row = []
                val_id = xl_sheet.cell(row, 1).value

                for col in xl_sheet.columns:

                    val_key = xl_sheet.cell(1, col).value
                    cel = xl_sheet.cell(row, col).value

                    if cel: # 해당 셀에 값이 있을 경우.
                        if val_key: # key 존재 시.
                            if val_id: # id 존재 시.
                                if v_list:
                                    val_row.append(v_list)
                                # 이 아래로 id 없는 값 존재 시.
                                if not xl_sheet.cell( row + 1, 1 ).value:
                                    no_id_num = self.col_ser_check(xl_sheet, row, col)

                                    v_list_col = [ xl_sheet.cell(list_row, col).value for list_row in xl_sheet.rows[row : row + no_id_num] ]
                                    # id 없는 모든 값들을 리스트로 흡수함.
                                    val_row.append(v_list_col)
                                else:
                                    val_row.append(cel)
                                    
                                v_list = []
                        else: # key 부재 시.
                            if val_id: # id 존재 시.
                                if pre_cel in val_row: 
                                    val_row.remove(pre_cel) 
                                    v_list.append(pre_cel)
                                    #TODO: id 없는 애들 어케 흡수하냐.
                                v_list.append(cel)
                        
                        # 값이 있는 셀은 이전 셀로 만든다.
                        pre_cel = cel
                vals.append(val_row)

    def keys_from_xl(self):
        xl_data = self.load_xlsx()


    # 본격적으로 딕셔너리 구축            
    def dix_builder(self, keys: [], vals: []):
        
        k_map = list(map(self.key_separator, keys))
        
        # 키맵 가장 끝자락의 k로 이루어진 리스트를 만든다.
        # 빈 리스트 하나 만들고 시작.
        k_last = []
        for k_serial in k_map:
            k_last.append(k_serial.pop())
        
        dix_last = dict(zip(k_last, vals))

        # 이제부턴 합칠 차례다.
        for k_serial in k_map:
            if k_serial:
                #TODO
                pass


    def json_to_templete(self):
        pass