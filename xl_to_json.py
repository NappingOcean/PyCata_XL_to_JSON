'''

구현 목표
1. xlsx 로부터 데이터를 가져와서 JSON 파일을 만든다.

2. JSON 파일로부터 형식만을 취하여 템플릿 xlsx 파일을 만든다.

3. 객체와 배열을 구분하여 데이터를 넣는다.
    - {}로서 표현되는 객체와 []로서 표현되는 배열.
    - 그럼 엑셀은 어떤 식으로 작성해둬야 하는가? 이는 나중에 마크다운 파일에 적어놓을 것.

4. JSON 파일이 들어있는 폴더를 연다.   

5. 이상의 사항들을 이벤트 루프 속에서 처리하는 창을 따로 띄우도록 한다.

'''

#import zone

import openpyxl as opyxl
import json as jo
import pandas as pd

class Data_Converter:
    # 생성자. 파일의 이름을 변수로서 받는다. 
    # 이 때 클래스 내부에서는 그 이름을 fn으로서 다룬다.
    def __init__(self, filename):
        self.fn = filename
    
    # xlsx 확장자로부터 데이터를 받는다.
    def load_xlsx(self):
        fn_path = "./" + self.fn
        load_wb = opyxl.load_workbook(fn_path, data_only=True)
        return load_wb
    
    # 1행에 있는 키 중에 복잡한 키를 분리하여 리스트화.
    def key_separator(self, keys):
        # 받은 keys를 / 별로 짤라서 리스트로 저장한다.
        key_serial = keys.split('/')
        return key_serial

    # 복합 딕셔너리에만 사용할 것! 이건 key_serial에만 대응한다.
    def dix_list_in_k(self, dix: dict, key, val):
        if isinstance(key, list):
            if not bool(dix.get(key, False)):
                dix[key] = []
            dix[key].append(val)

    #list가 아니라도 list로 만들어주마.
    def fc_append(self, is_it_list, not_list):
        if isinstance(is_it_list, list):
            return (lambda lst:lst.append(not_list))(is_it_list)
        else:
            return [is_it_list,not_list]


    # 해당 셀 아래로 데이터가 있는 셀이 얼마나 연달아 나오는가?
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
        vals_dict = {}
        for name in xl_data.sheetnames:
            xl_sheet = xl_data[name].values
            vals_list = []
            # 첫 줄 row 는 key가 있는 곳이다! 
            # 따라서 rows[0]은 배제한다.
            for row in list(xl_sheet.rows)[1:]:
                
                val_row = []
                val_id = xl_sheet.cell(row, 1).value

                for col in xl_sheet.columns:

                    val_key = xl_sheet.cell(1, col).value
                    cel = xl_sheet.cell(row, col).value

                    if cel: # 해당 셀에 값이 있을 경우.
                        if val_id: # id 존재 시.
                            if val_key: # key 존재 시.
                                if (row != xl_sheet.max_row) and not bool(xl_sheet.cell( row + 1, 1 ).value): # 여기가 끝 아님 and 이 아래로 id 없는 값 존재 시.
                                    no_id_num = self.col_ser_check(xl_sheet, row, col)

                                    v_list_col = [ xl_sheet.cell(list_row, col).value for list_row in xl_sheet.rows[row : row + no_id_num] ]
                                    # id 없는 모든 값들을 리스트로 흡수함.
                                    val_row.append(v_list_col)
                                else: # 다음 값 id 있다
                                    val_row.append(cel)
                                
                                # 리셋
                                v_list = []
                                bool_v_2d_list = False
                        
                            else: # key 부재 시.
                                if pre_cel in val_row: 
                                    val_row.remove(pre_cel) 
                                    v_list.append(pre_cel)
                                #pre_cel 이 있든 없든 아래는 진행.
                                v_list.append(cel)
                                
                                #이 아래에 id 없는 값이 있을 시.
                                if not xl_sheet.cell(row + 1, 1).value:
                                    # 이 때 cel은 이미 key가 없는 곳의 셀이다
                                    # 즉 앞의 pre_cel과 v_list를 이미 이루고 있다는 뜻.

                                    no_id_num = self.col_ser_check(xl_sheet, row, col)
                                    
                                    # 현재 셀에서 len(v_list) 만큼 왼쪽까지의 셀 값을 list로 만든다
                                    # 상기의 list 생성 과정을 현재 셀에서 no_id_num 갯수만큼 아래로 내려가며 반복한다
                                    v_2d_list_col = [[xl_sheet.cell(list_row, list_col).value for list_col in xl_sheet.columns[col+1 -len(v_list): col+1]] for list_row in xl_sheet.columns[row: row + no_id_num]]
                                    
                                    bool_v_2d_list = True
                                
                                if bool(xl_sheet.cell(1, col+1).value) or (col == xl_sheet.max_column):
                                    if bool_v_2d_list:
                                        # 2d list가 되었으면 그걸로 넣는다
                                        val_row.append(v_2d_list_col)
                                    else:
                                        # v_list 를 val 에 추가
                                        val_row.append(v_list)

                        # 값이 있는 셀은 이전 셀로 만든다.
                        pre_cel = cel
                vals_list.append(val_row)
            vals_dict[name] = vals_list
        # vals_dict에 저장되는 값은 다음과 같다
        # vals_dict --> {"name":
        #                   [
        #                       [v11, v12, v13, ...]
        #                       [v21, v22, v23, ...] 
        #                       ...
        #                   ]
        #                }
        return vals_dict




    # 본격적으로 딕셔너리 구축            
    def dix_builder(self, keys: list, vals: list):
        
        #k_map = jagged list
        k_map = list(map(self.key_separator, keys))
        
        k_last_list = list(map(lambda ks:ks[-1], k_map))
        dix_last = dict(zip(k_last_list, vals))
        
        bld_dix = {}

        for k_serial in k_map:
            tmp_idx = 0
            k_idx = 0
            for k in reversed(k_serial):
                k_idx-=1
                k_cut = k.removesuffix(':list')
                if k == k_serial[-1]:
                    tmp_dix = {}
                    if k.endswith(':list') and not isinstance(dix_last[k], list):
                        tmp_dix = {k_cut:[dix_last[k]]}
                    else:
                        tmp_dix = {k_cut:dix_last[k]}
                else: # k가 맨 끝에 있는 k 가 아니다.
                    if not k.endswith(':list'):
                        tmp_dix = {k:tmp_dix}
                    elif not k in [k_ser[k_idx] for k_ser in k_map[:k_map.index(k_serial)] if len(k_ser)+k_idx >= 0 and k_ser[:k_idx] == k_serial[:k_idx]]:
                        tmp_dix = {k_cut:[tmp_dix]}
                    
                    else: # k가 list로 끝남 & 앞으로의 k들이 일렬로 전부 들어있음
                        tmp_idx += 1

            tmp_adr = bld_dix

            if tmp_idx > 0: # tmp_dix 의 디렉토리와 겹치는 값이 하나 이상!
                for k2 in k_serial[:tmp_idx]:
                    k2_cut = k2.removesuffix(':list')
                    # 기존 값의 어디에 붙여야 하는지 탐색.
                    if not isinstance(tmp_adr,list):
                        tmp_adr = tmp_adr[k2_cut]
                    else:
                        for dix in tmp_adr:
                            if k2_cut in dix:
                                tmp_adr = dix

                    if k2 == k_serial[tmp_idx-1]:
                        if isinstance(tmp_adr,list):
                            tmp_dix = {k2_cut:tmp_adr.append(tmp_dix)}
                        else:
                            tmp_dix = {k2_cut:tmp_adr[k2_cut].append(tmp_dix)}
            else:
                tmp_adr.update(tmp_dix)

        return bld_dix

    def json_to_templete(self):
            pass
    
class File_Loader:
    def __init__(self, filename):
        self.fn = filename

    def read_file(self):
        pdxl = pd.ExcelFile("./"+self.fn)
        return pdxl
    
    def keys_from_file(self, pdxl:pd.ExcelFile) -> dict: 
        
        keys_dict = {}

        for s_name in pdxl.sheet_names:
            df = pdxl.parse(sheet_name=s_name)
            list_from_df = list(df.columns.values)
            for idx in range(len(list_from_df)):
                if list_from_df[idx].startswith('Unnamed: '):
                    list_from_df[idx] = None
                    
            keys_dict[s_name] = list_from_df

        return keys_dict

    def raw_vals_from_file(self, pdxl:pd.ExcelFile) -> dict: 
        
        raw_vals_dict = {}

        for s_name in pdxl.sheet_names:
            df = pdxl.parse(sheet_name=s_name)
            list_from_df = df.values.tolist()
            for lst in list_from_df:
                for idx in range(len(lst)):
                    if pd.isna(lst[idx]):
                        lst[idx] = None
            raw_vals_dict[s_name] = list_from_df

        return raw_vals_dict
    
    def vals_pro(self, pdxl:pd.ExcelFile) -> dict:

        vals_dict = {}

        keys = self.keys_from_file(pdxl)
        r_vals = self.raw_vals_from_file(pdxl)

        for s_name in pdxl.sheet_names:
            key_list = keys[s_name]
            val_list_list = r_vals[s_name]
            
            val_dict_in_name = {}
            
            for row in range(len(val_list_list)):
                val_list = val_list_list[row]
                val_id = val_list[0]
                val_in_row = []
                for col in range(len(key_list)):
                    key_now = key_list[col]
                    val_now = val_list[col]
                    if val_id:
                        exist_id = val_id
                        if key_now:
                            val_in_row.append(val_now)
                        else:
                            tmp_val = val_in_row.pop()

                            if isinstance(tmp_val, list):
                                tmp_val.append(val_now)
                            else:
                                tmp_val = [tmp_val, val_now]

                            val_in_row.append(tmp_val)
                    
                    else:   # id가 없는 value.
                        if val_now: # 그 자리 값 있음
                            if key_now: # 그 자리 키 있음
                                upper_val = val_dict_in_name[val_dict_in_name.keys()[-1]][col]
                                # 상위값에 접근이 가능할려나?

                            else: # 그 자리 키 없음
                                
                                pass
                # end of for-loop:col
                val_dict_in_name.update({exist_id:val_in_row})
            
            vals_dict[s_name] = [val_dict_in_name[key] for key in val_dict_in_name.keys()]
    

if __name__ == "__main__":
    
    # Test build
    main_file = File_Loader("test_XL.xlsx")
    
    main_keys = [
        "id",
        "name/str",
        "skills:list/skill:list/combat",
        "skills:list/skill:list/craft",
        "skills:list/level/dice:list",
        "skills:list/level/add"
    ]
    main_vals = [
        "id_sample1",
        "샘플 이름",
        "slash",
        "ALL",
        [2,6],
        4
    ] 

ori_xl = main_file.read_file()

main_key_dict = main_file.keys_from_file(ori_xl)
main_val_dict = main_file.raw_vals_from_file(ori_xl)

print(main_key_dict)
print(main_val_dict)


    