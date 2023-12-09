# WHAT'S THIS?


**[KOR]**

PyCata Converter 는 Cataclysm:BN 의 모드 제작 시에 사용되는 JSON을 빠르게 만들기 위해 고안된 프로그램입니다.

엑셀의 1행을 Object의 Property 로 채우고, 그 이하로는 여러분의 모드 데이터로 채우십시오!

**[ENG(in DeepL)]**

PyCata Converter is a program designed to quickly create JSON for modding Cataclysm:BN.

Fill the first of the excel with the properties of your object, and the rest with your mod data!

# HOW TO USE

1. EXE 파일 실행

1. [파일 로드] 드랍다운 목록에서 엑셀 파일 선택
   - 이 목록은 EXE 파일과 같은 디렉토리에 있는 Excel 폴더에 저장되어 있는 엑셀파일들입니다.

1. [가져오기] 클릭

1. [Extract to JSON!] 클릭

1. JSON/{엑셀파일 이름}/ 경로에 저장된 JSON 파일 확인
   
단지 버튼을 연달아 누르는 것만으로 JSON 을 만들 수 있습니다!

# EXCEL RULES

그래요. 사실은 여기가 중요합니다.

맨 먼저 알려드리고 싶은 것은, 이 프로그램은 엑셀에서 데이터만을 받기 때문에 수식이나 셀의 크기 및 색깔 등은 영향을 주지 않는다는 것입니다.


1. 엑셀 시트의 이름은, JSON 으로 만들 데이터의 `TYPE` 으로 합니다. 이 Property 는 자동으로 가장 앞자리에 입력됩니다.
1. 해당 `TYPE`의 JSON 데이터가 필요로 하는 Property name 을 엑셀의 첫 행에 모두 입력합니다.
   - <u>__첫 번째 행은 </u>`id`<u>여야 합니다.__</u>
   - 만약 그 Property 의 값이 배열이라면, 그 끝에 `:list` 라고 한번만 추가로 입력하십시오.
   - 만약 그 Property(name1) 의 값이 또다른 Property(name2) 라면, 그것까지 모두 포함하여 입력해주십시오. (name1/name2)

        ```json
        [
            
            {
                "type": "GENERIC",
                "id": "copper_knife",
                "category": "tools",
                "weapon_category": [
                    "KNIVES"
                ],
                "name": {
                    "str": "copper knife",
                    "str_pl": "copper knives"
                },
                ...
            }
        ]
        ``` 
        과 같이 구현하기 위해서는 이하와 같이 엑셀을 구성해야 합니다.
        ![image][엑셀 예시1]
 
1. Property name 을 모두 입력한 뒤에는 각각의 Property 에 해당하는 Column 에 값을 입력합니다.
   - 객체는 id로 구분합니다.
   - 만약 하나의 Property 에 여러 개의 값을 넣으려면 그 아래 셀에 입력합니다. 그 값들은 같은 row 에는 id가 없어야 합니다.
   - `name:[[x1, y1], [x2, y2], [x3, y3]]` <br>
   위와 같이 2차원 배열을 값으로 받는 경우, 3row x 2column 행렬의 셀에 값을 구성하면 됩니다. 이 때 Property name 바로 아래 셀에 있는 값은 x1 가 됩니다.
   - `name:[[z1,[[x1, y1], [x2, y2], [x3, y3]]]]`<br> 
   위와 같은 형태로 구성된 배열도 있습니다. 네. 진짜로요. `mutant.json` 에서 `"vitamins_absorb_multi"` 라는 Property 가 그런 유형입니다. 이것은 예제 파일을 참고하여 주십시오. 너무 복잡해서 이 Property 는 단 하나의 값만 (예제의 경우 flesh에 대해서만) 받을 수 있도록 구현할 수 밖에 없었습니다.

1. 새로운 `TYPE`의 JSON 을 만든다면 새 엑셀 시트를 통해 상기의 규칙대로 입력해주세요.  


# AFTER...

이상이 현재까지 구현된 기능입니다.  

나중에 더 개선하여 볼 수 있도록 하겠습니다.



2023-12-10 02:15



[엑셀 예시1]: https://github.com/NappingOcean/PyCata_XL_to_JSON/assets/129575271/7e2b89bc-d9af-4132-a495-36c474c081df "이것은 GENERIC 시트임을 상기해주세요. Remember, this is in GENERIC sheet."