import streamlit as st
import calendar
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# --- 頁面設定 ---
st.set_page_config(page_title="澳森托嬰中心 智慧教案與適性月計畫系統", layout="wide", initial_sidebar_state="collapsed")

# --- Word 樣式輔助工具 ---
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=70, bottom=70, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="1F497D", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

# --- 完整適性指標活動資料庫 (標準統一指標代碼與活動標題，無頁數雜訊) ---
DATABASE = {
    "0-12個月": {
        "顏色": {
            "身體動作": "1. 俯臥抬頭練習與抓握不同材質彩色布球【抬頭小娃/握緊拳頭】\n2. 趴姿支撐伸手摸眼前彩色懸吊物【抓抓樂】\n(指標 I-1-1 / II-1-4)",
            "社會情緒": "1. 與老師面對面互動發出社會性微笑【笑一個】\n2. 對鏡中彩色影像表現愉悅情緒【誰在鏡子裡】\n(指標 I-2-1 / II-2-2)",
            "語言溝通": "1. 聆聽繪本時注視老師口型變化【看誰在唱歌】\n2. 呼喚名字時會轉頭尋找聲源【聲音在哪裡】\n(指標 I-3-2 / I-3-1)",
            "認知探索": "1. 眼神追隨顏色鮮明之大色卡緩慢移動【追視物品】\n2. 觸摸柔軟絲巾與毛氈布感受觸覺刺激【搔癢遊戲】\n(指標 I-4-1 / I-4-3)",
            "生活自理": "1. 喝奶時能在協助下伸手扶握奶瓶【一起拿奶瓶】\n2. 練習吞嚥米糊/果汁等泥狀副食品【用湯匙喝果汁】\n(指標 II-5-1 / II-5-2)",
            "文化藝術": "1. 「彩色足印蓋印」：老師協助在腳底塗無毒顏料蓋出印記【色彩感知】\n2. 聆聽輕柔兒歌《小星星》，隨音樂給予肢體撫拍【寶寶會唱歌】\n(指標 II-4-5 / II-3-3)"
        },
        "中秋": {
            "身體動作": "1. 仰臥時練習雙腳踢懸吊黃色大氣球【踢踢樂】\n2. 撐起身體向前爬行探索月亮抱枕【爬爬樂】\n(指標 II-1-6 / III-1-2)",
            "社會情緒": "1. 抱在懷中進行親密拍撫與目光交流【建立感情】\n2. 和老師玩黃色毛巾「躲貓貓」遊戲【躲貓貓】\n(指標 III-2-5 / III-2-2)",
            "語言溝通": "1. 聆聽繪本，聆聽溫柔語調「月亮圓圓」【寶寶會唱歌】\n2. 咿呀發聲回應老師話語【發聲遊戲】\n(指標 II-3-3 / I-3-3)",
            "認知探索": "1. 觸摸光滑柚子皮與毛絨月亮抱枕【感受觸覺】\n2. 尋找被小毛巾稍微遮蓋的黃色小球【不見了】\n(指標 I-4-3 / III-4-2)",
            "生活自理": "1. 能自己拿住奶瓶進食【幫忙拿奶瓶】\n2. 嘗試用小湯匙吞嚥細緻水果泥【麥片時間】\n(指標 III-5-1 / III-5-2)",
            "文化藝術": "1. 「金黃圓月掌印畫」：小手掌沾黃色顏料拓印圓形【手掌蓋印】\n2. 聆聽搖籃曲《月光光》，享受舒緩音律【音律感受】\n(指標 II-4-5 / I-3-2)"
        },
        "長大": {
            "身體動作": "1. 練習由仰臥翻身成俯趴姿勢【翻身運動】\n2. 在支撐下練習雙手向前扶坐【坐穩囉】\n(指標 II-1-2 / II-1-3)",
            "社會情緒": "1. 對鏡中的自己揮手、展露笑容【鏡中人】\n2. 主動伸手向熟悉老師討抱【建立感情】\n(指標 I-2-1 / III-2-5)",
            "語言溝通": "1. 聆聽繪本，模仿發出單音「咕、叭」【寶寶學說話】\n2. 能注視說話者並轉頭回應【看誰在說話】\n(指標 II-3-2 / II-3-4)",
            "認知探索": "1. 用雙手搖晃小沙鈴聽取清脆聲響【搖一搖】\n2. 觸摸溫熱毛巾感受溫度變化【觸覺體驗】\n(指標 II-4-1 / I-4-3)",
            "生活自理": "1. 練習自行拿小米餅放入口中吃【我會自己拿】\n2. 協助拉下頭頂小帽子【脫帽子】\n(指標 III-5-3 / III-5-4)",
            "文化藝術": "1. 「破殼而出指印畫」：大拇指沾顏料點畫裝飾【指印拓印】\n2. 兒歌律動《小雞嗶嗶》肢體伸展操【舒活時間】\n(指標 II-4-5 / I-1-5)"
        },
        "父親": {
            "身體動作": "1. 仰臥時抓握手搖鈴由一手換至另一手【換一隻手】\n2. 練習在軟墊上藉由拉扶坐穩【能自己坐】\n(指標 III-1-5 / III-1-1)",
            "社會情緒": "1. 看著爸爸的照片能開心微笑並伸手觸摸【笑一個】\n2. 熟悉照顧者抱抱時給予依附回饋【建立感情】\n(指標 I-2-1 / III-2-5)",
            "語言溝通": "1. 聆聽繪本，模仿大人發出「ㄅㄚ、ㄅㄚ」疊音【跟我一起說】\n2. 能辨別溫柔安撫語氣【心情表情】\n(指標 III-3-1 / III-2-4)",
            "認知探索": "1. 觸摸爸爸的柔軟領帶與棉質衣物【觸覺刺激】\n2. 聽見叫喚自己名字時會轉頭注視【球兒滾來滾去】\n(指標 I-4-3 / III-4-4)",
            "生活自理": "1. 餵奶時能安穩吸吮並練習雙手捧奶瓶【一起拿奶瓶】\n2. 吃完副食品後配合毛巾擦拭嘴角【清潔衛生】\n(指標 II-5-1 / I-5-1)",
            "文化藝術": "1. 「給爸爸的愛心手印」：彩色掌印蓋在卡片上【手掌拓印】\n2. 聆聽溫馨兒歌《我的好爸爸》，進行嬰兒舒活按摩【舒活時間】\n(指標 I-1-5 / II-3-3)"
        },
        "default": {
            "身體動作": "1. 趴姿下用雙臂撐起上半身抬頭看前方【我會趴】\n2. 雙手互拍做出拍手動作【拍拍手】\n(指標 II-1-1 / III-1-7)",
            "社會情緒": "1. 在清涼水墊/軟墊上拍拍，展現興奮愉悅神情【飛天毛毯】\n2. 看到同伴時主動伸出小手觸碰【和我一起玩】\n(指標 II-2-3 / II-2-5)",
            "語言溝通": "1. 聆聽繪本語句，聆聽情境擬聲詞【看誰在唱歌】\n2. 牙牙學語發出不同高低語調【動物大集合】\n(指標 I-3-2 / III-3-2)",
            "認知探索": "1. 觸摸涼水袋與溫暖毛巾感受溫差【觸覺探索】\n2. 凝視色彩明亮之教玩具【瞧一瞧】\n(指標 I-4-3 / II-4-5)",
            "生活自理": "1. 練習用雙手捧住學習杯喝水【幫忙拿奶瓶】\n2. 沐浴/清潔時能放鬆四肢配合擦澡【舒活時間】\n(指標 III-5-1 / I-1-5)",
            "文化藝術": "1. 「手指膏拍拍畫」：手指沾水洗顏料在畫板塗抹【色彩感知】\n2. 配合輕快童謠做手腳輕拍律動【寶寶會唱歌】\n(指標 II-4-5 / II-3-3)"
        }
    },
    "7-12個月": {
        "顏色": {
            "身體動作": "1. 爬行穿過軟墊隧道，探索彩色大布球【爬爬樂】\n2. 練習用前兩指鉗握彩色小積木【換一隻手】\n(指標 III-1-2 / III-1-6)",
            "社會情緒": "1. 拿起彩色手偶向同儕揮手打招呼【打招呼】\n2. 與老師玩彩色絲巾「躲貓貓」遊戲【躲貓貓】\n(指標 III-2-1 / III-2-2)",
            "語言溝通": "1. 聆聽繪本，指認「熊熊、小鳥」等動物【表示意見】\n2. 模仿發出「ㄨㄤ、ㄨㄤ」或「ㄇㄚ、ㄇㄚ」簡單音節【跟我一起說】\n(指標 III-3-3 / III-3-1)",
            "認知探索": "1. 將彩色大積木放入相同顏色的大收納籃中【我會玩】\n2. 尋找被藏在彩布下的玩具【不見了】\n(指標 III-4-6 / III-4-2)",
            "生活自理": "1. 點心時間能自己用手拿米餅進食【我會自己拿】\n2. 喝完奶後能自己拿住奶瓶【幫忙拿奶瓶】\n(指標 III-5-3 / III-5-1)",
            "文化藝術": "1. 「手掌拍拍色彩畫」：手掌沾藍色與黃色水彩在大白紙上拍印【拍拍手】\n2. 兒歌律動《顏色歌》，隨音樂揮動雙手拍拍手【動物大集合】\n(指標 III-1-7 / III-3-2)"
        },
        "中秋": {
            "身體動作": "1. 扶著軟墊箱站立並嘗試彎腰撿拾黃色月亮球【站起來囉】\n2. 雙手合力推滾大球向前【球兒滾來滾去】\n(指標 III-1-3 / III-4-5)",
            "社會情緒": "1. 與同儕坐在一起玩中秋玩具，能互相注視【建立感情】\n2. 看到老師拿出柚子時露出期待神情【心情表情】\n(指標 III-2-5 / III-2-3)",
            "語言溝通": "1. 聆聽繪本，引導指認「圓圓的月亮」【我聽懂了】\n2. 用手指物表示想要吃月餅或玩玩具【表示意見】\n(指標 III-3-5 / III-3-3)",
            "認知探索": "1. 觸摸粗糙的柚子皮與光滑的月餅模型【感官知覺】\n2. 掀開毛巾找出被藏起來的黃色小球【不見了】\n(指標 III-4-6 / III-4-2)",
            "生活自理": "1. 戴上柚子帽後能自己練習拉下帽子【脫帽子】\n2. 嘗試用小湯匙吞嚥細碎副食品【麥片時間】\n(指標 III-5-4 / III-5-2)",
            "文化藝術": "1. 「黃金大月亮拓印」：利用圓形海綿塊沾黃色顏料拓印【模具拓印】\n2. 聆聽兒歌《月亮圓圓》，雙手跟著音樂拍拍手【拍拍手】\n(指標 III-1-7 / III-3-2)"
        },
        "長大": {
            "身體動作": "1. 爬行跨越低矮軟墊，練習雙手雙腳支撐協調【爬爬樂】\n2. 練習一手拿玩具，另一手敲擊發聲【我會玩】\n(指標 III-1-2 / III-4-6)",
            "社會情緒": "1. 聽到自己名字能主動轉頭並揮手【打招呼】\n2. 在鏡中看見自己能用手拍鏡子表達喜悅【鏡中人】\n(指標 III-2-1 / I-2-1)",
            "語言溝通": "1. 聆聽繪本，指認熟悉動物並嘗試仿說【跟我一起說】\n2. 聽懂簡單生活指令「抱抱、過來」【我聽懂了】\n(指標 III-3-1 / III-3-5)",
            "認知探索": "1. 操作按壓玩具，觀察燈光與聲響變化【我會玩】\n2. 看到水杯或圍兜能預期要吃點心【等一下做什麼】\n(指標 III-4-6 / III-4-3)",
            "生活自理": "1. 練習自己拿小塊水果丁送入口中【我會自己拿】\n2. 嘗試用雙手捧學習杯喝水【幫忙拿奶瓶】\n(指標 III-5-3 / III-5-1)",
            "文化藝術": "1. 「小手蓋印拓印畫」：手掌沾安全顏料蓋在白紙上【手掌蓋印】\n2. 聆聽兒歌隨旋律晃動身體與拍手【拍拍手】\n(指標 III-1-7 / II-3-3)"
        },
        "default": {
            "身體動作": "1. 撐起身體向前爬行追逐滾動的小皮球【爬爬樂】\n2. 練習在協助下扶著矮欄杆站立數秒【站起來囉】\n(指標 III-1-2 / III-1-4)",
            "社會情緒": "1. 聽到叫喚自己名字時主動轉頭並伸出手要抱【建立感情】\n2. 早上入園向家長揮手說拜拜【打招呼】\n(指標 III-2-5 / III-2-1)",
            "語言溝通": "1. 聆聽繪本，理解簡單語彙「喝水水、抱抱」【我聽懂了】\n2. 與老師進行輪流咿呀對話【聊聊天】\n(指標 III-3-5 / III-3-4)",
            "認知探索": "1. 操作按壓發聲玩具，觀察因果關係【我會玩】\n2. 看到奶瓶出現能預期喝奶時間到了【等一下做什麼】\n(指標 III-4-6 / III-4-3)",
            "生活自理": "1. 練習用雙手捧住雙耳學習杯喝水【幫忙拿奶瓶】\n2. 表達進食意願，能主動張口接受湯匙餵食【我要吃東西】\n(指標 III-5-1 / III-5-5)",
            "文化藝術": "1. 「繽紛手指膏塗塗樂」：手指沾可水洗顏料在光滑畫板上塗抹【色彩感知】\n2. 聆聽輕快童謠，隨節奏擺動身體與拍手【拍拍手】\n(指標 III-1-7 / III-3-2)"
        }
    },
    "12-24個月": {
        "顏色": {
            "身體動作": "1. 體能課：依顏色指令跨步走入「藍色/黃色呼啦圈步道」【大肌肉跨步】\n2. 小肌肉/認知：將彩色毛球用小夾子舀入同色杯中【精細操作】\n(指標 V-1-2 / V-1-5)",
            "社會情緒": "1. 與同伴一起進行「彩色絲巾捉迷藏」，體驗同儕互動樂趣【團體遊戲】\n2. 戴上手偶與同伴互相打招呼【打招呼】\n(指標 V-2-2 / V-2-6)",
            "語言溝通": "1. 聆聽繪本，複述重複句型與詞彙\n2. 能說出紅、黃、藍、綠、棕等顏色名稱【小博士】\n(指標 V-3-1 / V-3-2)",
            "認知探索": "1. 觸摸體驗不同材質彩色布料（滑滑絲巾、毛毛布、粗糙毛氈）【感覺一下】\n2. 彩色積木依顏色分類歸位【分類遊戲】\n(指標 V-4-1 / V-4-3)",
            "生活自理": "1. 點心後練習用小毛巾擦拭嘴巴與雙手【清潔寶寶】\n2. 遊戲後主動將彩色玩具送回指定顏色籃【自己收玩具】\n(指標 V-5-6 / V-5-7)",
            "文化藝術": "1. 「主題蓋印畫」：海綿刷沾顏料在圖畫紙上蓋印出角色【好玩的刷畫】\n2. 律動兒歌《顏色歌》，隨節奏揮舞彩帶【載歌載舞】\n(指標 V-4-6 / V-3-5)"
        },
        "中秋": {
            "身體動作": "1. 體能課：抱大球繞過圓形呼啦圈（月亮步道），練習推球與踢球【踢球樂】\n2. 手掌揉捏黃色黏土、用模型壓印月餅【小肌肉按壓】\n(指標 V-1-4 / V-1-5)",
            "社會情緒": "1. 在「中秋烤肉/吃月餅家家酒」中與同儕分享食物道具【同儕分享】\n2. 辨識故事中想吃月餅的調皮笑臉【喜怒哀樂】\n(指標 V-2-8 / V-2-3)",
            "語言溝通": "1. 聆聽繪本，學習詞彙「月亮、月餅、圓圓的」\n2. 練習說出「我要吃月餅」、「月亮好大」【以語言表達】\n(指標 V-3-1 / V-3-2)",
            "認知探索": "1. 觸摸體驗「柚子皮（粗糙凹凸）」與「柚子果肉（多汁軟軟）」【感覺一下】\n2. 觀察月亮形狀大小變化【分類遊戲】\n(指標 V-4-1 / V-4-3)",
            "生活自理": "1. 點心時間練習細嚼慢嚥品嚐柚子果肉【咬一咬】\n2. 戴上柚子帽時能練習自己拉下與戴上【脫帽子】\n(指標 V-5-2 / III-5-4)",
            "文化藝術": "1. 藝術創作：黃色水彩塗抹大月亮，貼上黃豆點綴【自由作畫】\n2. 律動兒歌《月亮圓圓》，搭配手部畫大圓動作【載歌載舞】\n(指標 V-4-6 / V-3-5)"
        },
        "長大": {
            "身體動作": "1. 體能課：扮演破殼而出，練習雙腳原地跳躍與平衡走【跳跳虎】\n2. 練習用前兩指撕開紙膠帶、操作拼圖【精細操作】\n(指標 V-1-3 / IV-1-7)",
            "社會情緒": "1. 在鏡子中觀察自己長大的模樣，指認「這是我」【自我認同】\n2. 扮演照顧小玩偶，展現關愛行為【照顧小娃娃】\n(指標 V-2-1 / V-2-8)",
            "語言溝通": "1. 聆聽繪本，模仿動物叫聲【小博士】\n2. 能用短句表達「我長大了」、「我自己做」【以語言表達】\n(指標 V-3-1 / V-3-2)",
            "認知探索": "1. 探索體驗「光滑冰涼觸覺袋」與「溫熱毛巾」【感覺一下】\n2. 進行大與小圖片大小配對【大大小小】\n(指標 V-4-1 / V-4-4)",
            "生活自理": "1. 展現長大常規：練習自己用湯匙把飯舀乾淨【用湯匙進食】\n2. 午睡前練習自己脫下鞋襪並擺整齊【自己脫】\n(指標 V-5-1 / V-5-3)",
            "文化藝術": "1. 藝術創作：蛋殼拼貼畫與羽毛黏貼小動物【撕撕樂】\n2. 配合輕快兒歌進行肢體擺動律動【載歌載舞】\n(指標 V-4-6 / V-3-5)"
        },
        "教師": {
            "身體動作": "1. 體能課：雙人推拉小推車「運送愛心禮物」，練習協調前進【拖拉玩具走】\n2. 將愛心積木依大中小依序堆疊【積木疊疊樂】\n(指標 IV-1-3 / V-1-5)",
            "社會情緒": "1. 主動給老師一個擁抱並說「謝謝老師」【表達情感】\n2. 能指認中心裡各位照顧老師的照片與名字【認識老師】\n(指標 V-2-4 / V-2-7)",
            "語言溝通": "1. 聆聽繪本，學習表達感謝詞彙「謝謝、愛你」\n2. 能回答簡單問題：「最喜歡哪位老師？」【我會問】\n(指標 V-3-1 / V-3-6)",
            "認知探索": "1. 觸摸愛心抱枕（軟綿綿）與感謝卡片（硬紙板）【感覺一下】\n2. 將愛心貼紙貼入指定卡片框線內【我會拼圖】\n(指標 V-4-1 / V-4-5)",
            "生活自理": "1. 用餐前後主動幫忙擦拭桌子、收小碗【自己收玩具】\n2. 如廁完畢後在老師引導下練習拉好褲子【練習穿衣服】\n(指標 V-5-7 / V-5-4)",
            "文化藝術": "1. 藝術創作：「給老師的愛心花束」掌印蓋印與色紙拼貼【手印畫】\n2. 配合兒歌《聽我說謝謝你》進行溫馨律動手指謠【手指謠】\n(指標 V-4-6 / V-3-5)"
        },
        "父親": {
            "身體動作": "1. 體能課：大步跨越軟墊障礙物，模仿爸爸高舉雙手做大樹【大肌肉活動】\n2. 用前兩指捏取撕貼紙裝飾卡片【精細操作】\n(指標 V-1-2 / V-1-6)",
            "社會情緒": "1. 看著爸爸的照片能開心地指認並說出「爸爸」【那是誰？】\n2. 練習給爸爸大大的擁抱表達感謝【表達愛意】\n(指標 V-2-1 / V-2-4)",
            "語言溝通": "1. 專心聆聽繪本，指認故事人物並仿說短語「爸爸抱抱」\n2. 能回答簡單問題「爸爸去哪裡？」【我會問】\n(指標 V-3-2 / V-3-6)",
            "認知探索": "1. 觸摸體驗爸爸的領帶、粗糙鬍渣卡與毛巾【感覺一下】\n2. 能分辨大皮鞋與小鞋子【大大小小】\n(指標 V-4-1 / V-4-4)",
            "生活自理": "1. 活動後能在引導下練習將鞋子整齊放回鞋櫃【自己脫】\n2. 能自行拿水杯喝水並送回水杯架【用湯匙進食】\n(指標 V-5-3 / V-5-1)",
            "文化藝術": "1. 「爸爸的帥氣領帶」手撕色紙黏貼與指印點點畫【撕撕樂】\n2. 配合兒歌《我的好爸爸》進行肢體擺動律動【載歌載舞】\n(指標 V-4-6 / V-3-5)"
        },
        "default": {
            "身體動作": "1. 體能課：跨越軟墊障礙物，練習平衡走與大步跨越【上樓梯/奔跑】\n2. 練習將積木堆疊 3-5 塊【積木疊疊樂】\n(指標 V-1-1 / V-1-5)",
            "社會情緒": "1. 參與團體傳球遊戲，練習輪流與等待【一起丟布球】\n2. 看到同伴哭泣能拿毛巾拍拍安慰【安慰別人】\n(指標 V-2-2 / V-2-5)",
            "語言溝通": "1. 專心聆聽繪本，回答簡單問句【故事魔毯】\n2. 能用短句表達日常生活需求【以語言表達】\n(指標 V-3-4 / V-3-2)",
            "認知探索": "1. 感官觸覺探索：體驗冷熱、乾濕與軟硬不同質感【感覺一下】\n2. 形狀與顏色配對分類遊戲【分類遊戲】\n(指標 V-4-1 / V-4-3)",
            "生活自理": "1. 練習自行用湯匙進食與小口咀嚼【用湯匙吃東西】\n2. 如廁前主動表示便意並配合脫褲【坐小馬桶】\n(指標 V-5-1 / V-5-8)",
            "文化藝術": "1. 自由作畫：利用無毒粗蠟筆或手指膏進行色彩塗鴉【自由作畫】\n2. 配合節奏音樂做大肌肉開展律動【載歌載舞】\n(指標 V-4-6 / V-3-5)"
        }
    },
    "24-36個月": {
        "顏色": {
            "身體動作": "1. 體能課：雙腳連續往前跳躍過色彩障礙線，單腳站立平衡【原地跳躍】\n2. 使用安全剪刀嘗試剪碎彩色紙條，串珠穿繩【工具操作】\n(指標 VI-1-1 / VI-1-3)",
            "社會情緒": "1. 與同伴合作完成彩色大拼圖，練習輪流與分享教具【友伴合作】\n2. 能辨識自己與他人的情緒並說出原因「他哭了因為...」【情緒表達】\n(指標 VI-2-1 / VI-2-6)",
            "語言溝通": "1. 聆聽繪本後能複述完整故事情節【語句理解】\n2. 能主動使用禮貌用語「請、謝謝、對不起」與老師對話【社交用語】\n(指標 VI-3-2 / VI-3-1)",
            "認知探索": "1. 進行三種以上顏色與形狀的「雙重屬性分類」【特徵歸類】\n2. 觀察藍色加黃色調配成綠色的科學色彩實驗【因果探索】\n(指標 VI-4-4 / VI-4-3)",
            "生活自理": "1. 用餐後能自行用小抹布擦拭桌面並將餐具整齊分類回收【自主家事】\n2. 獨立完成脫褲、坐馬桶、穿褲及洗手完整如廁流程【自主如廁】\n(指標 VI-5-7 / VI-5-8)",
            "文化藝術": "1. 藝術創作：利用滴管與水彩進行「色彩暈染拓印畫」創作角色【撕貼美術】\n2. 配合多節奏音樂進行律動舞蹈表演【音樂擺動】\n(指標 VI-1-7 / VI-1-6)"
        },
        "中秋": {
            "身體動作": "1. 體能課：抱球走平衡木、跨跳過呼啦圈月亮步道【推球平衡】\n2. 用雙手手掌搓揉黏土成圓球、壓扁並使用模具印製月餅【手眼協調】\n(指標 VI-1-2 / VI-1-3)",
            "社會情緒": "1. 在中秋扮演角扮演老闆與客人，進行買賣月餅互動【友伴遊戲】\n2. 願意主動分享中秋道具給同儕【公眾分享】\n(指標 VI-2-7 / VI-2-1)",
            "語言溝通": "1. 聆聽繪本，能回答「小星為什麼要吃月餅？」等思考問句【語言描述】\n2. 能說出中秋節相關詞彙（嫦娥、柚子、烤肉、賞月）【特徵詞彙】\n(指標 VI-3-11 / VI-3-3)",
            "認知探索": "1. 比較整顆柚子（重/大）與切片柚子（輕/小）之重量與大小【大小關係】\n2. 觀察月亮從新月、半月到滿月的形狀變化圖卡排序【步驟順序】\n(指標 VI-4-6 / VI-4-3)",
            "生活自理": "1. 嘗試自己剝開柚子果肉白皮，練習小肌肉細緻操作【生活自理】\n2. 能自己練習穿脫鞋子並將鞋帶黏扣帶貼好【自己穿脫】\n(指標 VI-5-7 / VI-5-3)",
            "文化藝術": "1. 「立體玉兔搗藥」：利用紙黏土與棉花球拼貼立體中秋夜景畫【撕貼創作】\n2. 兒歌律動《八月十五中秋節》，練習打節奏敲擊樂器【工具敲擊】\n(指標 VI-1-7 / VI-1-3)"
        },
        "長大": {
            "身體動作": "1. 體能課：雙腳靈活奔跑閃避障礙物，走斜坡學步梯【遵守安全】\n2. 使用鑷子夾取豆子、組裝大型積木建築【積木建構】\n(指標 VI-1-4 / VI-4-2)",
            "社會情緒": "1. 在小組活動中能主動邀請同伴「一起玩好嗎？」【友伴合作】\n2. 遇到挫折時能用語言表達求助而非哭鬧【情緒表達】\n(指標 VI-2-7 / VI-2-6)",
            "語言溝通": "1. 聆聽繪本，能有條理描述動物長大的順序【語句表達】\n2. 能清楚回答「你長大會做什麼？」並發表意見【主動要求】\n(指標 VI-3-2 / VI-3-8)",
            "認知探索": "1. 幼兒生活物品（奶瓶/水杯、尿布/小內褲）前後成長對照分類【特徵歸類】\n2. 完成 8-10 片情境拼圖【鑲嵌拼圖】\n(指標 VI-4-4 / VI-4-1)",
            "生活自理": "1. 練習自行穿脫有拉鍊或大鈕釦的外套【練習穿衣】\n2. 用餐後能自行使用牙刷練習清潔牙齒【自主刷牙】\n(指標 VI-5-4 / VI-5-5)",
            "文化藝術": "1. 「我長大了立體自畫像」：利用紙黏土與毛線拼貼自畫像【美術創作】\n2. 唱跳兒歌《長大歌》，隨音樂變換快慢節奏律動【音樂擺動】\n(指標 VI-1-7 / VI-1-6)"
        },
        "default": {
            "身體動作": "1. 體能課：獨立上下樓梯不扶欄杆、單腳原地跳躍【連續彈跳】\n2. 拿粗筆在框線內著色、拼排 6-10 片拼圖【精細操作】\n(指標 VI-1-1 / VI-1-3)",
            "社會情緒": "1. 能與同伴進行規則性團體遊戲，遵守遊戲常規【遵守原則】\n2. 能用語言表達「我不喜歡這樣，請停下來」【表達喜好】\n(指標 VI-1-4 / VI-2-2)",
            "語言溝通": "1. 專心聆聽繪本，能清楚用完整句子表達想法【語句理解】\n2. 能主動描述今天在中心發生的有趣事情【描述圖片】\n(指標 VI-3-2 / VI-3-11)",
            "認知探索": "1. 能依物品特徵進行「長短、高矮、粗細」比較與排序【大小關係】\n2. 探索兩個物體位置間的上下空間關係【上下關係】\n(指標 VI-4-6 / VI-4-5)",
            "生活自理": "1. 用餐時能正確握持湯匙乾淨進食，不掉飯粒【湯匙進食】\n2. 能在引導下主動將脫下的外套掛上衣架【自己穿脫】\n(指標 VI-5-1 / VI-5-3)",
            "文化藝術": "1. 混合媒材拼貼畫創作（色紙、毛線、大自然落葉花瓣）【撕貼創作】\n2. 配合快慢音樂節奏變換身體律動動作【音樂擺動】\n(指標 VI-1-7 / VI-1-6)"
        }
    }
}

def query_activity(domain, theme, book, age_group):
    group_dict = DATABASE.get(age_group, DATABASE["12-24個月"])
    matched_key = "default"
    for k in group_dict.keys():
        if k != "default" and (k in theme or k in book):
            matched_key = k
            break
    
    act_text = group_dict[matched_key].get(domain, group_dict["default"].get(domain, ""))
    if "{book}" in act_text:
        act_text = act_text.replace("{book}", book if book else "主題繪本")
    return act_text

# --- 產生文件 1：月教案輪值表 ---
def generate_roster_docx(branch_name, year_roc, month, teachers_str, table_data, headers):
    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    r_title = title_p.add_run(f"{branch_name} {year_roc} 年 {month} 月教案輪值表")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(8)
    sub_text = "幼兒發展領域：身體動作、社會情緒、語言溝通、認知探索、生活自理"
    if teachers_str.strip():
        sub_text += f"\n主帶托育人員：{teachers_str}"
    r_sub = sub_p.add_run(sub_text)
    r_sub.font.name = "微軟正黑體"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    table = doc.add_table(rows=len(table_data) + 1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="1F497D", sz="4")

    col_widths = [Inches(1.5), Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8)]

    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "1F497D")
        set_cell_margins(cell, top=80, bottom=80, left=60, right=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "微軟正黑體"
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.size = Pt(12)

    special_keywords = ["特約醫師", "牙醫師", "消防演練", "大型活動", "國定假日", "教案暫停", "連假", "中秋節", "教師節", "體能暫停"]
    for row_idx, row_content in enumerate(table_data, start=1):
        for col_idx, text in enumerate(row_content):
            cell = table.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            set_cell_margins(cell, top=70, bottom=70, left=60, right=60)
            p = cell.paragraphs[0]
            
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_background(cell, "DCE6F1")
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_cell_background(cell, "FAFAFA" if row_idx % 2 == 1 else "FFFFFF")

            r = p.add_run(text)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(12)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            elif any(k in text for k in special_keywords):
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- 產生文件 2：適性發展活動月計畫 (進階版) ---
def generate_month_plan_docx(branch_name, year_roc, month, age_group, teachers_str, weekly_schedule):
    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    r_title = title_p.add_run(f"{branch_name} 適性發展活動 {year_roc} 年 {month} 月計畫（進階版）")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(4)
    r_sub = sub_p.add_run(f"月份：{year_roc} 年 {month} 月 ｜ 年齡/組別：{age_group} ｜ 主責托育人員：{teachers_str}")
    r_sub.font.name = "微軟正黑體"
    r_sub.font.size = Pt(9.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    note_p = doc.add_paragraph()
    note_p.paragraph_format.space_after = Pt(4)
    r_n = note_p.add_run("填表原則：1. 週一需多了解週末嬰幼兒於家庭生活作息概況，多給予親密擁抱與個別接觸。 2. 依據嬰幼兒個別差異與興趣選擇活動，融入日常作息中重複進行。 3. 結合主題繪本與適性發展六大領域設計。")
    r_n.font.name = "微軟正黑體"
    r_n.font.size = Pt(8.5)
    r_n.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    table = doc.add_table(rows=len(weekly_schedule) + 1, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="1F497D", sz="4")

    col_widths = [Inches(1.4), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)]
    headers = ["週次 / 主題", "身體動作\n(粗大/精細動作)", "社會情緒\n(自我/同儕關係)", "語言溝通\n(表達/接收性語言)", "認知探索\n(感官/概念/問題解決)", "生活自理\n(衛生/健康習慣)", "文化藝術\n(音樂律動/美勞創作)"]

    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "1F497D")
        set_cell_margins(cell, top=50, bottom=50, left=50, right=50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "微軟正黑體"
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.size = Pt(9)

    for row_idx, item in enumerate(weekly_schedule, start=1):
        w_title = f"{item['week_range']}\n【主題：{item['theme']}】\n繪本：《{item['book']}》\n主責：{item['roster_summary']}"
        row_content = [
            w_title,
            query_activity("身體動作", item['theme'], item['book'], age_group),
            query_activity("社會情緒", item['theme'], item['book'], age_group),
            query_activity("語言溝通", item['theme'], item['book'], age_group),
            query_activity("認知探索", item['theme'], item['book'], age_group),
            query_activity("生活自理", item['theme'], item['book'], age_group),
            query_activity("文化藝術", item['theme'], item['book'], age_group)
        ]
        for col_idx, text in enumerate(row_content):
            cell = table.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            set_cell_margins(cell, top=45, bottom=45, left=45, right=45)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(text)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(8)
            if col_idx == 0:
                set_cell_background(cell, "DCE6F1")
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                set_cell_background(cell, "FAFAFA" if row_idx % 2 == 1 else "FFFFFF")

    doc.add_paragraph().paragraph_format.space_after = Pt(3)
    t_env = doc.add_table(rows=2, cols=1)
    t_env.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_env, color="1F497D", sz="4")

    c_h = t_env.cell(0, 0)
    set_cell_background(c_h, "1F497D")
    set_cell_margins(c_h, top=30, bottom=30, left=60, right=60)
    p = c_h.paragraphs[0]
    r = p.add_run("教材、情境、環境調整與教玩具規劃")
    r.font.name = "微軟正黑體"
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.size = Pt(9)

    c_b = t_env.cell(1, 0)
    set_cell_background(c_b, "FAFAFA")
    set_cell_margins(c_b, top=40, bottom=40, left=60, right=60)
    p = c_b.paragraphs[0]
    all_books = "、".join([f"《{it['book']}》" for it in weekly_schedule if it['book']])
    
    if age_group in ["0-12個月", "7-12個月"]:
        env_text = (
            f"1. 語文圖書角：配合各週主題展示繪本 {all_books}、布書與厚紙板書，設置於低矮防撞區域供嬰兒爬行探索。\n"
            "2. 感官探索區：備妥各色安全絲巾、手搖鈴、水感墊、觸覺球、溫熱毛巾，提供多樣觸覺與聽覺刺激。\n"
            "3. 益智操作與生活區：提供抓握玩具、形狀套杯、大按鍵發聲玩具、安全鏡面與雙耳學習杯。\n"
            "4. 大肌肉活動區：鋪設安全防撞爬行地墊、軟質斜坡、滾動大皮球，保持安全開闊的爬行與扶站動線。"
        )
    elif age_group == "12-24個月":
        env_text = (
            f"1. 語文圖書角：配合各週主題展示繪本 {all_books}，設置於 90 公分以下開放式矮櫃供自主取閱。\n"
            "2. 感官探索區：備妥各色絲巾、毛球/夾子、主題觸覺探索箱（水感袋/豆類/柚子皮）、無毒黏土與海綿滾筒刷。\n"
            "3. 益智操作與扮演區：提供主題形狀鑲嵌拼圖、彩色積木、大小套杯、辦家家酒道具、手偶與托育人員生活照片。\n"
            "4. 大肌肉活動區：設置安全防撞軟墊斜坡、呼啦圈步道、大球/布球、小推車、平衡木，規劃安全順暢的體能動線。"
        )
    else: # 24-36個月
        env_text = (
            f"1. 語文圖書角：配合各週主題展示繪本 {all_books} 與百科認知圖鑑，規劃安靜共讀角培養自主閱讀習慣。\n"
            "2. 感官探索區：備妥安全剪刀、色紙、穿線串珠板、水彩顏料、滴管、天平秤重儀與科學探索盆。\n"
            "3. 益智操作與扮演區：提供 6-10 片進階拼圖、雙重屬性分類盒、情境角色扮演服飾道具（廚房/醫院/中秋）。\n"
            "4. 大肌肉活動區：設置跳躍障礙線、體能跳箱、平衡木、投擲標靶籃，規劃具挑戰性的大肌肉體能路徑。"
        )

    r_body = p.add_run(env_text)
    r_body.font.name = "微軟正黑體"
    r_body.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(3)
    t_sign = doc.add_table(rows=1, cols=3)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_sign, color="1F497D", sz="4")
    sign_widths = [Inches(3.5), Inches(3.5), Inches(3.5)]
    sign_titles = ["帶班托育人員簽章：", "托育組長核閱：", "中心主任核閱："]

    for i, st_title in enumerate(sign_titles):
        c = t_sign.cell(0, i)
        c.width = sign_widths[i]
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, top=50, bottom=50, left=50, right=50)
        p = c.paragraphs[0]
        r = p.add_run(f"{st_title}\n\n日期：  年 月 日")
        r.font.name = "微軟正黑體"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- Web 介面佈局 ---
st.title("🏫 澳森托嬰中心 智慧教案與適性月計畫系統")
st.markdown("主任可選擇分園、年齡組別、設定月份與老師名單，系統會自動切換對應體能課與**標準適性指標代碼與活動標題**，並**一鍵生成兩種標準 Word 檔**！")

# 1. 基本設定區
st.subheader("壹、 基本設定")
c1, c2, c3, c4 = st.columns([1.2, 1, 1, 1.4])
with c1:
    branch = st.selectbox("園所分園：", ["澳森", "澳森文德"])

# 依分園動態提供對應之月齡版本
if branch == "澳森":
    available_ages = ["0-12個月", "12-24個月", "24-36個月"]
else:
    available_ages = ["7-12個月", "12-24個月", "24-36個月"]

with c4:
    age_group = st.selectbox("年齡階段版本：", available_ages)

with c2:
    year_roc = st.number_input("民國年份：", min_value=114, max_value=125, value=115)
with c3:
    month = st.selectbox("月份：", list(range(1, 13)), index=8)  # 預設 9 月

teachers_input = st.text_input("主帶老師名單（請以逗號分隔，若無可留空）：", value="", placeholder="例如：秋馨, 綺綺, 悅熏, 知容, 政勳, 閔興, 怡君, Candy, Panda, 小安, 樺樺, 均宜")

builtin_options = ["（請選擇）", "主任", "特約醫師", "牙醫師", "消防演練", "大型活動", "國定假日", "教案暫停"]
raw_teachers = [t.strip() for t in teachers_input.split(",") if t.strip()]
dropdown_options = builtin_options[:2] + raw_teachers + builtin_options[2:]

# 依分園動態調整表頭與體能課位置
if branch == "澳森":
    # 澳森：星期三為小肌肉/觸覺，星期四為體能課
    headers = ["主題", "星期一：繪本", "星期二：小肌肉／認知", "星期三：小肌肉／觸覺", "星期四：體能課", "星期五：藝術創作"]
    wed_title_prefix = "週三：觸覺"
    thu_title_prefix = "週四：體能"
else:
    # 澳森文德：星期三為體能課，星期四為小肌肉/觸覺
    headers = ["主題", "星期一：繪本", "星期二：小肌肉／認知", "星期三：體能課", "星期四：小肌肉／觸覺", "星期五：藝術創作"]
    wed_title_prefix = "週三：體能"
    thu_title_prefix = "週四：觸覺"

st.divider()

# 2. 月曆排班工作區
st.subheader(f"貳、 {branch}（{age_group}）{year_roc} 年 {month} 月 課程輪值排班（週一至週五）")

year_ad = year_roc + 1911
cal = calendar.monthcalendar(year_ad, month)

work_weeks = []
for w in cal:
    mon_to_fri = w[0:5]
    if any(d > 0 for d in mon_to_fri):
        work_weeks.append(mon_to_fri)

table_data = []
weekly_schedule_for_plan = []

for idx, w in enumerate(work_weeks, start=1):
    with st.expander(f"📌 第 {idx} 週排班設定", expanded=True):
        col_theme, col_m, col_t, col_w, col_th, col_f = st.columns([1.2, 2.2, 1.3, 1.3, 1.3, 1.3])
        
        d_mon = f"{month}/{w[0]}" if w[0] > 0 else ""
        d_tue = f"{month}/{w[1]}" if w[1] > 0 else ""
        d_wed = f"{month}/{w[2]}" if w[2] > 0 else ""
        d_thu = f"{month}/{w[3]}" if w[3] > 0 else ""
        d_fri = f"{month}/{w[4]}" if w[4] > 0 else ""

        valid_days = [d for d in w if d > 0]
        w_start = f"{month:02d}/{valid_days[0]:02d}"
        w_end = f"{month:02d}/{valid_days[-1]:02d}"
        week_range_str = f"第{['一','二','三','四','五'][idx-1]}週 ({w_start}-{w_end})"

        if idx == 1:
            theme_default = "顏色"
            book_name_default = "小藍與小黃"
            m_lead_default_idx = 1
        else:
            theme_default = ""
            book_name_default = ""
            m_lead_default_idx = 0

        with col_theme:
            theme = st.text_input("週主題", key=f"th_{idx}", value=theme_default, placeholder="請輸入主題")

        with col_m:
            st.markdown(f"**週一：繪本 ({d_mon})**" if d_mon else "**週一：繪本**")
            m_lead = st.selectbox("導讀人", dropdown_options, key=f"ml_{idx}", index=m_lead_default_idx)
            book_name = st.text_input("繪本名稱", key=f"bn_{idx}", value=book_name_default, placeholder="請輸入繪本名稱")
            
            lead_txt = "" if m_lead == "（請選擇）" else f" {m_lead}"
            bk_txt = f"\n{book_name}" if book_name.strip() else ""
            mon_str = f"{d_mon}{lead_txt}{bk_txt}".strip() if d_mon else ""

        with col_t:
            st.markdown(f"**週二：認知 ({d_tue})**" if d_tue else "**週二：認知**")
            tue_teacher = st.selectbox("負責人/狀態", dropdown_options, key=f"t_{idx}", index=0)
            t_txt = "" if tue_teacher == "（請選擇）" else f" {tue_teacher}"
            tue_str = f"{d_tue}{t_txt}".strip() if d_tue else ""

        with col_w:
            st.markdown(f"**{wed_title_prefix} ({d_wed})**" if d_wed else f"**{wed_title_prefix}**")
            wed_teacher = st.selectbox("負責人/狀態", dropdown_options, key=f"w_{idx}", index=0)
            w_txt = "" if wed_teacher == "（請選擇）" else f" {wed_teacher}"
            wed_str = f"{d_wed}{w_txt}".strip() if d_wed else ""

        with col_th:
            st.markdown(f"**{thu_title_prefix} ({d_thu})**" if d_thu else f"**{thu_title_prefix}**")
            thu_teacher = st.selectbox("負責人/狀態", dropdown_options, key=f"th_t_{idx}", index=0)
            th_txt = "" if thu_teacher == "（請選擇）" else f" {thu_teacher}"
            thu_str = f"{d_thu}{th_txt}".strip() if d_thu else ""

        with col_f:
            st.markdown(f"**週五：藝術 ({d_fri})**" if d_fri else "**週五：藝術**")
            fri_teacher = st.selectbox("負責人/狀態", dropdown_options, key=f"f_{idx}", index=0)
            f_txt = "" if fri_teacher == "（請選擇）" else f" {fri_teacher}"
            fri_str = f"{d_fri}{f_txt}".strip() if d_fri else ""

        table_data.append([theme, mon_str, tue_str, wed_str, thu_str, fri_str])
        
        assigned_teachers = [x for x in [m_lead, tue_teacher, wed_teacher, thu_teacher, fri_teacher] if x not in ["（請選擇）", "教案暫停", "國定假日"]]
        roster_summary = "/".join(assigned_teachers) if assigned_teachers else "全體托育人員"
        
        weekly_schedule_for_plan.append({
            "week_range": week_range_str,
            "theme": theme if theme else "主題探索",
            "book": book_name if book_name else "主題繪本導讀",
            "roster_summary": roster_summary
        })

# 3. 成果匯出區
st.divider()
st.subheader("參、 成果匯出")

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    st.markdown(f"#### 1. 下載【{branch} 月教案輪值表】")
    st.caption("格式：橫向 A4、字體 12pt、排定每日負責老師與活動類別。")
    doc_roster_bytes = generate_roster_docx(branch, year_roc, month, teachers_input, table_data, headers)
    st.download_button(
        label=f"📥 下載【{branch} {year_roc}年{month}月教案輪值表】(Word / 12pt)",
        data=doc_roster_bytes,
        file_name=f"{branch}_{year_roc}年{month}月教案輪值表.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with col_btn2:
    st.markdown(f"#### 2. 下載【{branch}（{age_group}）適性月計畫】")
    st.caption("格式：營運手冊範例 2-3-4 格式，自動套用專屬年齡之適性發展六大領域指標與環境規劃。")
    doc_plan_bytes = generate_month_plan_docx(branch, year_roc, month, age_group, teachers_input, weekly_schedule_for_plan)
    st.download_button(
        label=f"📥 下載【{branch}（{age_group}）適性月計畫】(Word / 進階版)",
        data=doc_plan_bytes,
        file_name=f"{branch}_{age_group}_適性發展活動_{year_roc}年{month}月計畫(進階版).docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
