#!/usr/bin/env python3
"""生成2万个旅行行业相关测试关键词"""

import random
import itertools

random.seed(42)

# ═══════════════════════════════════════════════════
# 基础词库
# ═══════════════════════════════════════════════════

CITIES = [
    '北京','上海','广州','深圳','成都','重庆','杭州','武汉','西安','南京',
    '长沙','苏州','天津','郑州','青岛','大连','沈阳','哈尔滨','长春','济南',
    '昆明','贵阳','南宁','海口','三亚','厦门','福州','泉州','南昌','合肥',
    '太原','石家庄','兰州','银川','西宁','呼和浩特','乌鲁木齐','拉萨','无锡','宁波',
    '温州','佛山','东莞','珠海','中山','惠州','扬州','徐州','常州','绍兴',
    '嘉兴','金华','台州','洛阳','开封','安阳','许昌','南阳','桂林','柳州',
    '北海','梧州','大理','丽江','香格里拉','西双版纳','腾冲','曲靖','玉溪','张家界',
    '凤凰','岳阳','衡阳','九江','景德镇','婺源','赣州','上饶','吉安','烟台',
    '威海','日照','泰安','曲阜','承德','秦皇岛','唐山','保定','邯郸','襄阳',
    '宜昌','恩施','十堰','遵义','安顺','凯里','铜仁','宜宾','泸州','乐山',
    '绵阳','德阳','南充','达州','汕头','潮州','湛江','河源','清远','韶关',
    '梅州','肇庆','阳江','茂名','丽水','衢州','舟山','湖州','连云港','盐城',
    '淮安','宿迁','泰州','南通','镇江','安吉','千岛湖','普陀山','九华山','五台山',
    '博鳌','文昌','万宁','康定','甘孜','阿坝','攀枝花','湘潭','株洲','常德',
    '益阳','郴州','邵阳','南平','龙岩','三明','漳州','莆田','延吉','吉林',
    '芜湖','蚌埠','马鞍山','黄山','淄博','潍坊','临沂','聊城','德州','滨州',
    '汕尾','揭阳','白银','天水','酒泉','嘉峪关','张掖','长治','大同','运城',
    '晋中','临汾','菏泽','信阳','汉中','渭南','楚雄','红河','文山','临沧',
    '普洱','德宏','怒江','包头','鄂尔多斯','涠洲岛','六安','丽江','敦煌','峨眉山',
]

INTL_CITIES = [
    '东京','大阪','京都','首尔','釜山','济州','曼谷','清迈','普吉','巴厘岛',
    '岘港','芽庄','吉隆坡','槟城','马尼拉','暹粒','河内','胡志明','金边','伦敦',
    '巴黎','罗马','米兰','威尼斯','佛罗伦萨','巴塞罗那','马德里','柏林','慕尼黑','法兰克福',
    '阿姆斯特丹','布鲁塞尔','苏黎世','维也纳','布拉格','布达佩斯','纽约','洛杉矶','旧金山','拉斯维加斯',
    '夏威夷','温哥华','多伦多','悉尼','墨尔本','伊斯坦布尔','圣托里尼','冲绳','苏梅岛','马六甲',
    '亚庇','明洞','北海道','芭提雅','长滩岛','宿务','富国岛','雅加达','奥克兰','迪拜',
]

COUNTRIES = [
    '日本','韩国','泰国','越南','新加坡','马来西亚','菲律宾','柬埔寨','缅甸','老挝',
    '印度','尼泊尔','斯里兰卡','马尔代夫','阿联酋','迪拜','卡塔尔','沙特','以色列','土耳其',
    '英国','法国','德国','意大利','西班牙','葡萄牙','荷兰','比利时','瑞士','奥地利',
    '希腊','捷克','匈牙利','波兰','瑞典','挪威','芬兰','丹麦','冰岛','爱尔兰','俄罗斯',
    '美国','加拿大','墨西哥','巴西','阿根廷','秘鲁','智利','古巴',
    '澳大利亚','新西兰','斐济','埃及','南非','摩洛哥','肯尼亚','坦桑尼亚',
    '香港','澳门','台湾',
]

HOTEL_WORDS = [
    '酒店','住宿','宾馆','旅馆','民宿','客栈','旅社','旅店','旅舍','公寓',
    '别墅','木屋','帐篷','度假村','青旅','招待所','旅馆','商务酒店','精品酒店','主题酒店',
]

HOTEL_QUALIFIERS = [
    '五星级','四星级','三星级','豪华','经济型','连锁','亲子','温泉','网红','精品',
    '度假','商务','情侣','海景','湖景','山景','江景','园景','套房','大床房',
    '双床房','家庭房','亲子房','标间','特价','便宜','性价比','高端','五星','三星',
    '泳池','浴缸','接机','海景房','阳台','早餐','自助','免费取消','即时确认','当日',
    '钟点房','月租','长租','短租','日租','今晚','明天','周末','节假日',
]

HOTEL_CHAINS = [
    '如家','锦江','汉庭','7天','全季','亚朵','维也纳','希尔顿','万豪','洲际',
    '香格里拉','喜来登','凯悦','皇冠','丽思','索菲特','诺富特','桔子','桔子水晶','宜必思',
    '速8','格林豪泰','布丁','尚客优','开元','君澜','雷迪森','金陵','花间堂','悦榕庄',
    '万枫','福朋','雅乐轩','艾美','美居','美爵','锦江之星','如家商旅','如家精选','素柏',
    '白玉兰','欢朋','枫渡','丽枫','喆啡','希岸','潮漫','ZMAX','云上四季','康铂',
    '美豪','书香世家','书香府邸',
]

FLIGHT_WORDS = [
    '机票','航班','飞机','航空','特价机票','打折机票','便宜机票','往返机票','单程机票',
    '机票预订','订机票','买机票','国际航班','国内航班','民航','飞往','航班号','舱位',
    '头等舱','商务舱','经济舱','公务舱','超级经济舱','机票查询','机票比价','联程机票',
    '中转','转机','经停','红眼航班','廉航','低成本航空','儿童机票','婴儿机票','留学生机票',
    '学生机票','随心飞','次卡','机票盲盒','盲盒','白金卡','金鹏卡','机票券','无限飞',
    '青春卡','敬老卡','家庭卡','套票','往返次卡','飞行卡','机票包','机酒套餐',
]

FLIGHT_AUX = [
    '托运','行李额','行李托运','随身行李','选座','值机','网上值机','在线值机',
    '自助值机','登机流程','飞机餐','餐食','机上wifi','机票退改','退票','改签',
    '误机','延误险','航意险','退改签','电子客票','行程单','燃油税','燃油附加费',
    '机建费','机建燃油','机场建设费','里程','积分','升舱','里程兑换','积分兑换',
    '常旅客','飞行里程','会员日','会员价','机场','火车站','高铁站','南站','北站',
    '东站','西站','候机','安检','通关','边检',
]

TRAIN_WORDS = [
    '火车票','高铁票','动车票','火车','高铁','动车','列车','车票','硬座','软座',
    '硬卧','软卧','无座','二等座','一等座','商务座','车次','时刻表','抢票','候补',
    '铁路','铁道','城际','城铁','12306',
]

SCENIC_WORDS = [
    '景区','景点','门票','古镇','古城','博物馆','纪念馆','长城','故宫','兵马俑',
    '西湖','黄山','泰山','九寨沟','张家界','迪士尼','环球影城','欢乐谷','长隆','方特',
    '海洋馆','水族馆','动物园','游乐园','乐园','缆车','索道','观光车','演出','表演',
    '瀑布','溶洞','预约','购票','免票','半票','开放时间','票价','门票价格',
]

DESTINATIONS = [
    '阿那亚','金町湾','日月湾','海棠湾','亚龙湾','三亚湾','大东海','清水湾','石梅湾',
    '神州半岛','蜈支洲岛','西岛','分界洲岛','南湾猴岛','涠洲岛','银滩','侨港','千岛湖',
    '太湖','瘦西湖','洱海','滇池','泸沽湖','抚仙湖','莫干山','安吉','桐庐','溧阳',
    '德清','北戴河','南戴河','黄金海岸','长隆','海昌','融创','亚布力','北大湖','松花湖',
    '长白山','牛首山','灵山','普陀山','九华山','峨眉山','武当山','青城山','华山',
    '庐山','雁荡山','武夷山','四姑娘山','稻城亚丁','黄龙','毕棚沟','海螺沟',
    '西双版纳','香格里拉','凤凰古城','丽江古城','乌镇','西塘','南浔','周庄','同里',
    '朱家角','鼓浪屿','曾厝垵','双月湾','巽寮湾','红海湾','武功山','三清山','龙虎山',
    '千户苗寨','三坊七巷','官也街','窑湖小镇','罗浮山','老君山','螺髻山','东极岛',
    '南澳岛','阿勒泰','中华恐龙园','宽窄巷子','平江路','望仙谷','圣水洞','赛里木湖',
    '松阳','玄武湖','顺德','锦州','忻州','本溪',
]

INTL_SCENIC = [
    '乐天世界','环球影城','圣托里尼','埃菲尔','卢浮宫','大英博物馆','浅草寺','明洞',
    '暹罗天地','鱼尾狮','自由女神','泰姬陵','金字塔','富士山','阿尔卑斯','峡湾','极光',
    '吴哥窟','蒲甘','大堡礁','马丘比丘',
]

FESTIVALS = [
    '五一','国庆','春节','清明','端午','中秋','元旦','寒假','暑假','黄金周',
    '小长假','情人节','七夕','圣诞','跨年','过年','除夕','元宵','春游','秋游',
]

GENERIC_INTENT = [
    '去哪','去哪玩','去哪旅游','去哪里','去哪里玩','去哪里旅游','适合去','适合去哪',
    '适合去哪里','适合旅游','适合旅行','适合度假','旅游去','旅行去','度假去',
    '假期去哪','放假去哪','周末去哪','暑假去哪','寒假去哪','周边游','短途游',
    '冷门城市','冷门景点','小众旅游','小众目的地','宝藏旅游','自驾游','房车','露营',
    '旅游推荐','旅行推荐','目的地推荐','出行','出游','出去玩','亲子旅游','亲子游',
    '带爸妈','带娃','毕业旅行','蜜月旅行','温泉旅游','温泉度假','泡温泉','私汤',
    '海边旅游','海边度假','海岛游','有雪的地方','滑雪','两天一夜','三天两夜',
    '美食推荐','美食攻略','穷游','背包客','攻略','自由行','必去','必玩','旅游路线',
    '行程','路线','怎么玩','好玩','避坑','打卡',
]

INTL_SCENE = [
    '免税','退税','换汇','代购','演唱会','购物','必买','免税店','药妆','化妆品',
    '血拼','奥特莱斯','签证','入境','过关','通关','货币','汇率','小费','必吃',
    '美食','小吃','夜市','餐厅推荐',
]

INTL_PRACTICAL = [
    '打车','租车','当地交通','天气','最佳时间','几月去','什么时候去','安全','注意事项',
    '电话卡','sim卡','拍照','网红',
]

PRICE_WORDS = [
    '便宜','性价比','特价','打折','划算','优惠','低价','百元','平价','实惠',
    '折扣','促销','秒杀','省钱','捡漏','比价','多少钱','最便宜','最低价','优惠价',
]

BRAND_WORDS = [
    '同程','同程旅行','携程','去哪儿','美团','飞猪','途牛','马蜂窝','驴妈妈','艺龙',
    'Booking','Agoda','Expedia','Airbnb','爱彼迎',
]

AIRLINES = [
    '国航','南航','东航','海航','川航','厦航','深航','山航','春秋航空','吉祥航空',
    '首都航空','九元航空','西部航空','长安航空','天津航空','祥鹏航空','华夏航空','成都航空',
    '国泰航空','港龙航空','澳门航空','长荣航空','中华航空','全日空','日航','大韩航空',
    '韩亚航空','济州航空','泰航','亚航','酷航','虎航','越捷航空','狮航','新航','马航',
    '达美航空','美联航','美国航空','阿联酋航空','卡塔尔航空','土耳其航空','英国航空',
    '汉莎航空','法航','荷航','瑞士航空','澳洲航空','新西兰航空',
]


# ═══════════════════════════════════════════════════
# 组合生成策略
# ═══════════════════════════════════════════════════

keywords = set()

# 1. 国内酒店词 (约5000个)
for city in CITIES:
    for hw in HOTEL_WORDS:
        keywords.add(f"{city}{hw}")
        keywords.add(f"{city}{random.choice(HOTEL_QUALIFIERS)}{hw}")
    for chain in random.sample(HOTEL_CHAINS, min(8, len(HOTEL_CHAINS))):
        keywords.add(f"{city}{chain}{random.choice(HOTEL_WORDS)}")
    for pq in random.sample(HOTEL_QUALIFIERS, min(5, len(HOTEL_QUALIFIERS))):
        keywords.add(f"{city}{pq}{random.choice(HOTEL_WORDS)}")

# 2. 国际酒店词 (约2000个)
for city in INTL_CITIES:
    for hw in random.sample(HOTEL_WORDS, 5):
        keywords.add(f"{city}{hw}")
    for pq in random.sample(HOTEL_QUALIFIERS, 3):
        keywords.add(f"{city}{pq}{random.choice(HOTEL_WORDS)}")
for country in COUNTRIES:
    for hw in random.sample(HOTEL_WORDS, 3):
        keywords.add(f"{country}{hw}")
    for pq in random.sample(HOTEL_QUALIFIERS, 2):
        keywords.add(f"{country}{pq}{random.choice(HOTEL_WORDS)}")

# 3. 机票词 (约3000个)
for city in CITIES[:40]:
    for fc in random.sample(CITIES[:40], 8):
        if city != fc:
            keywords.add(f"{city}到{fc}机票")
            keywords.add(f"{city}飞{fc}")
            keywords.add(f"{city}-{fc}机票")
for fw in FLIGHT_WORDS:
    for city in random.sample(CITIES, 5):
        keywords.add(f"{city}{fw}")
    for country in random.sample(COUNTRIES, 3):
        keywords.add(f"{country}{fw}")
for airline in random.sample(AIRLINES, 15):
    for fw in random.sample(FLIGHT_WORDS, 5):
        keywords.add(f"{airline}{fw}")
for fa in random.sample(FLIGHT_AUX, 15):
    for city in random.sample(CITIES, 5):
        keywords.add(f"{city}{fa}")

# 4. 火车票词 (约2000个)
for city in CITIES[:50]:
    for fc in random.sample(CITIES[:50], 8):
        if city != fc:
            keywords.add(f"{city}到{fc}高铁票")
            keywords.add(f"{city}到{fc}火车票")
            keywords.add(f"{city}-{fc}动车")
for tw in TRAIN_WORDS:
    for city in random.sample(CITIES, 5):
        keywords.add(f"{city}{tw}")
    keywords.add(f"12306{tw}")

# 5. 景区词 (约3000个)
for dest in DESTINATIONS:
    for sw in random.sample(SCENIC_WORDS, 5):
        keywords.add(f"{dest}{sw}")
    for pq in random.sample(PRICE_WORDS, 3):
        keywords.add(f"{dest}{pq}")
for city in CITIES[:40]:
    for sw in random.sample(SCENIC_WORDS, 5):
        keywords.add(f"{city}{sw}")
    for dest in random.sample(DESTINATIONS, 3):
        keywords.add(f"{city}到{dest}")

# 6. 国际景区词 (约1500个)
for city in INTL_CITIES:
    for sw in random.sample(SCENIC_WORDS, 3):
        keywords.add(f"{city}{sw}")
for scenic in INTL_SCENIC:
    for sw in random.sample(SCENIC_WORDS, 3):
        keywords.add(f"{scenic}{sw}")
    keywords.add(f"{scenic}门票")
    keywords.add(f"{scenic}攻略")

# 7. 通用旅游意图词 (约1500个)
for city in CITIES[:50]:
    for gi in random.sample(GENERIC_INTENT, 5):
        keywords.add(f"{city}{gi}")
for country in COUNTRIES:
    for gi in random.sample(GENERIC_INTENT, 3):
        keywords.add(f"{country}{gi}")
for dest in DESTINATIONS[:30]:
    for gi in random.sample(GENERIC_INTENT, 3):
        keywords.add(f"{dest}{gi}")

# 8. 品牌词 (约500个)
for brand in BRAND_WORDS:
    for hw in random.sample(HOTEL_WORDS, 3):
        keywords.add(f"{brand}{hw}")
    for fw in random.sample(FLIGHT_WORDS, 3):
        keywords.add(f"{brand}{fw}")
    for tw in random.sample(TRAIN_WORDS, 2):
        keywords.add(f"{brand}{tw}")
    keywords.add(f"{brand}优惠")
    keywords.add(f"{brand}特价")
    keywords.add(f"{brand}会员")

# 9. 航司词 (约500个)
for airline in AIRLINES:
    keywords.add(f"{airline}机票")
    keywords.add(f"{airline}航班")
    keywords.add(f"{airline}官网")
    keywords.add(f"{airline}会员")
    keywords.add(f"{airline}里程")

# 10. 节日+场景 (约500个)
for fest in FESTIVALS:
    for city in random.sample(CITIES, 8):
        keywords.add(f"{fest}{city}旅游")
        keywords.add(f"{fest}{city}酒店")
        keywords.add(f"{fest}{city}机票")
    for dest in random.sample(DESTINATIONS, 5):
        keywords.add(f"{fest}去{dest}")
        keywords.add(f"{fest}{dest}攻略")

# 11. 国际场景词 (约500个)
for city in INTL_CITIES:
    for sw in random.sample(INTL_SCENE, 3):
        keywords.add(f"{city}{sw}")
for country in COUNTRIES:
    for sw in random.sample(INTL_PRACTICAL, 2):
        keywords.add(f"{country}{sw}")
    for sw in random.sample(INTL_SCENE, 2):
        keywords.add(f"{country}{sw}")

# 12. 价格敏感词 (约300个)
for city in random.sample(CITIES, 20):
    for pw in random.sample(PRICE_WORDS, 5):
        keywords.add(f"{city}{pw}酒店")
        keywords.add(f"{city}{pw}机票")
        keywords.add(f"{pw}{city}旅游")

# 13. 混合长尾词 (补充到2万)
long_tail_templates = [
    '{city}到{city2}怎么走',
    '{city}旅游几天合适',
    '{city}必吃美食',
    '{city}旅游要多少钱',
    '{city}自由行攻略',
    '{city}跟团游价格',
    '{city}天气什么时候最好',
    '{city}酒店住哪里方便',
    '{city}机场到市区怎么走',
    '{city}有什么好玩的地方',
    '{city}周边一日游',
    '{city}两日游路线',
    '{city}三日游攻略',
    '{city}亲子游推荐',
    '{city}情侣旅游攻略',
    '{city}穷游攻略',
    '{city}自驾游路线',
    '{city}网红打卡地',
    '{country}旅游多少钱',
    '{country}签证怎么办',
    '{country}自由行攻略',
    '{country}跟团游价格',
    '{country}最佳旅游时间',
    '{country}安全吗',
    '{dest}门票多少钱',
    '{dest}开放时间',
    '{dest}怎么去',
    '{dest}附近酒店',
    '{dest}好玩吗',
    '{dest}避坑指南',
    '{airline}退改签规则',
    '{airline}行李额',
    '{airline}网上值机',
    '{brand}优惠券',
    '{brand}会员价',
    '{brand}新人优惠',
    '{fest}放假安排',
    '{fest}出行高峰',
]

counter = 0
while len(keywords) < 20000 and counter < 100000:
    template = random.choice(long_tail_templates)
    if '{city2}' in template:
        c1, c2 = random.sample(CITIES, 2)
        kw = template.replace('{city}', c1).replace('{city2}', c2)
    elif '{city}' in template:
        kw = template.replace('{city}', random.choice(CITIES))
    elif '{country}' in template:
        kw = template.replace('{country}', random.choice(COUNTRIES))
    elif '{dest}' in template:
        kw = template.replace('{dest}', random.choice(DESTINATIONS))
    elif '{airline}' in template:
        kw = template.replace('{airline}', random.choice(AIRLINES))
    elif '{brand}' in template:
        kw = template.replace('{brand}', random.choice(BRAND_WORDS))
    elif '{fest}' in template:
        kw = template.replace('{fest}', random.choice(FESTIVALS))
    else:
        continue
    keywords.add(kw)
    counter += 1

# 如果还不够，再补充随机组合
if len(keywords) < 20000:
    extra_templates = [
        '{city}{hw}{pq}',
        '{intl}{hw}',
        '{city}到{intl}机票',
        '{country}签证',
        '{city}周末{gi}',
        '{city}{tw}',
        '{dest}一日游',
        '{dest}两日游',
        '{airline}{fw}',
        '{brand}{hw}',
    ]
    while len(keywords) < 20000:
        t = random.choice(extra_templates)
        if t == '{city}{hw}{pq}':
            kw = random.choice(CITIES) + random.choice(HOTEL_WORDS) + random.choice(HOTEL_QUALIFIERS)
        elif t == '{intl}{hw}':
            kw = random.choice(INTL_CITIES) + random.choice(HOTEL_WORDS)
        elif t == '{city}到{intl}机票':
            kw = random.choice(CITIES) + '到' + random.choice(INTL_CITIES) + '机票'
        elif t == '{country}签证':
            kw = random.choice(COUNTRIES) + '签证'
        elif t == '{city}周末{gi}':
            kw = random.choice(CITIES) + '周末' + random.choice(GENERIC_INTENT)
        elif t == '{city}{tw}':
            kw = random.choice(CITIES) + random.choice(TRAIN_WORDS)
        elif t == '{dest}一日游':
            kw = random.choice(DESTINATIONS) + '一日游'
        elif t == '{dest}两日游':
            kw = random.choice(DESTINATIONS) + '两日游'
        elif t == '{airline}{fw}':
            kw = random.choice(AIRLINES) + random.choice(FLIGHT_WORDS[:15])
        elif t == '{brand}{hw}':
            kw = random.choice(BRAND_WORDS) + random.choice(HOTEL_WORDS)
        else:
            continue
        keywords.add(kw)

# 打乱顺序
result = list(keywords)
random.shuffle(result)

# 截取2万个
result = result[:20000]

# 写入文件
output_path = 'test_keywords_20k.txt'
with open(output_path, 'w', encoding='utf-8') as f:
    for kw in result:
        f.write(kw + '\n')

print(f"✅ 生成完成: {len(result)} 个关键词")
print(f"   输出文件: {output_path}")

# 统计分类分布预览
categories = {'酒店': 0, '机票': 0, '火车': 0, '景区': 0, '国际': 0, '其他': 0}
hotel_set = set(HOTEL_WORDS + HOTEL_CHAINS)
flight_set = set(FLIGHT_WORDS + FLIGHT_AUX)
train_set = set(TRAIN_WORDS)
scenic_set = set(SCENIC_WORDS)
intl_set = set(INTL_CITIES + COUNTRIES)

for kw in result:
    if any(h in kw for h in hotel_set):
        categories['酒店'] += 1
    elif any(f in kw for f in flight_set):
        categories['机票'] += 1
    elif any(t in kw for t in train_set):
        categories['火车'] += 1
    elif any(s in kw for s in scenic_set):
        categories['景区'] += 1
    elif any(i in kw for i in intl_set):
        categories['国际'] += 1
    else:
        categories['其他'] += 1

print("\n📊 关键词分类分布预估:")
for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
    pct = count / len(result) * 100
    print(f"   {cat}: {count} ({pct:.1f}%)")

# 打印前20个示例
print(f"\n📝 前20个示例:")
for kw in result[:20]:
    print(f"   {kw}")
