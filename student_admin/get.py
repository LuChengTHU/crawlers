import requests
from bs4 import BeautifulSoup
import json
import xlsxwriter

showLength = input('请输入获取数量(大于等于1): ')
username = input('info用户名: ')
password = input('密码: ')

print('正在制作社团类别对应表...')
classification = {
    '文化类': [
        '国学经典文化传播协会'
        '葡萄酒协会',
        '中医学社',
        '文学社',
        '文苑沙龙协会',
        '清莲诗社',
        '国学社',
        '红楼梦协会',
        '茶文化协会',
        '致知协会',
        '禅文化研究社',
        '紫苑学会',
        '古典文化研究协会',
        '咖啡文化交流协会',
        '科幻协会',
        '礼射研习协会',
        '好读书协会',
        '推理协会',
        '西南联合大学精神研究会',
        '走向全球协会',
        '海峡两岸交流协会',
        '对外交流协会',
        '国际文化交流协会',
        '国际事务交流协会',
        '中日友好交流协会',
        '中非交流协会',
        '中韩文化交流协会',
        '中法交流协会',
        '中巴文化交流协会',
        '蒙文化传播交流协会',
        '南粤文化交流协会',
        '黔文化交流协会',
        '赣文化交流协会',
        '闽文化交流协会',
        '湖湘文化交流协会',
        '荆楚文化发展研究会',
        '三晋文化交流协会',
        '八桂文化交流协会',
        '徽文化交流协会',
        '川渝文化发展研究会',
        '中原发展研究会',
        '苗文化研究会',
        '辽沈文化发展研究会',
        '齐鲁文化交流协会',
        '海派文化交流协会',
        '江苏文化交流协会',
        '滇文化交流协会',
        '津门文化交流协会',
        '越文化交流协会',
        '中德文化交流协会',
        '青海文化交流协会',
        '长吉文化发展研究会',
        '塞上江南文化交流协会',
        '陇文化交流协会',
        '燕赵文化交流协会'
    ],
    '公益类': [
        '科技教育交流协会',
        '绿色协会',
        '手语社',
        '教育扶贫公益协会',
        '爱心公益协会',
        '非物质文化遗产传播与保护协会',
        '粉刷匠工作室协会',
        '书脊支教团',
        '心理协会',
        '关注城市劳动者协会',
        '关注女性发展协会',
        '学工同行志愿者协会',
        '公益学术促进会',
        '教育互联网公益协会',
        '唐仲英爱心社',
        '小动物保护协会',
        '清源协会',
        '治安服务队',
        '思源社',
        '雁行社',
        '素食协会',
        '新能源微电网协会',
        '法律援助协会',
        '关爱留守儿童协会',
        '健康促进公益协会',
        '文博协会',
        '公益建造协会',
        '无障碍发展研究协会',
        '屋顶农场协会',
        '“暖·爱”协会',
        '益创咨询社',
        '曾宪备公益创新交流协会',
        '国际公益设计协会'
    ],
    '艺术类': [
        '舞蹈协会',
        '阿卡贝拉清唱社',
        '街舞社',
        '音乐剧社',
        '魔术协会',
        '笛子协会',
        '次世代动漫社',
        '电影协会',
        '吉他协会',
        '摄影协会',
        '书法协会',
        '采薇插花艺术协会',
        '古典爱乐社',
        '口琴社',
        '越剧协会',
        '古琴社',
        '腰鼓协会',
        '陶瓷体验协会',
        '纸艺社',
        'DIY协会',
        '手风琴协会',
        '戏剧创作社',
        '华语音乐交流协会',
        '中国音乐社',
        '京昆协会',
        '插画协会',
        '美妆社',
        '影视创作协会',
        '说唱社'
    ],
    '体育类': [
        '帆船协会',
        '跆拳道协会',
        '羽毛球协会',
        '游泳协会',
        '乒乓球协会',
        '网球协会',
        '高尔夫球协会',
        '滑雪协会',
        '艺术体操协会',
        '排球协会',
        '马拉松协会',
        '跳水协会',
        '击剑协会',
        '自行车协会',
        '滑板协会',
        '棒垒球协会',
        '篮球协会',
        '空手道协会',
        '腰旗橄榄球队',
        '绿茵协会',
        '弓箭协会',
        '马术协会',
        '剑道协会',
        '花样滑冰协会',
        '定向越野协会',
        '山野协会',
        '轮滑协会',
        '毽球协会',
        '健美操协会',
        '板球协会',
        '吴式太极拳协会',
        '杨式太极拳协会',
        '陈式太极拳协会',
        '四国军棋协会',
        '中国象棋协会',
        '飞盘协会',
        '围棋协会',
        '健美协会',
        '武术协会',
        '桥牌协会',
        '紫光国际象棋协会',
        '合气道协会',
        '台球协会',
        '桌游社',
        '跳绳协会',
        '瑜伽协会',
        '体育养生社',
        '冬泳协会',
        '混合健身社',
        '潜水协会'
    ],
    '科创类': [
        '创客空间协会',
        'STM32嵌入式协会',
        '未来通信兴趣团队',
        '未来航空兴趣团队',
        '未来汽车兴趣团队',
        '未来智能机器人兴趣团队',
        '未来城市与新能源兴趣团队',
        '未来动漫兴趣团队',
        '未来中医药兴趣团队',
        '未来医疗兴趣团队',
        '未来云计算兴趣团队',
        '未来互联网兴趣团队',
        '未来人居兴趣团队',
        '未来新媒体兴趣团队',
        '未来数字校园兴趣团队',
        '未来石墨烯兴趣团队',
        '创业协会',
        '未来企业家协会',
        '新能源协会',
        '超级计算机协会',
        '网络与开源软件协会',
        '技术娱乐设计协会（TEDxTHU）',
        '汽车行业研究会',
        'Lab μ校园极客社',
        '信息化服务与咨询协会',
        '大数据研究协会',
        '物联网协会',
        '网络安全技术协会',
        '天文协会',
        '微创造协会',
        '脑科学协会',
        '虚拟现实研习社',
        '区块链协会',
        '科学传播协会',
        '语言学协会',
        '碳立方创新研究协会'
    ],
    '素质拓展类': [
        '马克思主义学习研究协会',
        '实体行业研究协会',
        '三农问题学习研究会',
        '县域发展研究会',
        '城市中国研究会',
        '健康产业研究会',
        '经济学会',
        '咨询协会',
        '金融领导力协会',
        '翻译协会',
        '会计协会',
        '金融协会',
        '理财协会',
        '互联网金融协会',
        '信息管理协会',
        '动物协会',
        '在线教育协会',
        '植物协会',
        '保险与精算协会',
        '项目管理协会',
        '政治经济学与现代资本主义研究会',
        '金融数据与量化投资协会',
        '大数据管理及商业创新协会',
        '债券研究协会',
        '校园营造社',
        '风景园林协会',
        '商业智慧设计协会',
        '基层公共部门发展研究会',
        '职业发展协会',
        '求是学会',
        '学业发展协会',
        '国旗仪仗队',
        '军事爱好者协会',
        '模拟亚太经合组织协会',
        '就业服务协会',
        '记者团',
        '辩论协会',
        '演讲与口才协会',
        '英语辩论协会',
        '时政研究会',
        '科技经纪人协会',
        '县域经济研究会',
        '全球治理与国际组织发展协会',
        '交通文化交流协会',
        '军事特训队',
        '企业文化协会',
        '体育管理与产业发展研究会',
        '校友交流协会',
        '英雄文化协会'
    ]
}

def handleName(name):
    return name.replace('交流协会', '').replace('协会', '').replace('研究会', '').replace('社团', '').replace('社', '').replace('清华大学学生', '').replace('清华大学', '').replace('学生', '').replace('清华', '')

class_dict = {}
for key in classification.keys():
    for name in classification[key]:
        class_dict[handleName(name)] = key

sess = requests.session()
print('正在登录...')
try:
    login_req = sess.post(
        'https://id.tsinghua.edu.cn/do/off/ui/auth/login/post/c5d2f775bf1acacfb4f9b277ebe3604e/2?/j_spring_security_thauth_roaming_entry',
        data={
            'atOnce': 'true',
            'i_user': username,
            'i_pass': password
        }
    )
    sess.get(BeautifulSoup(login_req.text, 'html.parser').a.get('href'))
except Exception:
    print('网络错误,请重试')
    exit()
print('登录成功,正在请求数据...')
data = [
    {"name": "sEcho", "value": '1'},
    {"name": "iColumns", "value": '10'},
    {"name": "sColumns", "value": ""},
    {"name": "iDisplayStart", "value": '0'},
    {"name": "iDisplayLength", "value": showLength},
    {"name": "mDataProp_0", "value": "hdbh"},
    {"name": "mDataProp_1", "value": "hddlmc"},
    {"name": "mDataProp_2", "value": "hdlxmc"},
    {"name": "mDataProp_3", "value": "function"},
    {"name": "mDataProp_4", "value": "hdcyrs"},
    {"name": "mDataProp_5", "value": "hdrq"},
    {"name": "mDataProp_6", "value": "sqsj"},
    {"name": "mDataProp_7", "value": "yjspjg"},
    {"name": "mDataProp_8", "value": "ejspjg"},
    {"name": "mDataProp_9", "value": "function"},
    {"name": "iSortCol_0", "value": '6'},
    {"name": "sSortDir_0", "value": "desc"},
    {"name": "iSortingCols", "value": '1'},
    {"name": "bSortable_0", "value": 'false'},
    {"name": "bSortable_1", "value": 'false'},
    {"name": "bSortable_2", "value": 'false'},
    {"name": "bSortable_3", "value": 'false'},
    {"name": "bSortable_4", "value": 'false'},
    {"name": "bSortable_5", "value": 'false'},
    {"name": "bSortable_6", "value": 'false'},
    {"name": "bSortable_7", "value": 'false'},
    {"name": "bSortable_8", "value": 'false'},
    {"name": "bSortable_9", "value": 'false'}
]

try:
    one_page = sess.post(
        'http://oa.student.tsinghua.edu.cn/b/twbgzx/bmgly/yspxshd',
        data='aoData=' + json.dumps(data),
        headers={
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': 'oa.student.tsinghua.edu.cn',
            'Origin': 'http://oa.student.tsinghua.edu.cn',
            'Pragma': 'no-cache',
            'Referer': 'http://oa.student.tsinghua.edu.cn/f/twbgzx/bmgly/yspxshd',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
        }
    )
except Exception:
    print('网络错误,两次尝试间隔时间不要太短.')
    exit()
try:
    resList = json.loads(one_page.text)['object']['aaData']
except Exception:
    print('密码错误!')
    exit()
print('正在生成excel表格...')
book = xlsxwriter.Workbook('result.xlsx')
sheet = book.add_worksheet()
cell = book.add_format({'font_name': '楷体', 'align': 'center', 'font_size': 22, 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
sheet.write(0, 0, '社团类别')
sheet.set_column(0, 0, 20, cell)
sheet.write(0, 1, '社团名称')
sheet.set_column(1, 1, 80, cell)
sheet.write(0, 2, '活动日期')
sheet.set_column(2, 2, 30, cell)
sheet.write(0, 3, '活动具体时间段')
sheet.set_column(3, 3, 40, cell)
sheet.write(0, 4, '活动说明')
sheet.set_column(4, 4, 100, cell)
sheet.write(0, 5, '资源类型')
sheet.set_column(5, 5, 30, cell)
sheet.write(0, 6, '审批结果')
sheet.set_column(6, 6, 20, cell)
sheet.write(0, 7, '是否涉校外')
sheet.set_column(7, 7, 20, cell)
sheet.write(0, 8, '是否涉境外')
sheet.set_column(8, 8, 20, cell)
sheet.write(0, 9, '备注')
sheet.set_column(9, 9, 30, cell)
sheet.write(0, 10, '活动申请时间')
sheet.set_column(10, 10, 100, cell)

for i, res in enumerate(resList):
    if handleName(res['zbfmc']) in class_dict.keys():
        sheet.write(i + 1, 0, class_dict[handleName(res['zbfmc'])])
    sheet.write(i + 1, 1, res['zbfmc'])
    sheet.write(i + 1, 2, res['hdrq'][:4] + '/' + res['hdrq'][4:6] + '/' + res['hdrq'][6:])
    sheet.write(i + 1, 3, res['hdsj'])
    sheet.write(i + 1, 4, res['hdzt'])
    resources = []
    resources.append('教室' if res['sfsqjs'] == '是' else '')
    resources.append('展板' if res['sfsqzb'] == '是' else '')
    resources.append('电子屏' if res['sfsqdzp'] == '是' else '')
    resources.append('C楼活动室' if res['sfsqclhds'] == '是' else '')
    resources.append('室外场地' if res['sfsqswcd'] == '是' else '')
    resources = '、'.join(list(filter(None, resources)))
    sheet.write(i + 1, 5, resources)
    sheet.write(i + 1, 6, res['yjspjg'])
    sheet.write(i + 1, 7, res['sfsxw'])
    sheet.write(i + 1, 8, res['sfsjw'])
    sheet.write(i + 1, 10, res['sqsj'])

book.close()
print('成功写入到result.xlsx文件中！')
