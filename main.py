from winsun.database import Query, WeekSale, WeekSold
from winsun.database import ZHUZHAI, BIESHU, SHANGYE, BANGONG
from winsun.shuoli import Zoushi
from winsun.office import PPT, Excel
from winsun.utils import Week


class Report:
    def __init__(self):
        # 初始化ppt, excel, 数据库连接
        self.ppt = PPT()
        self.excel = Excel()
        self.q = Query()

        # 设置日期
        w = Week()
        self.week_str = f'{w.monday.year}第{w.N}周'
        w.str_format('%m.%d')
        self.date_str = f'{w.monday_str}-{w.sunday_str}'

        # 物业类型
        self.usg = {
            '住宅': ZHUZHAI,
            '别墅': BIESHU,
            '商业': SHANGYE,
            '办公': BANGONG
        }

    def intro(self, df, page):
        """导读"""

        # 根据页数找到导读中数据位置(首项4，公差6)
        idx = 4 + 6 * (page - 1)

        # 预处理数据
        value, rate = df.values
        value = list(map(lambda x: round(x, 2), value))
        value[-1] = int(value[-1])
        rate = list(map(lambda x: x.replace('增长', '↗').replace('下降', '↘'), rate))

        # 填入对应位置
        for i, each in enumerate(value + rate):
            self.ppt[[0, idx + i]] = each

    def shuoli(self, df, page):
        """说理"""
        zs = Zoushi(df, degree=0)
        self.ppt[[page, 1]] = f'本周{zs.shuoli}'
        return zs.df

    def market(self, usg_label, page):
        """市场量价"""
        usg = self.usg[usg_label]

        # 走势图
        self.ppt[[page, 2]] = f'南京近10周{usg_label}市场供销量价'
        df_trend = self.q.gxj(usage=usg)
        self.excel[f'{usg_label}量价'] = df_trend

        # 板块图
        self.ppt[[page, 3]] = f'{self.week_str}南京{usg_label}市场分板块供销量价'
        df_plate = self.q.gxj(usage=usg, period=1, by='plate')
        self.excel[f'{usg_label}板块'] = df_plate

        # 通过量价数据生成说理和导读
        df = self.shuoli(df_trend, page)
        self.intro(df, page)

    def rank(self, usg_label, page):
        """排行榜"""

        def adjust(df):
            """调整表格"""
            cols = ['rank', 'plate', 'pop_name', 'type', 'space', 'set', 'price']
            cols_ = ['排名', '板块', '项目', '类型', '面积(㎡)', '套数', '均价(元/㎡)']
            cols_dict = dict(zip(cols, cols_))

            df.loc[:, 'space'] = df['space'].round(0).astype('int')

            # 删除住宅、别墅排行中的“类型”列
            if usg_label in ['住宅', '别墅']:
                cols.remove('type')
            else:
                df['type'] = None

            # 删除上市表中的“均价”列
            if df is df_sale:
                cols.remove('price')
            else:
                df.loc[:, 'price'] = df['price'].round(0).astype('int')

            return df[cols].rename(columns=cols_dict)

        usg = self.usg[usg_label]

        # 上市
        self.ppt[[page, 7]] = f'{self.week_str}{usg_label}市场上市面积前三'
        df_sale = self.q.rank(WeekSale, usage=usg, num=3)
        df_sale = adjust(df_sale)
        self.ppt[[page, 9]] = df_sale

        # 成交
        self.ppt[[page, 8]] = f'{self.week_str}{usg_label}市场成交面积前三'
        df_sold = self.q.rank(usage=usg, num=3)
        df_sold = adjust(df_sold)
        self.ppt[[page, 10]] = df_sold

        # 成交功能
        if usg_label in ['商业', '办公']:
            pop_name = df_sold['项目'].tolist()
            self.excel[f'{usg_label}功能明细'] = self.type_(usg, pop_name)

    def type_(self, usg, pop_name):
        """成交榜中项目功能"""
        w = Week()
        week = f'{w.monday.year}{w.N:02d}'

        res = self.q.query(WeekSold).filter(
            WeekSold.pop_name.in_(pop_name),
            WeekSold.usage.in_(usg),
            WeekSold.week.between(week, week)
        )
        res = self.q.group(res, [WeekSold.pop_name, WeekSold.usage], [WeekSold.space])

        return self.q.to_df(res)


if __name__ == '__main__':
    # PPT('template.pptx').analyze_slides()

    r = Report()

    # 遍历4种物业类型
    for page, usg_label in enumerate(r.usg):
        page += 1
        r.ppt[[page, 0]] = f'{usg_label}市场'
        r.market(usg_label, page)
        r.rank(usg_label, page)

    # 保存
    r.excel.save('data.xlsx')
    r.ppt.save(f'E:/工作文件/报告/周报测试/{r.week_str}({r.date_str}).pptx')
