from winsun.datebase import Query, ZHUZHAI, SHANGYE, BANGONG, BIESHU, WeekSold, WeekSale
from winsun.shuoli import Zoushi
from winsun.office import PPT, Excel
from winsun.utils import Week
from itertools import product


class Report:
    def __init__(self):
        self.ppt = PPT()
        self.excel = Excel()
        self.q = Query()

        w = Week()
        self.week_str = f'{w.monday.year}第{w.N}周'
        w.str_format('%m.%d')
        self.date_str = f'{w.monday_str}-{w.sunday_str}'

        self.usg = {'住宅': ZHUZHAI, '商业': SHANGYE, '办公': BANGONG, '别墅': BIESHU}

    def liangjia_page(self, usage, page):
        def daodu():
            # 　根据物业类型所在页数计算出是第几种物业类型， 再x10
            i = (page - 1) / 2 * 10

            # 替换两张表中增长、下降为箭头
            dfs = zs_set.df, zs_space.df
            for df, col in product(dfs, ['sale', 'sold', 'price']):
                df.at['h', col] = df.at['h', col].replace('下降', '↘').replace('增长', '↗')

            # 在对应shape.text中填入数据
            for j, (col, df, idx) in enumerate(product(['sale', 'sold'], dfs, ['v', 'h'])):
                value = int(df.at[idx, col]) if idx == 'v' and df is zs_set.df else df.at[idx, col]
                self.ppt[[0, int(i + j + 4)]] = value
            for j, idx in enumerate(['v', 'h']):
                self.ppt[[0, int(i + j + 12)]] = int(dfs[0].at[idx, 'price']) if idx=='v' else dfs[0].at[idx, 'price']

        print(f'>>> 正在生成{usage}量价页...')

        usg = self.usg[usage]
        df_space = self.q.gxj(usage=usg)
        df_set = self.q.gxj(usage=usg, volumn='set')
        df_plate = self.q.gxj(usage=usg, by='plate', period=1)

        zs_space = Zoushi(df_space, degree=0)
        zs_set = Zoushi(df_set, degree=0, volumn='set')

        # 标题
        self.ppt[[page, 3]] = f'{self.week_str}南京{usage}市场板块供销量价'

        # 说理
        sale_set = zs_set.text('sale')
        sale_space = zs_space.text('sale')
        sold_set = zs_set.text('sold')
        sold_space = zs_space.text('sold')
        price = zs_space.text('price')
        self.ppt[[page, 4]] = f'本周{usage}市场{sale_set}{sale_space}\r' \
                              f'{sold_set}{sold_space}\r' \
                              f'{price}'

        # 图表数据
        columns = ['上市面积（万㎡）', '成交面积（万㎡）', '成交均价（元/㎡）']
        df_space.columns = columns
        self.excel[f'{usage}量价'] = df_space
        df_plate.columns = columns
        self.excel[f'{usage}板块'] = df_plate

        # 导读
        daodu()

    def paihang_page(self, usage, page):
        print(f'>>> 正在生成{usage}排行...')
        usg = self.usg[usage]
        if usage == '住宅':
            df_space = self.q.rank(usage=usg, outputs=['space'], num=10)
            df_space.space = df_space.space.round(0).astype('int')
            df_space.columns = ['排名', '项目推广名', '板块', '面积(㎡)']
            self.ppt[[page, 2]] = f'{self.week_str}（{self.date_str}）成交面积排行榜'
            self.ppt[[page, 5]] = df_space

            df_set = self.q.rank(usage=usg, outputs=['set'], by='set', num=10)
            df_set.columns = ['排名', '项目推广名', '板块', '套数']
            self.ppt[[page, 3]] = f'{self.week_str}（{self.date_str}）成交套数排行榜'
            self.ppt[[page, 6]] = df_set

        elif usage == '别墅':
            df_space = self.q.rank(usage=usg, outputs=['space'], num=5)
            df_space.space = df_space.space.round(0).astype('int')
            df_space.columns = ['排名', '项目推广名', '板块', '面积(㎡)']
            self.ppt[[page, 2]] = f'{self.week_str}（{self.date_str}）成交面积排行榜'
            self.ppt[[page, 5]] = df_space

            df_set = self.q.rank(usage=usg, outputs=['set'], by='set', num=5)
            df_set.columns = ['排名', '项目推广名', '板块', '套数']
            self.ppt[[page, 3]] = f'{self.week_str}（{self.date_str}）成交套数排行榜'
            self.ppt[[page, 6]] = df_set

        else:
            cols = ['rank', 'plate', 'pop_name', 'type', 'space', 'set', 'price']
            cols_ = ['排名', '板块', '项目', '类型', '面积(万㎡)', '套数', '均价(元/㎡)']
            cols_dict = dict(zip(cols, cols_))

            df_sale = self.q.rank(WeekSale, usage=usg, num=3)
            df_sold = self.q.rank(usage=usg, num=3)
            for each in [df_sale, df_sold]:
                each['type'] = None
                each['space'] = each['space'].round(0).astype('int')
            df_sale = df_sale[cols].drop('price', axis=1).rename(columns=cols_dict)
            df_sold['price'] = df_sold['price'].round(0).astype('int')
            df_sold = df_sold[cols].rename(columns=cols_dict)

            self.ppt[[page, 2]] = f'{self.week_str}（{self.date_str}）{usage}上市面积前三项目'
            self.ppt[[page, 5]] = df_sale
            self.ppt[[page, 3]] = f'{self.week_str}（{self.date_str}）{usage}成交面积前三项目'
            self.ppt[[page, 6]] = df_sold

    def shangshi(self, usage):
        usg = self.usg[usage]
        df = self.q.rank(WeekSale, usage=usg)
        self.excel[f'{usage}上市明细'] = df


if __name__ == '__main__':
    r = Report()
    page = 0
    for i, usage in enumerate(['住宅', '别墅', '商业', '办公']):
        page = i * 2 + 1
        r.liangjia_page(usage, page)
        page += 1
        r.paihang_page(usage, page)
        r.shangshi(usage)
    r.excel.save()
    r.ppt.save(f'E:/工作文件/报告/周报测试/{r.week_str}({r.date_str}).pptx')
