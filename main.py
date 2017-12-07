from winsun.datebase import Query, ZHUZHAI, SHANGYE, BANGONG, BIESHU, WeekSold, WeekSale
from winsun.shuoli import Zoushi
from winsun.office import PPT, Excel
from winsun.utils import Week


class Report:
    def __init__(self):
        self.ppt = PPT()
        self.excel = Excel()
        self.q = Query()

        w = Week()
        self.week_str = f'{w.monday.year}第{w.N}周'
        w.str_format('%m/%d')
        self.date_str = f'{w.monday_str}-{w.sunday_str}'

        self.usg = {'住宅': ZHUZHAI, '商业': SHANGYE, '办公': BANGONG, '别墅': BIESHU}

    def liangjia_page(self, usage, page):
        print(f'>>> 正在生成{usage}量价页...')
        usg = self.usg[usage]
        df_space = self.q.gxj(usage=usg)
        df_set = self.q.gxj(usage=usg, volumn='set')
        df_plate = self.q.gxj(usage=usg, by='plate')

        zs_space = Zoushi(df_space, degree=0)
        zs_set = Zoushi(df_set, degree=0, volumn='set')

        # 标题
        self.ppt[[page, 0]] = f'{usage}市场-量价'
        self.ppt[[page, 3]] = f'{self.week_str}南京{usage}市场板块供销量价'

        # 说理
        sale_set = zs_set.value_txt('sale')
        sale_set_ = zs_set.thb_txt('sale')
        sale_space = zs_space.value_txt('sale')
        sale_space_ = zs_space.thb_txt('sale')
        sold_set = zs_set.value_txt('sold')
        sold_space = zs_space.value_txt('sold')
        sold_set_ = zs_set.thb_txt('sold').replace('，', '')
        sold_space_ = zs_space.thb_txt('sold').replace('，', '')
        shuoli = f"本周{usage}新增供应{sale_set}{sale_set_}，面积合计{sale_space}{sale_space_}\r" \
                 f"本周{usage}成交{sold_set}，共{sold_space}，套数{sold_set_}，面积{sold_space_}" \
                 f"，{zs_space.text('price')}"
        self.ppt[[page, 4]] = shuoli

        # 图表数据
        columns = ['上市面积（万㎡）', '成交面积（万㎡）', '成交均价（元/㎡）']
        df_space.columns = columns
        self.excel[f'{usage}量价'] = df_space
        df_plate.columns = columns
        self.excel[f'{usage}板块'] = df_plate

    def paihang_page(self, usage, page):
        print(f'>>> 正在生成{usage}排行...')
        usg = self.usg[usage]
        if usage == '住宅':
            page += 1

            df_space = self.q.rank(usage=usg, outputs=['space'], num=10)
            df_space.space = df_space.space.round(2)
            self.ppt[[page, 2]] = f'{self.week_str}（{self.date_str}）成交面积排行榜'
            self.ppt[[page, 5]] = df_space

            df_set = self.q.rank(usage=usg, outputs=['set'], by='set', num=10)
            self.ppt[[page, 3]] = f'{self.week_str}（{self.date_str}）成交套数排行榜'
            self.ppt[[page, 6]] = df_set

        elif usage == '别墅':
            pass
        else:
            df_space = self.q.rank(usage=usg, num=3)
            prjs = [f'{x.plate}{x.pop_name}（{x.space/1e4:.2f}万㎡，{x.set}套，{x.price:.0f}元/㎡）' for _, x in df_space.iterrows()]
            self.ppt[[page, 4]] = f"{self.ppt[[page,4]].text}\r本周成交面积榜单前三：{'、'.join(prjs)}"

    def shangshi(self, usage):
        usg = self.usg[usage]
        df = self.q.rank(WeekSale, usage=usg)
        self.excel[f'{usage}上市明细'] = df

    def each_usage(self, usage, page):
        self.liangjia_page(usage, page)
        self.paihang_page(usage, page)
        self.shangshi(usage)


if __name__ == '__main__':
    r = Report()
    for i, usage in enumerate(['住宅', '商业', '办公', '别墅']):
        page = i if usage == '住宅' else i + 1
        r.each_usage(usage, page)
    r.excel.save()
    r.ppt.save()
