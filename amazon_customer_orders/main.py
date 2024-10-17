import pandas as pd
from pathlib import Path
from tab_stats_cb import stats_cb
from tab_glance_cb import glance_views
from tab_TEMPLATE import top_titles

df_stats_cb = stats_cb()
df_glance_cb = glance_views()

df_top20_fl_cb = top_titles(publisher="Chronicle", flbl="FL", num_rows=20)
df_top20_fl_dp = top_titles(publisher="!Chronicle", flbl="FL", num_rows=20)
df_top30_bl_cb = top_titles(publisher="Chronicle", flbl="BL", num_rows=30)

df_sc = top_titles(publisher="Sierra Club", flbl=None, num_rows=5)
df_ga = top_titles(publisher="Galison", flbl=None, num_rows=5)
df_lk = top_titles(publisher="Laurence King", flbl=None, num_rows=5)
df_hg = top_titles(publisher="Hardie Grant Publishing", flbl=None, num_rows=5)
df_lq = top_titles(publisher="Levine Querido", flbl=None, num_rows=5)
df_tw = top_titles(publisher="Tourbillon", flbl=None, num_rows=5)
df_do = top_titles(publisher="Do Books", flbl=None, num_rows=5)
df_cc = top_titles(publisher="Creative Company", flbl=None, num_rows=5)
df_pb = top_titles(publisher="Paperblanks", flbl=None, num_rows=5)

path = Path(fr'G:\SALES\Amazon\RBH\weekly_customer_order\atelier\amazon_weekly_customer_order_py.xlsx')

try:
    with pd.ExcelWriter(path,engine='xlsxwriter') as writer:
        df_stats_cb.to_excel(writer, sheet_name='stats_cb', index=False)
        df_glance_cb.to_excel(writer, sheet_name='glance_cb', index=False)
        df_top20_fl_cb.to_excel(writer, sheet_name='top20_fl_cb', index=False)
        df_top20_fl_dp.to_excel(writer, sheet_name='top20_fl_dp', index=False)
        df_top30_bl_cb.to_excel(writer, sheet_name='top30_cb_bl', index=False)
        df_sc.to_excel(writer, sheet_name='Sierra Club', index=False)
        df_ga.to_excel(writer, sheet_name='Galison', index=False)
        df_lk.to_excel(writer, sheet_name='Laurence_King', index=False)
        df_hg.to_excel(writer, sheet_name='Hardie Grant Publishing', index=False)
        df_lq.to_excel(writer, sheet_name='Levine Querido', index=False)
        df_tw.to_excel(writer, sheet_name='Tourbillon', index=False)
        df_do.to_excel(writer, sheet_name='Do Books', index=False)
        df_cc.to_excel(writer, sheet_name='Creative Company', index=False)
        df_pb.to_excel(writer, sheet_name='Paperblanks', index=False)
except Exception as e:
    print(e)