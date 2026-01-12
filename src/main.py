import worksheet_private_collection as wc
import temporary_team_private_collection as ttc
import complex_theme_private_collection as cx
from generator_word import GeneratorWord as gw
from generator_excel import GeneratorExcel as ge
import publication_collection as pc
import generator_html as gh
import generator_word
import outlay_tree

#===================================================================================================

def print_worksheets_statitics():
    """
    Print worksheet statistics.
    """

    osspv_all = wc.osspv.slots_sum()
    osspv_occupied = wc.osspv.occupied_slots_sum()
    kvvk_ovk_all = wc.kvvk_ovk.slots_sum()
    kvvk_ovk_occupied = wc.kvvk_ovk.occupied_slots_sum()

    print(f'Статистика по ставкам:')
    print(f'\tОССиПВ : занято {osspv_occupied} из {osspv_all}'
          f', ОВК : занято {kvvk_ovk_occupied} из {kvvk_ovk_all}'
          f', всего : {osspv_occupied + kvvk_ovk_occupied}')

#---------------------------------------------------------------------------------------------------

def print_temporary_teams_statistics():
    """
    Print temporary teams statistics.
    """

    cx1_slots = ttc.cx1.slots_sum()
    cx2_slots = ttc.cx2.slots_sum()
    oth_slots = ttc.other.slots_sum()

    print(f'Статистика по временному трудовому коллективу:')
    print(f'\t6Ф-СИ.1 : {cx1_slots}, 6Ф-СИ.2 : {cx2_slots}, другие : {oth_slots}'
          f', всего : {cx1_slots + cx2_slots + oth_slots}')

#---------------------------------------------------------------------------------------------------

def generate_documents_pack(y, out_dir):
    """
    Generate documents pack.

    Parameters
    ----------
    y : int
        Start year.
    out_dir : str
        Out directory.
    """

    # Walk all themes.
    for (theme, team) in [(cx.cx1, ttc.cx1), (cx.cx2, ttc.cx2)]:
        pre = f'{out_dir}/{theme.short_title}'

        # We need to create outlays for thematics.
        # NB! Thematic outlays we calculate only for 'y' year.

        # Get all three thematics.
        [th1, th2, th3] = theme.thematics

        # Get all thematic weights.
        w1, w2, w3 = \
            0.01 * th1.funding_part(y), 0.01 * th2.funding_part(y), 0.01 * th3.funding_part(y)

        # Split.
        th1.outlay = \
            outlay_tree.duplicate_outlay_each_line(theme.outlay,
                                                   'тематике исследований',
                                                   w1, w1,
                                                   1.0, 1.0, 1.0,
                                                   1.0, 1.0, 1.0,
                                                   1.0, 1.0, 1.0,
                                                   1.0, 1.0,
                                                   1.0, 1.0,
                                                   1.0, 1.0, 1.0,
                                                   w1, w1, w1, w1, w1, w1, w1, w1, w1)
        th2.outlay = \
            outlay_tree.duplicate_outlay_each_line(theme.outlay,
                                                   'тематике исследований',
                                                   w2, w2,
                                                   0.0, 0.0, 0.0,
                                                   0.0, 0.0, 0.0,
                                                   0.0, 0.0, 0.0,
                                                   0.0, 0.0,
                                                   0.0, 0.0,
                                                   0.0, 0.0, 0.0,
                                                   w2, w2, w2, w2, w2, w2, w2, w2, w2)

        th3.outlay = \
            outlay_tree.duplicate_outlay_each_line(theme.outlay,
                                                   'тематике исследований',
                                                   w3, w3,
                                                   0.0, 0.0, 0.0,
                                                   0.0, 0.0, 0.0,
                                                   0.0, 0.0, 0.0,
                                                   0.0, 0.0,
                                                   0.0, 0.0,
                                                   0.0, 0.0, 0.0,
                                                   w3, w3, w3, w3, w3, w3, w3, w3, w3)

        # Form gos assignment (order 3188, form, supplements 7 - 11).
        prex = f'{pre}-{y}-{y + 2}-формирование'
        prexf = f'{prex}-форма'
        gw.generate_form_gos_assignment_3188_07_technical_task(theme, y, f'{prexf}-07-ТЗ')
        gw.generate_form_gos_assignment_3188_08_calendar_plan(theme, y, f'{prexf}-08-КП')
        gw.generate_form_gos_assignment_3188_09_outlay(theme, y, f'{prexf}-09-смета')
        gw.generate_form_gos_assignment_3188_10_team(theme, team, y, f'{prexf}-10-ВТК')
        gw.generate_form_gos_assignment_3188_11_equipment(theme, f'{prexf}-11-оборудование')
        ge.generate_PTNI_researchers(theme, f'{prex}-ПТНИ-исследователи')

        # Exec gos assignment (order 3188, exec, supplements 1 - 5).
        prex = f'{pre}-{y}-{y + 2}-приказ'
        prexp = f'{prex}-приложение'
        gw.generate_exec_gos_assignment_3188_01_technical_task(theme, y, f'{prexp}-1-ТЗ')
        gw.generate_exec_gos_assignment_3188_02_calendar_plan(theme, y, f'{prexp}-2-КП')
        gw.generate_exec_gos_assignment_3188_03_outlay(theme, y, f'{prexp}-3-смета')
        gw.generate_exec_gos_assignment_3188_04_team(theme, team, y, f'{prexp}-4-ВТК')
        gw.generate_exec_gos_assignment_3188_05_equipment(theme, f'{prexp}-5-оборудование')
        gw.generate_exec_gos_assignment_3188_order(theme, y, prex)

#===================================================================================================

if __name__ == '__main__':

    # plans
    gh.generate_publications_info(pc.nrcki_2025, '../out/publications_2025.html')
    gh.generate_plan(cx.cx1, '../out/plan_6f_si_1.html', year_from=2026, year_to=2029)
    gh.generate_plan(cx.cx2, '../out/plan_6f_si_2.html', year_from=2026, year_to=2029)

    # generate documents for complex themes
    #generate_documents_pack(2027, '../out/docs')

#===================================================================================================
