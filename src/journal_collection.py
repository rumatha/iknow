from journal import Journal

#===================================================================================================

bibliosfera =                               Journal('Библиосфера')
computational_and_theor_chemistry =         Journal('Computational and Theoretical Chemistry')
current_organic_chemistry =                 Journal('Current Organic Chemistry')
georesursy =                                Journal('Георесурсы')
inf_tech =                                  Journal('Информационные технологии',             '1684-6400', '',          '4', 'К1', '', '',   '', '',     '',                     '', '', '', '', '', '+', '+', '+', 'http://novtex.ru/IT/', )
inzhenerno_phis_journal =                   Journal('Инженерно-физический журнал')
informazionnye_processy =                   Journal('Информационные процессы')
informazionnye_tehnologii =                 Journal('Информационные технологии')
informazionnye_tehnologii_i_vych_systemy =  Journal('Информационные технологии и вычислительные системы')
journal_of_coordination_chemistry =         Journal('Journal of Coordination Checmistry')
journal_of_eng_phis_and_thermophis =        Journal('Journal of Engineering Physiscs and Thermophysiscs')
journal_of_porphyrins_and_phthalocyanines = Journal('Journal of Porphyrins and Phthalocyanines')
lncs =                                      Journal('Lecture Notes in Computer Science',     '0302-9743', '1611-3349', '',  '',   '', 'Q2', '', '',     '',                     '', '', '', '', '', '+', '+', '',  'https://www.springer.com/de/it-informatik/lncs')
lobachevskii =                              Journal('Lobachevskii Journal of Mathematics')
mathematical_modeling =                     Journal('Математическое Моделирование',          '0234-0879', '',          '3', 'К~', '', '',   '', '3595', 'se:00002821|00007173', '', '', '', '', '', '+', '+', '+', 'https://www.mathnet.ru/php/journal.phtml?jrnid=mm&option_lang=rus')
math_models_and_computer_simulations =      Journal('Mathematical Models and Computer Simulations')
nauchnye_i_tehnich_biblioteki =             Journal('Научные и технические библиотеки')
polyhedron =                                Journal('Polyhedron')
prog_syst_teor_i_prilozheniya =             Journal('Программные системы: Теория и приложения')
reviews_and_advances_in_chemistry =         Journal('Reviews and Advances in Chemistry')
sci_and_tech_inf_processing =               Journal('Scientific and Technical Information Processing')
sistemy_i_sredstva_informazii =             Journal('Системы и средства информации')
software_and_systems =                      Journal('Программные продукты и системы',        '0236-235X', '2311-2735', '4', 'К1', '', '',   '', '',     '',                     '', '', '', '', '', '+', '+', '+', 'https://swsys.ru')
tcomm =                                     Journal('T-Comm - Телекоммуникации и транспорт', '2072-8735', '2072-8743', '',  'К1', '', '',   '', '',     '',                     '', '', '', '', '', '',  '+', '+', 'https://media-publisher.ru/abouttcomm/')
uch_zapiski_kaz_univer_estestv =            Journal('Ученые записки Казанского университета. Серия Естественные науки.')
uch_zapiski_kaz_univer_phis_mat =           Journal('Ученые записки Казанского университета. Серия Физико-математические науки.')
vestnik_technologich_universiteta =         Journal('Вестник технологического университета')
electrosvyaz =                              Journal('Электросвязь')

all = \
[
    mathematical_modeling, lncs, tcomm
]

#===================================================================================================

if __name__ == '__main__':
    for j in all:
        print(j)

#===================================================================================================
