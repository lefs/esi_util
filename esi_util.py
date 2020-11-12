import sys
import datetime
import json
from pathlib import Path
import dateutil.relativedelta
import fire
import pygal
import pandas as pd
from terminaltables import SingleTable


class ESIDataWrapper:
    """Convenience class for the background work related to the ESI xlsx."""
    ENTITY_CODES = [
        'eu',
        'ea',
        'at',
        'be',
        'dk',
        'de',
        'el',
        'es',
        'fr',
        'it',
        'nl',
        'pl',
        'pt',
        'fi',
        'se',
        'uk'
    ]
    # Two digit country/entity codes (as used in the ESI) mapped to
    # countries/entities.
    ESI_ENTITIES = {
        'eu': 'Europe',
        'ea': 'Euro Area',
        'at': 'Austria',
        'be': 'Belgium',
        'dk': 'Denmark',
        'de': 'Germany',
        'el': 'Greece',
        'es': 'Spain',
        'fr': 'France',
        'it': 'Italy',
        'nl': 'Netherlands',
        'pl': 'Poland',
        'pt': 'Portugal',
        'fi': 'Finland',
        'se': 'Sweden',
        'uk': 'United Kingdom'
    }
    # The columns corresponding to each entity in the ESI xlsx file.
    ENTITY_COLS = dict(
        eu='A,C:H',
        ea='A,K:P',
        at='A,FO:FT',
        be='A,S:X',
        dk='A,AQ:AV',
        de='A,AY:BD',
        el='A,BW:CB',
        es='A,CE:CJ',
        fr='A,CM:CR',
        it='A,DC:DH',
        nl='A,FG:FL',
        pl='A,FW:GB',
        pt='A,GE:GJ',
        fi='A,HK:HP',
        se='A,HS:HX',
        uk='A,IA:IF'
    )
    # These are used in the ESI xlsx as column headers.
    ESI_COMPONENTS = [
        '.INDU',  # <Entity Code>.INDU
        '.SERV',  # <Entity Code>.SERV
        '.CONS',  # <Entity Code>.CONS
        '.RETA',  # <Entity Code>.RETA
        '.BUIL',  # <Entity Code>.BUIL
        '.ESI'  # <Entity Code>.ESI
    ]

    def __init__(self, data_dir=Path('.'),
                 esi_filename='main_indicators_nace2.xlsx',
                 esi_sheet_name='MONTHLY'):
        self.data_dir = data_dir
        self.esi_filename = esi_filename
        self.esi_sheet_name = esi_sheet_name

    def _entity_csv_filename(self, entity_code):
        """Constructs a path to an entity's CSV file."""
        return Path(self.data_dir) / '{}_esi.csv'.format(entity_code)

    def _create_esi_csv_tables(self, esi_tables):
        """Creates a CSV file for each country/entity."""
        for code in self.ENTITY_CODES:
            # Each value of the esi_tables dict is a Pandas DataFrame.
            esi_tables[code].to_csv(
                self._entity_csv_filename(code), encoding='utf-8'
            )

    def _import_esi_tables_from_xlsx(self):
        """Imports the ESI numbers for each country/entity we're interested
        in into DataFrames.
        """
        # ESI xlsx file and relevant sheet.
        esi_file_path = Path(self.data_dir) / self.esi_filename
        # This will hold a DataFrame for each country's/entity's ESI numbers.
        esi_tables = {}

        for entity, cols in self.ENTITY_COLS.items():
            esi_tables[entity] = pd.read_excel(
                esi_file_path,
                sheet_name=self.esi_sheet_name,
                header=0,
                index_col=0,
                usecols=cols
            )

        return esi_tables

    def _load_esi_tables_from_csv(self):
        esi_tables = {}
        for ec in self.ENTITY_CODES:
            esi_tables[ec] = pd.read_csv(
                self._entity_csv_filename(ec), index_col=0, parse_dates=[0]
            )

        return esi_tables

    def _fetch_esi_tables(self):
        """Returns a dict where each key is an entity code and its
        corresponding value is a pandas DataFrame with the ESI measurements
        for this entity.
        """
        # Check if we have CSV files for the ESI tables.
        we_have_csvs = True
        for ec in self.ENTITY_CODES:
            if not Path(self._entity_csv_filename(ec)).is_file():
                we_have_csvs = False
                break

        if we_have_csvs:
            esi_tables = self._load_esi_tables_from_csv()
        else:
            esi_tables = self._import_esi_tables_from_xlsx()
            self._create_esi_csv_tables(esi_tables)

        # Convert date indices to monthly frequency.
        for ec in self.ENTITY_CODES:
            esi_tables[ec].index = esi_tables[ec].index.to_period(freq='M')

        return esi_tables

    def get_latest_rankings(self, date=None):
        """Returns a dict where keys are the ESI components and their values
        are lists with country measurements. For example:

        {'construction_confidence': [('Netherlands', 4.2),
                                     ('Germany', 1.8),
                                     ('Austria', 0.6),
                                     # ...
                                     ('Sweden', -28.7),
                                     ('Greece', -52.1),
                                     ('United Kingdom', nan)],
         'consumer_confidence': [('Sweden', 2.7),
                                 ('Denmark', 0.5),
                                 # ...
                                 ('Finland', -5.9),
                                 ('Portugal', -28.6),
                                 ('Greece', -41.0)],
         'esi': [('France', 96.6),
                 ('Germany', 95.5),
                 # ...
                 ('Denmark', 80.6),
                 ('Poland', 77.9)],
         'industrial_confidence': [('Sweden', -1.4),
                                   ('Netherlands', -9.8),
                                   ('Greece', -18.1),
                                   # ...
                                   ('Poland', -20.3),
                                   ('United Kingdom', -21.5),
                                   ('Finland', -21.7)],
         'retail_confidence': [('Sweden', 11.8),
                               ('Denmark', 11.2),
                               # ...
                               ('Netherlands', 6.2),
                               ('Spain', -24.6)],
         'services_confidence': [('Germany', 0.5),
                                 ('Austria', -5.2),
                                 # ...
                                 ('Spain', -35.8),
                                 ('United Kingdom', -36.7)]}
        """

        INDU_idx = 0  # INDU - Industrial confidence indicator (40%)
        SERV_idx = 1  # SERV - Services confidence indicator (30%)
        CONS_idx = 2  # CONS - Consumer confidence indicator (20%)
        RETA_idx = 3  # RETA - Retail trade confidence indicator (5%)
        BUIL_idx = 4  # BUIL - Construction confidence indicator (5%)
        ESI_idx = 5  # ESI - Economic sentiment indicator, composite.

        if date is None:
            now = datetime.datetime.now()
            month_ago = now + dateutil.relativedelta.relativedelta(months=-1)
            start_date = '{}-{}'.format(month_ago.year, month_ago.month)
            end_date = '{}-{}'.format(now.year, now.month)
        else:
            start_date = date
            end_date = date

        esi_tables = self._fetch_esi_tables()

        latest_values = {}
        for ec in self.ESI_ENTITIES.keys():
            try:
                latest_values[ec] = (
                    esi_tables[ec][start_date:end_date]
                    .tail(1)
                    .values.tolist()[0]
                )
            except IndexError:
                sys.exit('Date given is out of range')

        industrial_ranking = {}
        services_ranking = {}
        consumer_ranking = {}
        retail_ranking = {}
        construction_ranking = {}
        esi_ranking = {}
        for ec, label in self.ESI_ENTITIES.items():
            industrial_ranking[label] = latest_values[ec][INDU_idx]
            services_ranking[label] = latest_values[ec][SERV_idx]
            consumer_ranking[label] = latest_values[ec][CONS_idx]
            retail_ranking[label] = latest_values[ec][RETA_idx]
            construction_ranking[label] = latest_values[ec][BUIL_idx]
            esi_ranking[label] = latest_values[ec][ESI_idx]
        rankings = {
            'industrial_confidence': industrial_ranking,
            'services_confidence': services_ranking,
            'consumer_confidence': consumer_ranking,
            'retail_confidence': retail_ranking,
            'construction_confidence': construction_ranking,
            'esi': esi_ranking
        }
        for ranking, values in rankings.items():
            rankings[ranking] = sorted(
                values.items(), key=lambda x: x[1], reverse=True
            )

        return rankings

    def get_historical_values(self, esi_component, months=12):
        """Returns a data structure with historical values for a given ESI
        component. For example:

        {'countries': {'at': [100.7,
                              98.8,
                              # ...
                              87.0,
                              89.4],
                       'be': [94.3,
                              93.9,
                              # ...
                              83.5,
                              88.8],
                       'de': [98.2,
                              98.6,
                              99.1,
                              # ...
                              94.3,
                              95.5],
                       'dk': [96.3,
                              100.9,
                              # ...
                              76.2,
                              77.7,
                              80.6],
                       'ea': [100.2,
                              100.7,
                              # ...
                              87.5,
                              91.1],
                       'el': [107.8,
                              108.1,
                              # ...
                              90.7,
                              89.5],
                       'pt': [107.1,
                              108.2,
                              # ...
                              85.9,
                              87.1],
                       'se': [95.8,
                              95.0,
                              # ...
                              88.9,
                              94.3],
                       'uk': [88.9,
                              89.7,
                              # ...
                              75.1,
                              83.0]},
         'dates': [Period('2019-10', 'M'),
                   Period('2019-11', 'M'),
                   # ...
                   Period('2020-08', 'M'),
                   Period('2020-09', 'M')]}
        """
        if esi_component not in self.ESI_COMPONENTS:
            esi_component = '.ESI'
        esi_tables = self._fetch_esi_tables()

        values = {'countries': {}, 'dates': []}
        for ec in self.ENTITY_CODES:
            col = '{}{}'.format(ec.upper(), esi_component)
            values['countries'][ec] = esi_tables[ec][col].tail(months).tolist()
        values['dates'] = (
            esi_tables[self.ENTITY_CODES[0]].tail(months).index.tolist()
        )

        return values


def display_latest_rankings(date=None, json_output=False, data_dir=None,
                            esi_filename=None, esi_sheet_name=None):
    """Display ESI rankings in the console or output as JSON."""
    BOLD = '\033[1m'
    ENDC = '\033[0m'
    GREEN = '\033[92m'
    RED = '\033[91m'
    indicators = [
        'esi',
        'industrial_confidence',
        'services_confidence',
        'consumer_confidence',
        'retail_confidence',
        'construction_confidence'
    ]
    esi = ESIDataWrapper()
    if data_dir:
        esi.data_dir = data_dir
    if esi_filename:
        esi.esi_filename = esi_filename
    if esi_sheet_name:
        esi.esi_sheet_name = esi_sheet_name
    num_entries = len(esi.ESI_ENTITIES)
    rankings = esi.get_latest_rankings(date=date)

    if json_output:
        print(json.dumps(rankings))
    else:
        table_data = [
            # Headers.
            [
                BOLD + 'ESI' + ENDC,
                'Industrial Confidence (40%)',
                'Services Confidence (30%)',
                'Consumer Confidence (20%)',
                'Retail Trade Confidence (5%)',
                'Construction Confidence (5%)'
            ]
        ]
        for i in range(num_entries):
            if i == 0:
                tmpl = GREEN + '{} ({})' + ENDC
            elif i == 15:
                tmpl = RED + '{} ({})' + ENDC
            else:
                tmpl = '{} ({})'

            row = [
                tmpl.format(
                    rankings[indicator][i][0], rankings[indicator][i][1]
                )
                for indicator in indicators
            ]
            table_data.append(row)

        rankings_table = SingleTable(table_data)
        rankings_table.inner_heading_row_border = False
        rankings_table.inner_heading_row_border = True
        if date:
            rankings_table.title = 'Rankings for {}'.format(date)

        print(rankings_table.table)


def historical_esi_values_chart(esi_component, title, filename=None,
                                months=12, data_dir=None, esi_filename=None,
                                esi_sheet_name=None):
    """Generates an SVG chart with historical values for an ESI component."""
    disable_xml_declaration = True
    if filename is not None:
        disable_xml_declaration = False
    title = 'ESI - {} (past {} months)'.format(title, months)
    esi = ESIDataWrapper()
    if data_dir:
        esi.data_dir = data_dir
    if esi_filename:
        esi.esi_filename = esi_filename
    if esi_sheet_name:
        esi.esi_sheet_name = esi_sheet_name
    values = esi.get_historical_values(esi_component, months)

    chart = pygal.Line(
        dots_size=1,
        show_y_guides=False,
        x_label_rotation=90,
        disable_xml_declaration=disable_xml_declaration
    )
    chart.title = title
    chart.x_labels = map(lambda d: d.strftime('%Y-%m'), values['dates'])
    for country, val in values['countries'].items():
        chart.add(country, val)

    if filename:
        chart.render_to_file(filename)
    else:
        return chart.render()


# Convenience functions.


def industrial_esi_chart(filename=None, months=12, data_dir=None,
                         esi_filename=None, esi_sheet_name=None):
    """Render an SVG chart with ESI Industrial Confidence data."""
    return historical_esi_values_chart(
        '.INDU',
        'Industrial Confidence',
        filename=filename,
        months=months,
        data_dir=data_dir,
        esi_filename=esi_filename,
        esi_sheet_name=esi_sheet_name
    )


def services_esi_chart(filename=None, months=12, data_dir=None,
                       esi_filename=None, esi_sheet_name=None):
    """Render an SVG chart with ESI Services Confidence data."""
    return historical_esi_values_chart(
        '.SERV',
        'Services Confidence',
        filename=filename,
        months=months,
        data_dir=data_dir,
        esi_filename=esi_filename,
        esi_sheet_name=esi_sheet_name
    )


def consumer_esi_chart(filename=None, months=12, data_dir=None,
                       esi_filename=None, esi_sheet_name=None):
    """Render an SVG chart with ESI Consumer Confidence data."""
    return historical_esi_values_chart(
        '.CONS',
        'Consumer Confidence',
        filename=filename,
        months=months,
        data_dir=data_dir,
        esi_filename=esi_filename,
        esi_sheet_name=esi_sheet_name
    )


def retail_trade_esi_chart(filename=None, months=12, data_dir=None,
                           esi_filename=None, esi_sheet_name=None):
    """Render an SVG chart with ESI Retail Trade Confidence data."""
    return historical_esi_values_chart(
        '.RETA',
        'Retail Trade Confidence',
        filename=filename,
        months=months,
        data_dir=data_dir,
        esi_filename=esi_filename,
        esi_sheet_name=esi_sheet_name
    )


def construction_esi_chart(filename=None, months=12, data_dir=None,
                           esi_filename=None, esi_sheet_name=None):
    """Render an SVG chart with ESI Construction Confidence data."""
    return historical_esi_values_chart(
        '.BUIL',
        'Construction Confidence',
        filename=filename,
        months=months,
        data_dir=data_dir,
        esi_filename=esi_filename,
        esi_sheet_name=esi_sheet_name
    )


def esi_chart(filename=None, months=12, data_dir=None, esi_filename=None,
              esi_sheet_name=None):
    """Render an SVG chart with ESI data."""
    return historical_esi_values_chart(
        '.ESI',
        'ESI',
        filename=filename,
        months=months,
        data_dir=data_dir,
        esi_filename=esi_filename,
        esi_sheet_name=esi_sheet_name
    )


if __name__ == '__main__':
    fire.Fire(
        {
            'latest_rankings': display_latest_rankings,
            'industrial_chart': industrial_esi_chart,
            'services_chart': services_esi_chart,
            'consumer_chart': consumer_esi_chart,
            'retail_trade_chart': retail_trade_esi_chart,
            'construction_chart': construction_esi_chart,
            'esi_chart': esi_chart
        }
    )
