import unittest

from deepseek_ocr.parsers import (
    extract_ahu,
    parse_airborne_particle,
    parse_air_change_rate,
    parse_air_velocity,
    parse_airflow_pattern,
    parse_hepa_filter,
)


AIRBORNE_ROOMS = [
    ('B', '2142', '무균 실험실', [(121, 7), (194, 0), (676, 9), (576, 9), (107, 1), (121, 4)]),
    ('D', '2165', '균주접종실', [(5420, 400), (570, 20), (660, 50), (4000, 620), (690, 110), (600, 10)]),
    ('C', '2147', '탈의실', [(30760, 1340), (17690, 550)]),
    ('D', '2169', '복도', [(11140, 1670), (4120, 650), (4710, 590)]),
    ('D', '2145', '탈의실', [(54260, 1170)]),
    ('B', '2142', 'PASS BOX(QHA-745)', [(539, 4), (317, 0)]),
    ('B', '2142', 'PASS BOX(QHA-744)', [(90, 1), (114, 1)]),
    ('D', '2169', 'PASS BOX(QHA-743)', [(25830, 2240), (30810, 2960)]),
    ('D', '2169', 'PASS BOX(QHA-742)', [(60, 0), (70, 0)]),
    ('A', '2142', '무균시험실 BSC', [(3, 0), (1, 0)]),
    ('B', '2165', '균주접종실 BSC', [(4, 0), (0, 0)]),
]


def airborne_page(rows, include_header=False):
    header = ''
    if include_header:
        header = (
            '<tr><th>NO</th><th>청정 등급</th><th>실번호</th><th>실명</th>'
            '<th>측정번호</th><th>0.5 μm</th><th>5.0 μm</th></tr>'
        )
    body = ''.join(
        f'<tr><td>{no}</td><td>{grade}</td><td>{room_number}</td><td>{name}</td>'
        f'<td>{point}</td><td>{value_05}</td><td>{value_50}</td></tr>'
        for no, grade, room_number, name, point, value_05, value_50 in rows
    )
    return f'<table>{header}{body}</table>'


class AirborneParticleParserTest(unittest.TestCase):
    def test_does_not_treat_particle_size_as_ahu_zero(self):
        self.assertEqual(extract_ahu('해당 공조기 0.5 μm'), 'unknown')

    def test_reuses_header_for_continuation_pages_and_keeps_all_30_points(self):
        rows = []
        no = 1
        for grade, room_number, name, measurements in AIRBORNE_ROOMS:
            for point, (value_05, value_50) in enumerate(measurements, start=1):
                rows.append((no, grade, room_number, name, point, value_05, value_50))
                no += 1

        pages = [
            (
                '<table><tr><td>해당 공조기</td><td>공조기-33</td></tr></table>'
                + airborne_page(rows[:12], include_header=True)
            ),
            airborne_page(rows[12:13]),
            airborne_page(rows[13:]),
        ]
        result = parse_airborne_particle(pages)

        self.assertEqual(result['ahu'], '33')
        self.assertEqual(len(result['rooms']), 11)
        self.assertEqual(sum(len(room['measurements']) for room in result['rooms']), 30)
        self.assertEqual(
            [room['room_name'] for room in result['rooms']],
            [room[2] for room in AIRBORNE_ROOMS],
        )
        self.assertEqual(result['rooms'][2]['no_start'], 13)
        self.assertEqual(result['rooms'][2]['no_end'], 14)
        self.assertEqual(
            result['rooms'][2]['measurements'],
            [
                {'point': 1, 'value_05': 30760, 'value_50': 1340},
                {'point': 2, 'value_05': 17690, 'value_50': 550},
            ],
        )


class AhuMetadataParserTest(unittest.TestCase):
    def test_reads_horizontal_and_vertical_ahu_metadata(self):
        horizontal = '<table><tr><td>해당 공조기</td><td>공조기-33</td></tr></table>'
        vertical = """
        <table>
          <tr><td>측정기준</td><td>해당<br>공조기</td><td>측정일자</td><td>측정결과</td></tr>
          <tr><td>0.01%↓</td><td>공조기-33</td><td>2025.08.03</td><td>적합</td></tr>
        </table>
        """
        self.assertEqual(extract_ahu(horizontal), '33')
        self.assertEqual(extract_ahu(vertical), '33')

    def test_reads_ahu_number_label_variants(self):
        numbered = '<table><tr><td>공조기 번호</td><td>33호기</td></tr></table>'
        english = '<table><tr><td>AHU No.</td><td>34</td></tr></table>'
        self.assertEqual(extract_ahu(numbered), '33')
        self.assertEqual(extract_ahu(english), '34')

    def test_does_not_use_measurement_number_when_ahu_is_blank(self):
        page = """
        <table>
          <tr><td>해당 공조기</td><td>측정일자</td></tr>
          <tr><td></td><td>2025.08.03</td></tr>
        </table>
        <table><tr><td>NO</td><td>측정번호</td></tr><tr><td>1</td><td>1</td></tr></table>
        """
        self.assertEqual(extract_ahu(page), 'unknown')

    def test_air_velocity_reads_horizontal_ahu_metadata(self):
        page = """
        <table><tr><td>해당 공조기</td><td>공조기-34</td></tr></table>
        <table>
          <tr><th>NO</th><th>청정등급</th><th>실번호</th><th>실명</th><th>측정번호</th><th>측정값</th></tr>
          <tr><td>1</td><td>A</td><td>2142</td><td>무균시험실 BSC</td><td>1</td><td>0.45</td></tr>
        </table>
        """
        self.assertEqual(parse_air_velocity([page])['ahu'], '34')

    def test_hepa_reads_ahu_value_below_header(self):
        page = """
        <table>
          <tr><td>측정기준</td><td>해당<br>공조기</td><td>측정일자</td><td>측정결과</td></tr>
          <tr><td>0.01%↓</td><td>공조기-33</td><td>2025.08.03</td><td>적합</td></tr>
        </table>
        <table>
          <tr><th>NO</th><th>실번호</th><th>실명</th><th>측정번호</th><th>측정값</th></tr>
          <tr><td>1</td><td>2142</td><td>무균시험실 BSC</td><td>1</td><td>0.003%</td></tr>
        </table>
        """
        result = parse_hepa_filter([page])
        self.assertEqual(result['ahu'], '33')
        self.assertEqual(len(result['items']), 1)


class AirChangeRateParserTest(unittest.TestCase):
    def test_parses_spaced_metadata_and_groups_measurement_rows(self):
        page = """
        <table>
          <tr><td>측 정 일 자</td><td>2025. 03. 27</td><td>해 당 공 조 기</td><td>공 조 기 − 37</td></tr>
          <tr><td>측 정 결 과</td><td>적 합</td></tr>
        </table>
        <table>
          <tr>
            <th rowspan="2">NO.</th><th rowspan="2">청정 등급</th><th rowspan="2">실번호</th>
            <th rowspan="2">실명</th><th rowspan="2">체적 (m3)</th><th rowspan="2">측정 번호</th>
            <th colspan="2">측정값</th>
          </tr>
          <tr><th>풍량 (m3/hr)</th><th>환기횟수 (회/hr)</th></tr>
          <tr>
            <td rowspan="3">1</td><td rowspan="3">D</td><td rowspan="3">3101</td>
            <td rowspan="3">칭량실</td><td rowspan="3">50.0</td><td>1</td><td>480.2</td><td rowspan="3">20</td>
          </tr>
          <tr><td>2</td><td>500.3</td></tr>
          <tr><td>합계</td><td>980.5</td></tr>
        </table>
        """

        result = parse_air_change_rate([page])

        self.assertEqual(result['ahu'], '37')
        self.assertEqual(result['date'], '2025.03.27')
        self.assertEqual(result['result'], '적합')
        self.assertEqual(len(result['rooms']), 1)
        room = result['rooms'][0]
        self.assertEqual(room['room_number'], '3101')
        self.assertEqual(room['total_air_flow'], 980.5)
        self.assertEqual(room['ach'], 20)
        self.assertEqual(
            room['air_flow_measurements'],
            [{'point': 1, 'air_flow': 480.2}, {'point': 2, 'air_flow': 500.3}],
        )


class AirflowPatternParserTest(unittest.TestCase):
    def test_extracts_only_field_values_from_each_page(self):
        names = [
            '무균시험실 BSC',
            '균주접종실 BSC',
            '미생물 시험실 C/B(852)',
            '미생물 시험실 C/B(853)',
            'PASS BOX (QHA-745)',
            'PASS BOX (QHA-744)',
            'PASS BOX (QHA-743)',
            'PASS BOX (QHA-742)',
        ]
        criteria = (
            '1. 육안상 단일방향류가 형성되어야 함.\n'
            '2. 측정대상 크린장비 내부에 난류가 형성되는 구역이 없어야 함.'
        )
        pages = []
        for name in names:
            pages.append(f"""
            <table>
              <tr>
                <td>측정대상</td><td>{name}</td><td>측정일자</td><td>2025.08.02</td>
                <td>결재</td><td>측정자</td><td>확인자</td>
              </tr>
            </table>
            <table><tr><td>측정사진</td><td>image</td></tr></table>
            <table>
              <tr><td>측정기준</td><td>1. 육안상 단일방향류가 형성되어야 함.<br>
              2. 측정대상 크린장비 내부에 난류가 형성되는 구역이 없어야 함.</td></tr>
            </table>
            <table>
              <tr><td rowspan="2">측정결과</td><td>동영상 첨부</td><td>판정결과</td></tr>
              <tr><td>첨부</td><td>적합</td></tr>
            </table>
            """)

        result = parse_airflow_pattern(pages)

        self.assertEqual(result['ahu'], 'unknown')
        self.assertEqual(result['date'], '2025.08.02')
        self.assertEqual(len(result['items']), 8)
        self.assertEqual([item['name'] for item in result['items']], names)
        for item in result['items']:
            self.assertEqual(item['date'], '2025.08.02')
            self.assertEqual(item['criteria'], criteria)
            self.assertEqual(item['video_attached'], '첨부')
            self.assertEqual(item['judgment'], '적합')


if __name__ == '__main__':
    unittest.main()
