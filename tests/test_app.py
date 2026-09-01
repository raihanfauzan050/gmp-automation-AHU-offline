import unittest

from ahu_utils import ahu_sort_key, default_ahu_for_test, extract_ahu_number


class AhuNumberTest(unittest.TestCase):
    def test_airborne_and_airflow_default_to_ahu_33(self):
        self.assertEqual(default_ahu_for_test('airborne_particle'), '33')
        self.assertEqual(default_ahu_for_test('airflow_pattern'), '33')
        self.assertEqual(default_ahu_for_test('air_velocity'), 'unknown')
        self.assertEqual(
            extract_ahu_number(
                'unknown',
                '/tmp/airborne.pdf',
                default=default_ahu_for_test('airborne_particle'),
            ),
            '33',
        )

    def test_sorts_numeric_and_text_ahu_values(self):
        self.assertEqual(
            sorted(['unknown', '37', '5'], key=ahu_sort_key),
            ['5', '37', 'unknown'],
        )

    def test_normalizes_ocr_value(self):
        self.assertEqual(extract_ahu_number('공 조 기 - 37'), '37')
        self.assertEqual(extract_ahu_number('AHU-42'), '42')
        self.assertEqual(extract_ahu_number('공조기 번호 33'), '33')
        self.assertEqual(extract_ahu_number('AHU No. 34'), '34')

    def test_falls_back_to_filename(self):
        filename = '/tmp/uuid_AHU-37_air_change_rate.pdf'
        self.assertEqual(extract_ahu_number('unknown', filename), '37')

    def test_rejects_zero_and_falls_back_to_filename(self):
        filename = '/tmp/uuid_AHU-33_airborne_particle.pdf'
        self.assertEqual(extract_ahu_number('0', filename), '33')
        self.assertEqual(extract_ahu_number('AHU-0'), 'unknown')

    def test_filename_ahu_overrides_an_incorrect_ocr_number(self):
        filename = '/tmp/uuid_AHU-33_hepa_filter.pdf'
        self.assertEqual(extract_ahu_number('1', filename), '33')

    def test_airflow_uses_filename_before_default(self):
        self.assertEqual(
            extract_ahu_number('unknown', '/tmp/AHU-34_airflow.pdf', default='33'),
            '34',
        )
        self.assertEqual(
            extract_ahu_number('unknown', '/tmp/airflow.pdf', default='33'),
            '33',
        )


if __name__ == '__main__':
    unittest.main()
