import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import load_workbook

from demo.generate_inputs import OIL, PRODUCT, create_inputs
from desktop_app import SITE_CONFIGURATIONS, load_generator_module


class DemoIntegrationTests(unittest.TestCase):
    def test_both_sites_create_scenario_and_preserve_output_on_error(self):
        for site, sheet_count in (("Евдаково", 49), ("ЭПИ", 39)):
            with self.subTest(site=site), TemporaryDirectory() as temporary:
                root = Path(temporary)
                source = root / "sources"
                as_of = pd.Timestamp.today().date()
                next_month = date(
                    as_of.year + (as_of.month == 12), as_of.month % 12 + 1, 1
                )
                create_inputs(source, site, as_of)
                config = SITE_CONFIGURATIONS[site]
                pipeline = load_generator_module(config.generator_path)
                output = root / config.output_filename
                pipeline.create_scenario(source, output)

                with output.open("rb") as stream:
                    workbook = load_workbook(stream, read_only=True, data_only=True)
                    try:
                        self.assertEqual(
                            workbook.sheetnames, list(pipeline.SHEET_ORDER)
                        )
                        self.assertEqual(len(workbook.sheetnames), sheet_count)
                        demand = list(workbook["DemandRecord"].values)
                        order = dict(zip(*demand, strict=True))
                        self.assertEqual(order["material"], PRODUCT)
                        self.assertEqual(order["mass"], 42)
                        self.assertEqual(order["beginDate"].date(), next_month)
                        stock = list(workbook["InitialStock"].values)
                        row = dict(zip(stock[0], stock[1], strict=True))
                        self.assertEqual(row["material"], OIL)
                        self.assertEqual(row["mass"], 2.5)
                        if site == "Евдаково":
                            header, values = workbook["packagingChangeover"].values
                            row = dict(zip(header, values, strict=True))
                            self.assertEqual(row["duration"], 2)
                            self.assertEqual(row["cleaningMaterial"], OIL)
                            self.assertEqual(row["cleaningMaterialQuantity"], 0.5)
                    finally:
                        workbook.close()

                original = output.read_bytes()
                marker, _ = pipeline.SOURCES["materials"]
                missing = next(source.glob(f"{Path(marker).stem}*.xlsx"))
                missing.unlink()
                with self.assertRaisesRegex(
                    pipeline.ScenarioGenerationError, "Material"
                ):
                    pipeline.create_scenario(source, output)
                self.assertEqual(output.read_bytes(), original)
