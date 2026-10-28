# -*- coding: utf-8 -*-
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(__file__))
import core


class CoreTests(unittest.TestCase):
    def test_classification_fire(self):
        self.assertEqual(core.wall_classification("Yes", "No", ""), "FIRE")

    def test_classification_smoke(self):
        self.assertEqual(core.wall_classification("No", "Yes", ""), "SMOKE")

    def test_classification_both(self):
        self.assertEqual(core.wall_classification("1", "1", ""), "BOTH")

    def test_missing_rating_is_unknown(self):
        self.assertEqual(core.wall_classification("", "", ""), "UNKNOWN")

    def test_nonrated_is_not_unknown(self):
        self.assertEqual(core.wall_classification("No", "No", ""), "NOT_RATED")

    def test_box_overlap_true(self):
        self.assertTrue(core.boxes_overlap(([0, 0, 0], [2, 2, 2]), ([1, 1, 1], [3, 3, 3])))

    def test_box_overlap_false(self):
        self.assertFalse(core.boxes_overlap(([0, 0, 0], [1, 1, 1]), ([2, 0, 0], [3, 1, 1])))

    def test_segment_through_box(self):
        hit = core.segment_box_intersection((-1, 0, 0), (1, 0, 0), (-0.5, -1, -1), (0.5, 1, 1))
        self.assertIsNotNone(hit)
        self.assertAlmostEqual(hit[0], 0.25)

    def test_near_miss(self):
        hit = core.segment_box_intersection((-1, 2, 0), (1, 2, 0), (-0.5, -1, -1), (0.5, 1, 1))
        self.assertIsNone(hit)

    def test_tangential_contact_is_not_volume_intersection(self):
        hit = core.segment_box_intersection((-1, -1, 0), (1, -1, 0), (-0.5, -1, -1), (0.5, 1, 1), tolerance=0.0)
        self.assertIsNotNone(hit)
        self.assertEqual(hit[0], 0.25)

    def test_finding_id_is_stable(self):
        a = core.make_finding_id("H", "L", "W", "D", "G")
        b = core.make_finding_id("H", "L", "W", "D", "G")
        self.assertEqual(a, b)

    def test_finding_id_changes_with_geometry(self):
        a = core.make_finding_id("H", "L", "W", "D", "G1")
        b = core.make_finding_id("H", "L", "W", "D", "G2")
        self.assertNotEqual(a, b)

    def test_unit_conversion_round_trip(self):
        self.assertAlmostEqual(core.ft_to_mm(core.mm_to_ft(1000.0)), 1000.0, places=5)

    def test_csv_export(self):
        handle, path = tempfile.mkstemp(suffix=".csv")
        os.close(handle)
        try:
            count = core.write_csv(path, [{"finding_id": "TEST", "status": "CONFIRMED_GEOMETRY"}])
            self.assertEqual(count, 1)
            with open(path, "rb") as stream:
                self.assertIn(b"finding_id", stream.read())
        finally:
            os.remove(path)

    def test_parameter_aliases_separated_by_multiple_delimiters(self):
        aliases = core.split_parameter_aliases("NWB_Material, Fire_Rating; Smoke_Rating\nLegacy_Rating")
        self.assertEqual(aliases, ["NWB_Material", "Fire_Rating", "Smoke_Rating", "Legacy_Rating"])

    def test_parameter_aliases_remove_duplicates_case_insensitively(self):
        aliases = core.split_parameter_aliases("NWB_Material, nwb_material; NWB_Material")
        self.assertEqual(aliases, ["NWB_Material"])

    def test_wall_material_empty_is_unknown(self):
        self.assertEqual(core.classify_wall_material(""), "UNKNOWN")

    def test_wall_material_explicit_fire(self):
        self.assertEqual(core.classify_wall_material("Fire Rated"), "FIRE")

    def test_wall_material_explicit_smoke(self):
        self.assertEqual(core.classify_wall_material("Smoke Rated"), "SMOKE")

    def test_wall_material_explicit_both(self):
        self.assertEqual(core.classify_wall_material("Fire Rated / Smoke Rated"), "BOTH")

    def test_wall_material_explicit_nonrated(self):
        self.assertEqual(core.classify_wall_material("Non Rated"), "NOT_RATED")


if __name__ == "__main__":
    unittest.main()
