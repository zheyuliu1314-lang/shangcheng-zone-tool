import io
import unittest

import openpyxl

from app import app, normalize_zones, classify_points_data, build_export, classify_gaode_error


class AppRegressionTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()

    def make_xlsx(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["公司名称", "公司地址", "备注"])
        sheet.append(["测试公司", "杭州市上城区清江路1号", "保留"])
        stream = io.BytesIO()
        workbook.save(stream)
        stream.seek(0)
        return stream

    def test_polygon_validation_rejects_self_intersection(self):
        zones, errors = normalize_zones([{
            "name": "自交片区", "color": "#123456",
            "polygon": [[120.1, 30.1], [120.2, 30.2], [120.1, 30.2], [120.2, 30.1]],
        }])
        self.assertEqual(zones, [])
        self.assertTrue(errors)

    def test_classification_marks_low_confidence_for_review(self):
        result = classify_points_data([{
            "row_index": 0, "lng": 120.15, "lat": 30.15,
            "confidence": 0.65, "accuracy_warning": False,
        }], [{
            "name": "片区A", "color": "#123456",
            "polygon": [[120.1, 30.1], [120.2, 30.1], [120.2, 30.2], [120.1, 30.2]],
        }])
        self.assertEqual(result[0]["zone"], "片区A")
        self.assertEqual(result[0]["review_status"], "待复核")

    def test_classification_distinguishes_outside_zone(self):
        result = classify_points_data([{
            "row_index": 0, "lng": 120.25, "lat": 30.25,
            "confidence": 0.95, "accuracy_warning": False,
        }], [{
            "name": "片区A", "color": "#123456",
            "polygon": [[120.1, 30.1], [120.2, 30.1], [120.2, 30.2], [120.1, 30.2]],
        }])
        self.assertEqual(result[0]["zone"], "片区外")
        self.assertEqual(result[0]["review_status"], "片区外")

    def test_gaode_error_mapping_is_specific(self):
        self.assertEqual(classify_gaode_error({"status": "0", "infocode": "10001", "info": "INVALID_USER_KEY"})["message"], "Key无效")
        self.assertEqual(classify_gaode_error({"status": "0", "infocode": "10003", "info": "DAILY_QUERY_OVER_LIMIT"})["message"], "调用额度不足")
        self.assertEqual(classify_gaode_error(exc=TimeoutError("timeout"))["message"], "网络错误")

    def test_upload_creates_job_and_preserves_headers(self):
        response = self.client.post(
            "/api/upload",
            data={"file": (self.make_xlsx(), "sample.xlsx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload["job_id"])
        self.assertEqual(payload["headers"], ["公司名称", "公司地址", "备注"])
        self.assertEqual(payload["total"], 1)

    def test_export_contains_review_columns(self):
        stream = build_export([{
            "公司名称": "测试公司", "公司地址": "杭州市上城区清江路1号",
            "zone": "片区A", "lng": 120.15, "lat": 30.15,
            "geocode_status": "成功", "confidence": 0.95,
            "level": "兴趣点", "source": "geo", "manual_reviewed": False,
        }], ["公司名称", "公司地址"], {"片区A": "#123456"})
        workbook = openpyxl.load_workbook(stream, data_only=True)
        headers = [cell.value for cell in workbook.active[1]]
        self.assertIn("所属片区", headers)
        self.assertIn("复核备注", headers)


if __name__ == "__main__":
    unittest.main()
