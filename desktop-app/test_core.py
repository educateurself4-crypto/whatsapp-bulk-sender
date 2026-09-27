"""Run:  python -m unittest test_core -v   (no network, no GUI needed)"""
import csv
import json
import os
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, HTTPServer

from openpyxl import Workbook

import core
from core import N8nClient, N8nError, build_contacts, normalize_phone, parse_template, read_table, run_campaign


class PhoneTests(unittest.TestCase):
    def test_valid_formats(self):
        for raw in ["9876543210", "+91 98765 43210", "919876543210", "09876543210", "0091-98765-43210", "98765-43210"]:
            self.assertEqual(normalize_phone(raw)[0], "919876543210", raw)

    def test_foreign_number(self):
        self.assertEqual(normalize_phone("+1 415 555 0100")[0], "14155550100")
        self.assertEqual(normalize_phone("14155550100")[0], "14155550100")

    def test_invalid(self):
        for raw in ["", "abc", "12345", "5876543210", "+91 12345"]:
            n, err = normalize_phone(raw)
            self.assertIsNone(n, raw)
            self.assertTrue(err)


class TemplateTests(unittest.TestCase):
    def test_supported(self):
        t = parse_template({"name": "sale", "language": "en", "category": "MARKETING", "components": [
            {"type": "BODY", "text": "Hi {{1}}, {{2}} off till {{3}}"},
            {"type": "FOOTER", "text": "Reply STOP"},
            {"type": "BUTTONS", "buttons": [{"type": "QUICK_REPLY", "text": "Stop"}]}]})
        self.assertTrue(t.supported)
        self.assertEqual(t.var_count, 3)

    def test_unsupported(self):
        img = parse_template({"name": "a", "language": "en", "category": "MARKETING", "components": [
            {"type": "HEADER", "format": "IMAGE"}, {"type": "BODY", "text": "x"}]})
        self.assertFalse(img.supported)
        named = parse_template({"name": "b", "language": "en", "category": "MARKETING", "components": [
            {"type": "BODY", "text": "Hi {{first_name}}"}]})
        self.assertFalse(named.supported)
        urlvar = parse_template({"name": "c", "language": "en", "category": "MARKETING", "components": [
            {"type": "BODY", "text": "x"}, {"type": "BUTTONS", "buttons": [{"type": "URL", "url": "https://x.com/{{1}}"}]}]})
        self.assertFalse(urlvar.supported)
        static_url = parse_template({"name": "d", "language": "en", "category": "MARKETING", "components": [
            {"type": "BODY", "text": "x"}, {"type": "BUTTONS", "buttons": [{"type": "URL", "url": "https://x.com/shop"}]}]})
        self.assertTrue(static_url.supported)


class FileAndContactTests(unittest.TestCase):
    def make_xlsx(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Mobile", "Offer", "Offer"])
        ws.append(["Rahul", 9876543210, "Diwali", 20.0])          # numeric phone + float
        ws.append(["Priya", "+91 98765 43211", "Diwali\nSale", 15])  # newline in value
        ws.append(["Dup", "9876543210", "Diwali", 20])              # duplicate
        ws.append(["Bad", "12345", "Diwali", 20])                   # invalid phone
        ws.append(["Empty", "9876543212", "", 20])                  # empty variable
        ws.append(["Optout", "9876543213", "Diwali", 20])
        ws.append([None, None, None, None])                         # blank row
        f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        wb.save(f.name)
        return f.name

    def test_pipeline(self):
        headers, rows = read_table(self.make_xlsx())
        self.assertEqual(headers, ["Name", "Mobile", "Offer", "Offer (2)"])
        self.assertEqual(len(rows), 6)  # blank row dropped
        rep = build_contacts(rows, 1, [0, 2, 3], "91", opted_out={"919876543213"})
        self.assertEqual([c["phone"] for c in rep.valid], ["919876543210", "919876543211"])
        self.assertEqual(rep.valid[0]["params"], ["Rahul", "Diwali", "20"])
        self.assertEqual(rep.valid[1]["params"], ["Priya", "Diwali Sale", "15"])
        self.assertEqual(rep.duplicates, 1)
        self.assertEqual(rep.opted_out, 1)
        self.assertEqual(len(rep.invalid), 2)
        self.assertEqual(rep.invalid[0][0], 5)  # Excel row number of the bad phone

    def test_csv(self):
        f = tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, newline="", encoding="utf-8-sig")
        csv.writer(f).writerows([["Name", "Phone"], ["A", "9876543210"]])
        f.close()
        h, r = read_table(f.name)
        self.assertEqual((h, r), (["Name", "Phone"], [["A", "9876543210"]]))


class MockN8n(BaseHTTPRequestHandler):
    """Pretends to be n8n: records chunks, 'processes' them instantly."""
    sent = []
    fail_processing = False

    def log_message(self, *a):
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])) or b"{}")
        if self.headers.get("x-api-key") != "secret":
            self.send_response(403); self.end_headers(); return
        path = self.path.rsplit("/", 1)[-1]
        out = {}
        if path == "wa-send-campaign":
            if not MockN8n.fail_processing:
                MockN8n.sent.extend(body["contacts"])
        elif path == "wa-campaign-status":
            n = len(MockN8n.sent)
            out = {"processed": n, "accepted": n, "send_failed": 0, "delivered": 0, "read": 0, "delivery_failed": 0, "failures": []}
        elif path == "wa-optouts":
            out = {"phones": ["919876543213"]}
        elif path == "wa-list-templates":
            out = {"templates": [{"name": "x"}]}
        else:
            self.send_response(404); self.end_headers(); return
        data = json.dumps(out).encode()
        self.send_response(200)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


class RunnerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.srv = HTTPServer(("127.0.0.1", 0), MockN8n)
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.url = f"http://127.0.0.1:{cls.srv.server_port}"

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()

    def setUp(self):
        MockN8n.sent = []
        MockN8n.fail_processing = False
        self.tpl = core.Template("sale", "en", "MARKETING", "Hi {{1}}", 1, True)
        self.contacts = [{"phone": f"9198765{i:05d}", "params": ["A"]} for i in range(120)]

    def test_bad_key(self):
        with self.assertRaises(N8nError):
            N8nClient(self.url, "wrong").optouts()

    def test_chunks_sent_in_order_and_all_delivered_to_n8n(self):
        log = []
        n = run_campaign(N8nClient(self.url, "secret"), "c1", self.tpl, self.contacts,
                         threading.Event(), lambda m, d, t: log.append(m), chunk_size=50, poll_seconds=0.05)
        self.assertEqual(n, 120)
        self.assertEqual(MockN8n.sent, self.contacts)

    def test_cancel(self):
        cancel = threading.Event()
        cancel.set()
        n = run_campaign(N8nClient(self.url, "secret"), "c2", self.tpl, self.contacts, cancel, lambda *a: None, chunk_size=50)
        self.assertEqual(n, 0)

    def test_timeout_does_not_continue(self):
        MockN8n.fail_processing = True
        orig = core.time.time
        t = [1000.0]
        core.time.time = lambda: t.__setitem__(0, t[0] + 500) or t[0]  # fast-forward the clock
        try:
            with self.assertRaises(N8nError):
                run_campaign(N8nClient(self.url, "secret"), "c3", self.tpl, self.contacts,
                             threading.Event(), lambda *a: None, chunk_size=50, poll_seconds=0.01)
        finally:
            core.time.time = orig
        self.assertEqual(MockN8n.sent, [])


if __name__ == "__main__":
    unittest.main()
