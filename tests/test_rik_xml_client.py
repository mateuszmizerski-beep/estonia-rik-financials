from __future__ import annotations

from pathlib import Path
import unittest

from rik_xml_client import RikAuthenticationError, RikXmlClient


FIXTURES = Path(__file__).parent / "fixtures"


class FakeResponse:
    def __init__(self, content: bytes, status_code: int = 200) -> None:
        self.content = content
        self.status_code = status_code


class FakeSession:
    def __init__(self, responses: list[bytes]) -> None:
        self.responses = list(responses)
        self.headers: dict[str, str] = {}

    def post(self, *args, **kwargs) -> FakeResponse:
        return FakeResponse(self.responses.pop(0))


def fixture(name: str) -> bytes:
    return (FIXTURES / name).read_bytes()


class RikXmlClientTests(unittest.TestCase):
    def client(self, *fixtures: str) -> RikXmlClient:
        session = FakeSession([fixture(name) for name in fixtures])
        return RikXmlClient("fixture-user", "fixture-password", session=session)

    def test_company_summary(self) -> None:
        company = self.client("company_summary.xml").get_company_summary("70000310")
        self.assertEqual(company.registry_code, "70000310")
        self.assertEqual(company.name, "Registrite ja Infosüsteemide Keskus")
        self.assertIn("Tallinn", company.address or "")

    def test_report_list(self) -> None:
        reports = self.client("annual_report_list.xml").list_annual_reports("70000310")
        self.assertEqual({item.report_type for item in reports}, {"14", "35"})
        self.assertTrue(all(item.fiscal_year == 2024 for item in reports))
        self.assertEqual(reports[0].period_end.isoformat(), "2024-12-31")

    def test_statement_rows_and_comparative_column(self) -> None:
        lines = self.client("income_statement.xml").get_annual_report_lines(
            "70000310", "35", 2024
        )
        revenue = [line for line in lines if line.row_number == "10"]
        self.assertEqual([line.column_code for line in revenue], ["A1", "A2"])
        self.assertEqual(str(revenue[0].value), "5534707")
        self.assertIsNotNone(revenue[0].retrieved_at.tzinfo)

    def test_document_list_exposes_pdf(self) -> None:
        documents = self.client("document_list.xml").list_company_documents("70000310")
        self.assertEqual(len(documents), 1)
        self.assertEqual(documents[0].document_type, "A")
        self.assertEqual(documents[0].extension, ".pdf")

    def test_http_500_soap_authentication_fault_is_not_reported_as_generic_500(self) -> None:
        fault = b"""<?xml version="1.0" encoding="UTF-8"?>
        <SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/">
          <SOAP-ENV:Body>
            <SOAP-ENV:Fault>
              <faultcode>SOAP-ENV:Server</faultcode>
              <faultstring>Incorrect user name or password.</faultstring>
            </SOAP-ENV:Fault>
          </SOAP-ENV:Body>
        </SOAP-ENV:Envelope>"""

        class FaultSession(FakeSession):
            def post(self, *args, **kwargs) -> FakeResponse:
                return FakeResponse(fault, status_code=500)

        client = RikXmlClient(
            "fixture-user", "fixture-password", session=FaultSession([])
        )
        with self.assertRaisesRegex(
            RikAuthenticationError, "RIK rejected the configured credentials"
        ):
            client.get_company_summary("70000310")


if __name__ == "__main__":
    unittest.main()
