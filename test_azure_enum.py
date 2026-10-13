import unittest
from unittest.mock import patch

import azure_enum


class GraphPaginationTests(unittest.TestCase):
    def test_follows_graph_continuation_and_combines_results(self):
        pages = [
            {
                "value": [{"id": "first"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next",
            },
            {"value": [{"id": "second"}]},
        ]

        with patch("azure_enum.graph_get", side_effect=pages) as graph_get:
            result = azure_enum.list_users("mock-token")

        self.assertEqual(result, [{"id": "first"}, {"id": "second"}])
        self.assertEqual(graph_get.call_count, 2)
        self.assertEqual(graph_get.call_args_list[1].args[0], pages[0]["@odata.nextLink"])

    def test_stops_and_fails_instead_of_returning_truncated_inventory(self):
        continued_page = {
            "value": [{"id": "record"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next",
        }

        with patch("azure_enum.MAX_GRAPH_PAGES", 2), patch(
            "azure_enum.graph_get", return_value=continued_page
        ) as graph_get:
            with self.assertRaisesRegex(RuntimeError, "2-page safety limit"):
                azure_enum.graph_get_all("users?$top=999", "mock-token")

        self.assertEqual(graph_get.call_count, 2)

    def test_rejects_untrusted_continuation_url_before_request(self):
        with patch("azure_enum.urllib.request.urlopen") as urlopen:
            with self.assertRaisesRegex(ValueError, "unsafe continuation URL"):
                azure_enum.graph_get(
                    "https://attacker.example/v1.0/users?$skiptoken=x",
                    "mock-token",
                )
        urlopen.assert_not_called()

    def test_does_not_treat_graph_errors_as_empty_inventory(self):
        with patch("azure_enum.graph_get", return_value={"error": "mock failure"}):
            with self.assertRaisesRegex(RuntimeError, "request failed"):
                azure_enum.graph_get_all("users?$top=999", "mock-token")


if __name__ == "__main__":
    unittest.main()
